// enhance-pptx.mjs — lossless post-processor for PptxGenJS decks.
//
// PptxGenJS builds the deck; this layer opens the .pptx zip and surgically injects
// raw DrawingML (gradients, patterns, glow, soft edge, shadow, reflection, gradient
// text) into shapes tagged with an `objectName`. Every other part of the package is
// repackaged byte-for-byte, so charts / media / embedded workbooks are untouched
// (unlike a python-pptx round-trip, which drops parts it doesn't model).
//
// Usage:  const outBuf = await enhancePptx(pptxBuffer, enhancements)
//
// enhancements: [ {
//    objectName: "enh:hero",                       // matches addShape/addText({objectName})
//    fill: { type:"gradient", stops:[{pos:0,color:"008C64"},{pos:1,color:"222A30"}], angle:45 }
//        | { type:"pattern", preset:"wave", fore:"008C64", back:"FFFFFF" },
//    glow:       { color:"00E0A0", size:10, alpha:60 },
//    softEdge:   { size:5 },
//    shadow:     { color:"000000", blur:6, dist:4, angle:45, alpha:50 },
//    reflection: true,
//    textGradient: { stops:[{pos:0,color:"008C64"},{pos:1,color:"222A30"}], angle:90 }
// } ]
import JSZip from "jszip";

const PT = 12700;            // EMU per point
const ANG = 60000;          // 60000ths of a degree
const hex = (c) => String(c || "000000").replace(/^#/, "").toUpperCase();
const clampAlpha = (a) => Math.max(0, Math.min(100, a == null ? 100 : a));

// ---- DrawingML builders --------------------------------------------------
function gradFillXml(stops, angleDeg = 90) {
  const gs = stops.map((s) =>
    `<a:gs pos="${Math.round((s.pos ?? 0) * 100000)}"><a:srgbClr val="${hex(s.color)}"/></a:gs>`
  ).join("");
  return `<a:gradFill flip="none" rotWithShape="1"><a:gsLst>${gs}</a:gsLst>` +
         `<a:lin ang="${Math.round(angleDeg * ANG)}" scaled="1"/></a:gradFill>`;
}
function pattFillXml(preset, fore, back) {
  return `<a:pattFill prst="${preset || "pct50"}">` +
         `<a:fgClr><a:srgbClr val="${hex(fore)}"/></a:fgClr>` +
         `<a:bgClr><a:srgbClr val="${hex(back)}"/></a:bgClr></a:pattFill>`;
}
function shapeFillXml(fill) {
  if (!fill) return null;
  if (fill.type === "gradient") return gradFillXml(fill.stops || [], fill.angle ?? 90);
  if (fill.type === "pattern")  return pattFillXml(fill.preset, fill.fore, fill.back);
  if (fill.type === "solid")    return `<a:solidFill><a:srgbClr val="${hex(fill.color)}"/></a:solidFill>`;
  if (fill.type === "none")     return `<a:noFill/>`;
  return null;
}
function effectLstXml(e) {
  let out = "";
  if (e.glow) {
    const a = clampAlpha(e.glow.alpha) * 1000;
    out += `<a:glow rad="${Math.round((e.glow.size ?? 8) * PT)}">` +
           `<a:srgbClr val="${hex(e.glow.color)}"><a:alpha val="${a}"/></a:srgbClr></a:glow>`;
  }
  if (e.shadow) {
    const a = clampAlpha(e.shadow.alpha) * 1000;
    out += `<a:outerShdw blurRad="${Math.round((e.shadow.blur ?? 4) * PT)}" ` +
           `dist="${Math.round((e.shadow.dist ?? 3) * PT)}" ` +
           `dir="${Math.round((e.shadow.angle ?? 45) * ANG)}" rotWithShape="0">` +
           `<a:srgbClr val="${hex(e.shadow.color)}"><a:alpha val="${a}"/></a:srgbClr></a:outerShdw>`;
  }
  if (e.reflection) {
    out += `<a:reflection blurRad="6350" stA="50000" stPos="0" endA="300" endPos="55000" ` +
           `dist="0" dir="5400000" fadeDir="5400000" sx="100000" sy="-100000" ` +
           `kx="0" ky="0" algn="bl" rotWithShape="0"/>`;
  }
  if (e.softEdge) out += `<a:softEdge rad="${Math.round((e.softEdge.size ?? 3) * PT)}"/>`;
  return out ? `<a:effectLst>${out}</a:effectLst>` : "";
}

// ---- surgical XML helpers (scoped to one shape block) --------------------
const SP_RE = /<p:sp>[\s\S]*?<\/p:sp>/g;
const FILL_BEFORE_LN =
  /<a:(?:solidFill|noFill|gradFill|pattFill|blipFill|grpFill)\b[^>]*(?:\/>|>[\s\S]*?<\/a:(?:solidFill|gradFill|pattFill|blipFill|grpFill)>)(?=\s*<a:ln\b)/;

function applyToShape(spBlock, enh, warn) {
  let sp = spBlock;
  // 1) shape fill — replace the fill element that sits just before <a:ln>
  const newFill = shapeFillXml(enh.fill);
  if (newFill) {
    if (FILL_BEFORE_LN.test(sp)) sp = sp.replace(FILL_BEFORE_LN, newFill);
    else warn(`could not locate shape fill for "${enh.objectName}"`);
  }
  // 2) effectLst — insert just before </p:spPr> (after <a:ln>, per schema order)
  const eff = effectLstXml(enh);
  if (eff) {
    if (sp.includes("<a:effectLst>"))
      sp = sp.replace(/<a:effectLst>[\s\S]*?<\/a:effectLst>/, eff);
    else
      sp = sp.replace("</p:spPr>", `${eff}</p:spPr>`);
  }
  // 3) gradient text — swap every run's solidFill inside the text body
  if (enh.textGradient) {
    const grad = gradFillXml(enh.textGradient.stops || [], enh.textGradient.angle ?? 90);
    sp = sp.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, (tx) =>
      tx.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/g, grad));
  }
  return sp;
}

function enhanceSlideXml(xml, byName, warn) {
  return xml.replace(SP_RE, (sp) => {
    const name = (sp.match(/<p:cNvPr[^>]*\bname="([^"]*)"/) || [])[1];
    const enh = name && byName.get(name);
    if (!enh) return sp;
    byName.get(name)._hit = true;
    return applyToShape(sp, enh, warn);
  });
}

// ---- public entry --------------------------------------------------------
export async function enhancePptx(buffer, enhancements) {
  if (!Array.isArray(enhancements) || enhancements.length === 0) return buffer;
  const warnings = [];
  const warn = (m) => warnings.push(m);
  const byName = new Map();
  for (const e of enhancements) if (e && e.objectName) byName.set(e.objectName, e);

  const zip = await JSZip.loadAsync(buffer);
  const slidePaths = Object.keys(zip.files)
    .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p));

  for (const p of slidePaths) {
    const xml = await zip.file(p).async("string");
    const out = enhanceSlideXml(xml, byName, warn);
    if (out !== xml) zip.file(p, out);
  }
  for (const [name, e] of byName) if (!e._hit) warn(`objectName "${name}" not found on any slide`);

  const outBuf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  if (warnings.length) console.warn("[enhancePptx] warnings:\n  - " + warnings.join("\n  - "));
  return outBuf;
}

export default enhancePptx;

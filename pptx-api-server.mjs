import express from 'express';
import cors from 'cors';
import PptxGenJS from 'pptxgenjs';
import axios from 'axios';
import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import { fileURLToPath } from 'url';
import { Buffer } from 'buffer';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// The company part of the deck's file name: keeps Japanese and other letters (before: everything but a-z0-9 became "_",
// so 「サンプル商事株式会社」 gave "___________"); replaces only what file systems reject, spaces and controls; max 80 chars.
export function safeFilePart(name) {
  const s = String(name == null ? '' : name).normalize('NFC')
    .replace(/[\\/:*?"<>|\u0000-\u001f\u007f]/g, '_').replace(/\s+/g, '_')
    .replace(/_+/g, '_').replace(/^[_.]+|[_.]+$/g, '').slice(0, 80);
  return s || 'Company';
}

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// Optional bearer-token auth. If RENDER_TOKEN is set (recommended when the
// server is publicly reachable, e.g. on Easypanel), /generate-pptx and
// /download require it — via `Authorization: Bearer <token>` or `?token=`.
// /health stays open. When RENDER_TOKEN is unset (local use), auth is off.
// GET /d/<id>/<file> download links are open too: the 128-bit random id is the
// capability, so the master token never ends up in a link shown to users.
const RENDER_TOKEN = process.env.RENDER_TOKEN || '';
app.use((req, res, next) => {
  if (!RENDER_TOKEN || req.path === '/health' || (req.method === 'GET' && req.path === '/')) return next();
  if (req.method === 'GET' && req.path.startsWith('/d/')) return next();
  const auth = req.headers.authorization || '';
  if (auth === 'Bearer ' + RENDER_TOKEN || req.query.token === RENDER_TOKEN) return next();
  return res.status(401).json({ success: false, error: 'unauthorized' });
});

const PORT = process.env.PORT || 3456;
const OUTPUT_DIR = path.join(__dirname, 'output');

// Ensure output directory exists
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

// Per-run download links (?link=1): each deck goes to output/dl/<random id>/<filename>
// and expires after DOWNLOAD_TTL_HOURS. Separate folders also stop two runs for the
// same company from overwriting each other's file.
const DOWNLOAD_DIR = path.join(OUTPUT_DIR, 'dl');
const DOWNLOAD_TTL_MS = Number(process.env.DOWNLOAD_TTL_HOURS || 24) * 3600 * 1000;
fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });

function purgeExpiredDownloads() {
  for (const id of fs.readdirSync(DOWNLOAD_DIR)) {
    const dir = path.join(DOWNLOAD_DIR, id);
    try {
      if (Date.now() - fs.statSync(dir).mtimeMs > DOWNLOAD_TTL_MS) fs.rmSync(dir, { recursive: true, force: true });
    } catch { /* already gone */ }
  }
}
setInterval(purgeExpiredDownloads, 3600 * 1000).unref();


// ============================================================================
// REQUEST PARSING & NORMALIZATION HELPERS (supports Dify string payloads)
// ============================================================================

function safeJsonParse(maybeJson, maxDepth = 2) {
  let v = maybeJson;
  for (let i = 0; i < maxDepth; i++) {
    if (typeof v !== 'string') return v;
    const s = v.trim();
    if (!s) return v;
    try {
      v = JSON.parse(s);
    } catch {
      return v; // not JSON
    }
  }
  return v;
}

function toNum(v) {
  const n = Number(String(v ?? 0).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : 0;
}

function normalizeMetric(m) {
  // Accept a variety of inputs and always return the expected shape.
  if (m == null) return { name: '', current: 0, comparison: 0, items: [] };
  if (typeof m === 'number') return { name: '', current: toNum(m), comparison: 0, items: [] };
  if (typeof m === 'string') return { name: m, current: 0, comparison: 0, items: [] };

  const name = m.name ?? m.label ?? m.title ?? '';
  const current = toNum(m.current ?? m.actual ?? m.act ?? 0);
  const comparison = toNum(m.comparison ?? m.budget ?? m.prev ?? 0);

  const items = Array.isArray(m.items)
    ? m.items.map(it => ({
        name: it?.name ?? it?.label ?? it?.title ?? '',
        current: toNum(it?.current ?? it?.actual ?? it?.act ?? 0),
        comparison: toNum(it?.comparison ?? it?.budget ?? it?.prev ?? 0),
      }))
    : [];

  return { name, current, comparison, items };
}

function normalizeFinancialData(raw) {
  let d = safeJsonParse(raw, 3);

  // Some workflows send { financial_data: {...} }
  if (d && typeof d === 'object' && d.financial_data && typeof d.financial_data === 'object') {
    d = d.financial_data;
  }

  // Some workflows send { financialData: {...} } inside the string
  if (d && typeof d === 'object' && d.financialData && typeof d.financialData === 'object') {
    d = d.financialData;
  }

  // Period labels
  const current_period_label =
    d.current_period_label ??
    d.currentPeriodLabel ??
    d.period_label ??
    d.period_columns?.current ??
    d.periodColumns?.current ??
    d.period?.current ??
    '';
  const comparison_period_label =
    d.comparison_period_label ??
    d.comparisonPeriodLabel ??
    d.comparison_label ??
    d.period_columns?.comparison ??
    d.periodColumns?.comparison ??
    d.period?.comparison ??
    '';

  // Metric name mapping (Dify sample uses net_profit)
  const profit_after_tax_src = d.profit_after_tax ?? d.net_profit ?? d.net_income ?? d.netProfit ?? d.profitAfterTax ?? {};
  const finance_costs_src = d.finance_costs ?? d.other_expenses ?? d.otherExpenses ?? {};

  const out = {
    company_name: d.company_name ?? d.companyName ?? d.company ?? 'Company',
    current_period_label,
    comparison_period_label,
    revenue: normalizeMetric(d.revenue),
    cogs: normalizeMetric(d.cogs),
    gross_profit: normalizeMetric(d.gross_profit),
    operating_expenses: normalizeMetric(d.operating_expenses),
    other_income: normalizeMetric(d.other_income),
    finance_costs: normalizeMetric(finance_costs_src),
    profit_before_tax: normalizeMetric(d.profit_before_tax),
    tax_expense: normalizeMetric(d.tax_expense),
    profit_after_tax: normalizeMetric(profit_after_tax_src),
    // Pass through any budget payload if already present (optional)
    budget_data: d.budget_data ?? d.budgetData ?? {},
    // n8n v2: Actual / Budget / Prior Year lines. Its presence switches on the v2 slides.
    three_way: d.three_way ?? null,
    account_names_ja: d.account_names_ja ?? null,
  };

  // Ensure required nested structures for slide 2 logic
  if (!out.operating_expenses.items) out.operating_expenses.items = [];
  if (!Array.isArray(out.operating_expenses.items)) out.operating_expenses.items = [];

  // If the source had opex item arrays elsewhere, try to pick them up
  if (out.operating_expenses.items.length === 0 && Array.isArray(d.operating_expense_items)) {
    out.operating_expenses.items = d.operating_expense_items.map(it => ({
      name: it.name ?? '',
      current: toNum(it.current ?? it.actual ?? 0),
      comparison: toNum(it.comparison ?? it.budget ?? 0),
    }));
  }

  return out;
}

// ============================================================================
// NATIVE P&L WATERFALL (slide 2)
// One stacked-column chart: an invisible spacer series lifts the floating bars.
// Bridge per period: Revenue -> COGS -> Gross Profit -> OPEX -> Other, net -> Net Income,
// where "Other, net" = Net Income - (Gross Profit - OPEX) so the bridge always closes
// (other income - finance costs - tax).
// ============================================================================

const WATERFALL_SPACER = '(spacer)';

function buildWaterfallSeries(d, periods, names = ['Revenue', 'COGS', 'Gross Profit', 'OPEX', 'Other, net', 'Net Income']) {
  const steps = [];
  for (const p of periods) {
    const v = (m) => toNum(d?.[m]?.[p.key] ?? 0);
    const revenue = v('revenue');
    const grossProfit = v('gross_profit') || (revenue - v('cogs'));
    const cogs = revenue - grossProfit;
    const opex = v('operating_expenses');
    const netIncome = v('profit_after_tax');
    const other = netIncome - (grossProfit - opex);
    steps.push(
      { label: names[0], period: p.label, kind: 'total', value: revenue },
      { label: names[1], period: p.label, kind: 'delta', value: -cogs },
      { label: names[2], period: p.label, kind: 'total', value: grossProfit },
      { label: names[3], period: p.label, kind: 'delta', value: -opex },
      { label: names[4], period: p.label, kind: 'delta', value: other },
      { label: names[5], period: p.label, kind: 'total', value: netIncome },
    );
  }

  const spacer = [], total = [], increase = [], decrease = [];
  let running = 0;
  for (const s of steps) {
    if (s.kind === 'total') {
      running = s.value;
      spacer.push(0); total.push(s.value); increase.push(0); decrease.push(0);
    } else {
      const next = running + s.value;
      spacer.push(Math.min(running, next));
      total.push(0);
      increase.push(s.value > 0 ? s.value : 0);
      decrease.push(s.value < 0 ? -s.value : 0);
      running = next;
    }
  }
  // Floating bars only work while the bridge stays above zero; callers fall back otherwise.
  const valid = steps.every((s, i) => spacer[i] >= 0 && total[i] >= 0) && steps.some(s => s.value !== 0);
  const outer = steps.map((s, i) => (i % 6 === 0 ? s.period : ''));
  // Floating-bar labels sit just outside the bar (on the invisible spacer), in thousands.
  const thousands = (n) => Math.round(Math.abs(n) / 1000).toLocaleString('en-US');
  const deltaLabels = steps.map(s => (s.kind === 'delta' ? `${s.value < 0 ? '−' : '+'}${thousands(s.value)}` : null));
  return { valid, steps, inner: steps.map(s => s.label), outer, spacer, total, increase, decrease, deltaLabels };
}

// pptxgenjs cannot hide one series or format labels per series, so patch the chart XML:
// spacer -> no fill + no labels; totals -> white labels; deltas -> dark labels, "−" prefix on decreases.
async function patchWaterfallCharts(pptxBase64, deltaLabels, fontFace) {
  const esc = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const spacerLabels = deltaLabels
    .map((text, idx) => (text === null ? '' : `<c:dLbl><c:idx val="${idx}"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="800"><a:solidFill><a:srgbClr val="1F2937"/></a:solidFill><a:latin typeface="${esc(fontFace)}"/></a:rPr><a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx><c:dLblPos val="inEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbl>`))
    .join('');
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(pptxBase64, { base64: true });
  let patched = 0;
  for (const name of Object.keys(zip.files).filter(n => /^ppt\/charts\/chart\d+\.xml$/.test(n))) {
    let xml = await zip.file(name).async('string');
    if (!xml.includes(`<c:v>${WATERFALL_SPACER}</c:v>`)) continue;
    xml = xml.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, (ser) => {
      const seriesName = (/<c:tx>[\s\S]*?<c:v>([^<]*)<\/c:v>/.exec(ser) || [])[1];
      if (seriesName === WATERFALL_SPACER) {
        return ser
          .replace(/<c:spPr>[\s\S]*?<\/c:spPr>/, '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>')
          .replace(/<c:dLbls>[\s\S]*?<\/c:dLbls>/, `<c:dLbls>${spacerLabels}<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`);
      }
      if (seriesName !== 'Total') {
        return ser.replace(/<c:dLbls>[\s\S]*?<\/c:dLbls>/, '<c:dLbls><c:delete val="1"/></c:dLbls>');
      }
      return ser
        .replace(/(<c:dLbls>[\s\S]*?)<c:numFmt [^>]*\/>/, `$1<c:numFmt formatCode="#,##0,;&quot;−&quot;#,##0,;;" sourceLinked="0"/>`)
        .replace(/(<c:dLbls>[\s\S]*?<a:solidFill>)<a:srgbClr val="[0-9A-Fa-f]{6}"\/>/, '$1<a:srgbClr val="FFFFFF"/>');
    });
    zip.file(name, xml);
    patched++;
  }
  if (!patched) return pptxBase64;
  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}

// Kanji/kana render in Yu Gothic while Latin text and numbers stay in the primary font:
// pptxgenjs writes the same face into <a:ea>, so point every East Asian slot at Yu Gothic
// (charts only carry <a:latin>, so add the slot there). Japanese decks are tagged ja-JP.
async function applyEastAsianFont(pptxBase64, ja) {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(pptxBase64, { base64: true });
  for (const name of Object.keys(zip.files).filter(n => /^ppt\/(slides\/slide|charts\/chart)\d+\.xml$/.test(n))) {
    let xml = await zip.file(name).async('string');
    xml = xml.replace(/<a:ea typeface="[^"]*"/g, '<a:ea typeface="Yu Gothic"');
    xml = xml.replace(/(<a:latin typeface="[^"]*"[^>]*\/>)(?!<a:ea)/g, '$1<a:ea typeface="Yu Gothic"/>');
    if (ja) xml = xml.replace(/ lang="en-US"/g, ' lang="ja-JP"');
    zip.file(name, xml);
  }
  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}

// ============================================================================
// PPTX GENERATION FUNCTION (From your n8n code)
// ============================================================================

async function generatePresentation(config) {
  const { pptxConfig, financialData, coverImageUrl, chartImageUrl, chartData, shopCommentary, executiveSummary } = config;
  // Deck language: 'ja' or 'en' (default). Every fixed label goes through t(english, japanese).
  const JA = String(config.language || pptxConfig?.language || '').toLowerCase() === 'ja';
  const t = (en, ja) => (JA ? ja : en);
  // Account lines from the CSV shown under their Japanese names (map sent by the workflow).
  const ACCOUNT_JA = (financialData && financialData.account_names_ja) || {};
  const acct = (name) => (JA && ACCOUNT_JA[String(name ?? '').trim()]) || name;
  
  // Configuration loading
  const cfg = pptxConfig || {};
  const globalStyles = cfg;
  const layoutConfig = cfg;
  const slideConfig = cfg;
  const textContent = cfg;

  // ============================================================================
  // MERGED CONFIGURATION WITH DEFAULTS
  // ============================================================================

  const CONFIG = {
    // === TYPOGRAPHY ===
    fonts: {
      primary: globalStyles.font_primary || 'Aptos',
      fallback: globalStyles.font_fallback || 'Calibri, Arial'
    },
    
    fontSizes: {
      title: globalStyles.font_size_title || 20,
      subtitle: globalStyles.font_size_subtitle || 10,
      table: globalStyles.font_size_table || 8,
      small: globalStyles.font_size_small || 8,
      chartLabel: globalStyles.font_size_chart_label || 8,
      chartLegend: globalStyles.font_size_chart_legend || 8,
      takeaway: globalStyles.font_size_takeaway || 10,
      kpi: globalStyles.font_size_kpi || 12
    },
    
    // === COLORS ===
    colors: {
      primary: globalStyles.color_primary || '059669',
      primaryDark: globalStyles.color_primary_dark || '047857',
      black: globalStyles.color_black || '000000',
      white: globalStyles.color_white || 'FFFFFF',
      highlight: globalStyles.color_highlight || 'E2DDF9',
      positive: globalStyles.color_positive || '059669',
      negative: globalStyles.color_negative || 'DC2626',
      gray50: globalStyles.color_gray_50 || 'F9FAFB',
      gray100: globalStyles.color_gray_100 || 'F3F4F6',
      gray500: globalStyles.color_gray_500 || '6B7280',
      gray700: globalStyles.color_gray_700 || '374151',
      gray800: globalStyles.color_gray_800 || '1F2937',
      gray900: globalStyles.color_gray_900 || '111827',
      tableBorder: globalStyles.color_table_border || 'E5E7EB',
      tableHeaderBg: globalStyles.color_table_header_bg || '000000',
      takeawayBg: globalStyles.color_takeaway_bg || 'E8F5E9',
      chartBar1: globalStyles.color_chart_bar_1 || '111827',
      chartBar2: globalStyles.color_chart_bar_2 || '9CA3AF',
      chartBar3: globalStyles.color_chart_bar_3 || '10B981'
    },
    
    // === LAYOUT POSITIONS (in inches) ===
    layout: {
      marginLeft: layoutConfig.layout_margin_left || 0.5,
      marginRight: layoutConfig.layout_margin_right || 0.5,
      titleY: layoutConfig.layout_title_y || 0.15,
      subtitleY: layoutConfig.layout_subtitle_y || 0.6,
      contentStartY: layoutConfig.layout_content_start_y || 1.05,
      takeawaysY: layoutConfig.layout_takeaways_y || 4.65,
      footerY: layoutConfig.layout_footer_y || 7.27
    },
    
    // === SLIDE 2 (Income Statement) ===
    slide2: {
      graphX: slideConfig.slide2_graph_x || 0.5,
      graphW: slideConfig.slide2_graph_w || 4.331,
      graphH: slideConfig.slide2_graph_h || 3.425,
      tableX: slideConfig.slide2_table_x || 5.163,
      tableW: slideConfig.slide2_table_w || 4.331,
      tableColMetric: slideConfig.slide2_table_col_metric || 1.9685,
      tableColValue: slideConfig.slide2_table_col_value || 0.7874,
      tableRowHeight: slideConfig.slide2_table_row_height || 0.16,
      takeawayBoxX: slideConfig.slide2_takeaway_box_x || 0.467,
      takeawayBoxW: slideConfig.slide2_takeaway_box_w || 1.5,
      takeawayItemW: slideConfig.slide2_takeaway_item_w || 2.402,
      takeawayItemH: slideConfig.slide2_takeaway_item_h || 0.35,
      takeawayGap: slideConfig.slide2_takeaway_gap || 0.122
    },
    
    // === SLIDE 3 (Dashboard) ===
    slide3: {
      gaugeY: slideConfig.slide3_gauge_y || 0.8,
      gaugeW: slideConfig.slide3_gauge_w || 1.5,
      gaugeH: slideConfig.slide3_gauge_h || 0.7,
      gaugeSpacing: slideConfig.slide3_gauge_spacing || 1.65,
      gaugeHoleSize: slideConfig.slide3_gauge_hole_size || 70,
      tableX: slideConfig.slide3_table_x || 7.2,
      tableW: slideConfig.slide3_table_w || 2.7,
      chartEbitX: slideConfig.slide3_chart_ebit_x || 0.5,
      chartEbitY: slideConfig.slide3_chart_ebit_y || 2.15,
      chartEbitW: slideConfig.slide3_chart_ebit_w || 3.3,
      chartEbitH: slideConfig.slide3_chart_ebit_h || 1.4,
      chartOpexX: slideConfig.slide3_chart_opex_x || 4.0,
      chartOpexW: slideConfig.slide3_chart_opex_w || 3.0,
      chartRevenueY: slideConfig.slide3_chart_revenue_y || 3.95
    },
    
    // === SLIDE 4 (Budget vs Actual) ===
    slide4: {
      tableX: slideConfig.slide4_table_x || 0.57,
      tableY: slideConfig.slide4_table_y || 1.17,
      tableW: slideConfig.slide4_table_w || 9.06,
      tableRowH: slideConfig.slide4_table_row_h || 0.45,
      chartY: slideConfig.slide4_chart_y || 5.95,
      chartW: slideConfig.slide4_chart_w || 2.5,
      chartH: slideConfig.slide4_chart_h || 1.2,
      chart1X: slideConfig.slide4_chart1_x || 1.0,
      chart2X: slideConfig.slide4_chart2_x || 4.5,
      chart3X: slideConfig.slide4_chart3_x || 8.0
    },
    
    // === COVER SLIDE ===
    cover: {
      bgColor: slideConfig.cover_bg_color || '0d4d29',
      diagonalTopPct: slideConfig.cover_diagonal_top_pct || 47.5,
      diagonalBottomPct: slideConfig.cover_diagonal_bottom_pct || 30,
      titleX: slideConfig.cover_title_x || 7.05,
      titleY: slideConfig.cover_title_y || 1.35,
      titleFontSize: slideConfig.cover_title_font_size || 54
    },
    
    // === CHART STYLING ===
    charts: {
      lineSize: globalStyles.chart_line_size || 2,
      symbolSize: globalStyles.chart_symbol_size || 3,
      legendPosition: globalStyles.chart_legend_position || 't',
      gridColor: globalStyles.chart_grid_color || 'E5E7EB',
      gridStyle: globalStyles.chart_grid_style || 'solid',
      showLegend: globalStyles.chart_show_legend !== false,
      showValues: globalStyles.chart_show_values || false
    },
    
    // === TABLE STYLING ===
    tables: {
      borderWidth: globalStyles.table_border_width || 1,
      borderColor: globalStyles.table_border_color || 'E5E7EB',
      headerBg: globalStyles.table_header_bg || '000000',
      headerText: globalStyles.table_header_text || 'FFFFFF',
      altRowBg: globalStyles.table_alt_row_bg || 'F3F4F6',
      highlightRowBg: globalStyles.table_highlight_row_bg || 'E2DDF9',
      totalRowBg: globalStyles.table_total_row_bg || '059669',
      cellPaddingTop: globalStyles.table_cell_padding_top || 0.02,
      cellPaddingSide: globalStyles.table_cell_padding_side || 0.1
    },
    
    // === TEXT CONTENT ===
    text: {
      slide2Title: JA ? '主要財務指標：売上総利益と当期純利益' : (textContent.text_slide2_title || 'Key monthly financials metrics with gross and net profit'),
      slide3Title: JA ? '財務ダッシュボードと損益計算書' : (textContent.text_slide3_title || 'Monthly financial metrics dashboard with income statement'),
      slide4Title: JA ? '店舗別売上高と予算差異' : (textContent.text_slide4_title || 'Monthly financials metrics with revenue and gross profit'),
      footer: JA ? 'グラフはExcelデータと連動しています。グラフを右クリックし「データの編集」から更新できます。' : (textContent.text_footer || 'This graph/chart is linked to excel, and changes automatically based on data. Just left click on it and select "edit data"'),
      takeawayTemplate1: JA ? '売上高の前年同期比 {var}%' : (textContent.text_takeaway_template_1 || 'Year-on-year revenue change of {var}%.'),
      takeawayTemplate2: JA ? '売上総利益率 {var}%' : (textContent.text_takeaway_template_2 || 'Gross profit margin of {var}% of revenue.'),
      takeawayTemplate3: JA ? '当期純利益の前年同期差 {var}百万（現地通貨）' : (textContent.text_takeaway_template_3 || 'Year-on-year net profit change of {var}M (local currency).')
    },
    
    // === OPEX BREAKDOWN ===
    opex: {
      salesPct: slideConfig.opex_sales_pct || 50,
      marketingPct: slideConfig.opex_marketing_pct || 22,
      adminPct: slideConfig.opex_admin_pct || 28
    },
    
    // === PRESENTATION METADATA ===
    metadata: {
      layout: slideConfig.pptx_layout || 'LAYOUT_16x9',
      author: slideConfig.pptx_author || 'Financial Reporting System',
      filenamePrefix: slideConfig.filename_prefix || 'Financial_Metrics'
    }
  };

  // ============================================================================
  // HELPER FUNCTIONS
  // ============================================================================

	// --------------------------------------------------------------------------
	// Robust input defaults (prevents crashes when optional fields are missing)
	// --------------------------------------------------------------------------
	const inputData = (() => {
		const d = financialData || {};
		const obj = (v, def = {}) => (v && typeof v === 'object' && !Array.isArray(v) ? v : def);
		const arr = (v) => (Array.isArray(v) ? v : []);
		const metric = (v) => {
			const o = obj(v, {});
			return {
				current: toNum(o.current ?? 0),
				comparison: toNum(o.comparison ?? 0),
				items: arr(o.items).map(it => ({
					name: String(it?.name ?? ''),
					current: toNum(it?.current ?? 0),
					comparison: toNum(it?.comparison ?? 0),
				}))
			};
		};

		const normalized = {
			...d,
			company_name: d.company_name || d.companyName || '',
			current_period_label: d.current_period_label || d.currentPeriodLabel || '',
			comparison_period_label: d.comparison_period_label || d.comparisonPeriodLabel || '',
			revenue: metric(d.revenue || d.revenueData),
			cogs: metric(d.cogs || d.cogsData),
			gross_profit: metric(d.gross_profit || d.grossProfit),
			operating_expenses: metric(d.operating_expenses || d.operatingExpenses || d.opex),
			other_income: metric(d.other_income || d.otherIncome),
			finance_costs: metric(d.finance_costs || d.financeCosts),
			profit_before_tax: metric(d.profit_before_tax || d.profitBeforeTax),
			tax_expense: metric(d.tax_expense || d.taxExpense),
			profit_after_tax: metric(d.profit_after_tax || d.profitAfterTax),
			budget_data: obj(d.budget_data || d.budgetData || {}, {}),
		};

		// Ensure operating_expenses.items exists even if caller omits it
		normalized.operating_expenses.items = arr(normalized.operating_expenses.items);
		return normalized;
	})();

	const COVER_IMAGE_URL = coverImageUrl || pptxConfig?.cover_image_url || pptxConfig?.coverImageUrl || null;
	const CHART_IMAGE_URL = chartImageUrl || pptxConfig?.chart_image_url || pptxConfig?.chartImageUrl || null;

  async function downloadImageAsDataUri(urlOrPath) {
    if (!urlOrPath || urlOrPath === '') return null;
    if (urlOrPath.startsWith('http://') || urlOrPath.startsWith('https://')) {
      try {
        const response = await axios.get(urlOrPath, { responseType: 'arraybuffer', timeout: 30000 });
        const mime = response.headers['content-type'] || 'image/png';
        return `data:${mime};base64,${Buffer.from(response.data).toString('base64')}`;
      } catch (e) { 
        console.error('Failed to download image:', urlOrPath, e.message);
        return null; 
      }
    }
    try {
      if (!fs.existsSync(urlOrPath)) return null;
      const ext = path.extname(urlOrPath).toLowerCase().replace(".", "");
      const mime = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "webp" ? "image/webp" : "image/png";
      return `data:${mime};base64,${fs.readFileSync(urlOrPath).toString("base64")}`;
    } catch (e) { return null; }
  }

  function svgToDataUriCover(svg) {
    return "data:image/svg+xml;base64," + Buffer.from(svg, "utf8").toString("base64");
  }

  function formatCurrency(num) {
    const n = Number(num ?? 0);
    if (!Number.isFinite(n)) return '0';
    return Math.round(n).toLocaleString();
  }

  function calculateVariance(current, comparison) {
    if (comparison === 0) return '0';
    const variance = ((current - comparison) / comparison * 100);
    return (variance >= 0 ? '+ ' : '') + variance.toFixed(0) + '%';
  }

  // "Posting periods 001/2025 to 003/2025" -> "Jan–Mar 2025" (a single period -> "Mar 2025").
  // Japanese: "2025年1月〜3月" / "2025年3月".
  function periodRange(label, months) {
    const s = String(label ?? '');
    const m = s.match(/(\d{3})\/(\d{4})\s*to\s*(\d{3})\/(\d{4})/i);
    const name = (p) => months[parseInt(p, 10) - 1];
    if (JA) {
      const mo = (p) => parseInt(p, 10);
      if (m && mo(m[1]) >= 1 && mo(m[3]) <= 12) {
        if (m[1] === m[3] && m[2] === m[4]) return `${m[4]}年${mo(m[3])}月`;
        if (m[2] === m[4]) return `${m[4]}年${mo(m[1])}月〜${mo(m[3])}月`;
        return `${m[2]}年${mo(m[1])}月〜${m[4]}年${mo(m[3])}月`;
      }
      const one = s.match(/(\d{3})\/(\d{4})/);
      return one ? `${one[2]}年${mo(one[1])}月` : '';
    }
    if (m && name(m[1]) && name(m[3])) {
      if (m[1] === m[3] && m[2] === m[4]) return `${name(m[3])} ${m[4]}`;
      if (m[2] === m[4]) return `${name(m[1])}–${name(m[3])} ${m[4]}`;
      return `${name(m[1])} ${m[2]}–${name(m[3])} ${m[4]}`;
    }
    const one = s.match(/(\d{3})\/(\d{4})/);
    return one && name(one[1]) ? `${name(one[1])} ${one[2]}` : '';
  }

  function extractPeriodLabel(label) {
    return periodRange(label, ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']) || String(label ?? '');
  }

  const even = (n) => 2 * Math.round(n / 2);

  function addSlideTitle(slide, text) {
    slide.addText(text, {
      x: CONFIG.layout.marginLeft,
      y: CONFIG.layout.titleY,
      w: 9,
      h: 0.5,
      fontSize: CONFIG.fontSizes.title,
      bold: true,
      color: CONFIG.colors.gray900,
      fontFace: CONFIG.fonts.primary,
      align: 'left',
      valign: 'top'
    });
  }

  function addSlideSubtitle(slide, text) {
    slide.addText(text, {
      x: CONFIG.layout.marginLeft,
      y: CONFIG.layout.subtitleY,
      w: 9,
      h: 0.3,
      fontSize: CONFIG.fontSizes.subtitle,
      color: CONFIG.colors.gray500,
      fontFace: CONFIG.fonts.primary,
      align: 'left',
      valign: 'top'
    });
  }

  function addKeyTakeaways(slide, pptx, takeaways) {
    slide.addShape(pptx.ShapeType.rect, {
      x: CONFIG.slide2.takeawayBoxX,
      y: CONFIG.layout.takeawaysY,
      w: CONFIG.slide2.takeawayBoxW,
      h: CONFIG.slide2.takeawayItemH,
      fill: { color: CONFIG.colors.black },
      line: { type: 'none' }
    });

    slide.addText(t('Key Takeaways', '要点'), {
      x: CONFIG.slide2.takeawayBoxX,
      y: CONFIG.layout.takeawaysY,
      w: CONFIG.slide2.takeawayBoxW,
      h: CONFIG.slide2.takeawayItemH,
      fontSize: CONFIG.fontSizes.takeaway,
      bold: true,
      color: CONFIG.colors.white,
      align: 'center',
      valign: 'middle',
      fontFace: CONFIG.fonts.primary
    });

    const startX = CONFIG.slide2.takeawayBoxX + CONFIG.slide2.takeawayBoxW + CONFIG.slide2.takeawayGap;
    
    takeaways.forEach((text, idx) => {
      const x = startX + (idx * (CONFIG.slide2.takeawayItemW + CONFIG.slide2.takeawayGap));
      
      slide.addText(text, {
        x: x,
        y: CONFIG.layout.takeawaysY,
        w: CONFIG.slide2.takeawayItemW,
        h: CONFIG.slide2.takeawayItemH,
        fontSize: CONFIG.fontSizes.takeaway,
        color: CONFIG.colors.gray700,
        align: 'center',
        valign: 'middle',
        fontFace: CONFIG.fonts.primary,
        fill: { color: CONFIG.colors.takeawayBg }
      });
    });
  }

  function createTableHeader(columns) {
    return columns.map(col => ({
      text: col,
      options: {
        bold: true,
        fontSize: CONFIG.fontSizes.table,
        fill: CONFIG.colors.tableHeaderBg,
        color: CONFIG.colors.white,
        align: col === columns[0] ? 'left' : 'center',
        valign: 'middle',
        fontFace: CONFIG.fonts.primary
      }
    }));
  }

  function addGaugeChart(slide, pptx, options) {
    const { x, value, title, color } = options;
    const y = options.y != null ? options.y : CONFIG.slide3.gaugeY;
    const w = CONFIG.slide3.gaugeW;
    const h = CONFIG.slide3.gaugeH;
    
    slide.addText(title, {
      x, y, w, h: 0.2,
      fontSize: CONFIG.fontSizes.table,
      bold: true,
      color: CONFIG.colors.gray800,
      fontFace: CONFIG.fonts.primary,
      align: 'center'
    });
    
    slide.addChart(pptx.ChartType.doughnut, [{
      name: 'Metric',
      labels: ['Value', 'Remaining'],
      values: [value, 100 - value]
    }], {
      x, y: y + 0.2, w, h,
      chartColors: [color, 'E5E7EB'],
      holeSize: CONFIG.slide3.gaugeHoleSize,
      showLegend: false,
      showValue: false,
      showPercent: false,
      showTitle: false
    });
    
    slide.addText(`${value}%`, {
      x, y: y + 0.42, w, h: 0.3,
      fontSize: CONFIG.fontSizes.chartLabel,
      bold: true,
      color: color,
      fontFace: CONFIG.fonts.primary,
      align: 'center'
    });
    
    slide.addText(options.periodLabel || '', {
      x, y: y + 0.95, w, h: 0.15,
      fontSize: CONFIG.fontSizes.small,
      color: CONFIG.colors.gray500,
      fontFace: CONFIG.fonts.primary,
      align: 'center'
    });
  }

  function addBudgetChart(slide, pptx, options) {
    // Clean KPI tile (actual vs budget) \u2014 the detailed bars live in the table above.
    const { x, title, actual, budget, varPct } = options;

    slide.addShape(pptx.ShapeType.rect, {
      x, y: 4.5, w: 2.75, h: 0.92,
      fill: { color: CONFIG.colors.gray50 }, line: { color: 'D1D5DB', width: 1 },
    });
    slide.addShape(pptx.ShapeType.rect, {
      x, y: 4.5, w: 0.05, h: 0.92,
      fill: { color: varPct >= 0 ? CONFIG.colors.positive : CONFIG.colors.negative }, line: { type: 'none' },
    });

    slide.addText(title, {
      x: x + 0.18, y: 4.56, w: 2.5, h: 0.24,
      fontSize: CONFIG.fontSizes.small, color: CONFIG.colors.gray500, fontFace: CONFIG.fonts.primary,
    });
    slide.addText(formatCurrency(actual), {
      x: x + 0.18, y: 4.78, w: 1.7, h: 0.34,
      fontSize: 16, bold: true, color: CONFIG.colors.gray900, fontFace: CONFIG.fonts.primary, valign: 'middle',
    });
    slide.addText(`${varPct >= 0 ? '\u25b2' : '\u25bc'} ${Math.abs(varPct)}%`, {
      x: x + 1.75, y: 4.78, w: 0.95, h: 0.34,
      fontSize: 12, bold: true, align: 'right', valign: 'middle',
      color: varPct >= 0 ? CONFIG.colors.positive : CONFIG.colors.negative, fontFace: CONFIG.fonts.primary,
    });
    slide.addText(t(`vs Budget ${formatCurrency(budget)}`, `予算 ${formatCurrency(budget)}`), {
      x: x + 0.18, y: 5.14, w: 2.5, h: 0.22,
      fontSize: CONFIG.fontSizes.small, color: CONFIG.colors.gray500, fontFace: CONFIG.fonts.primary,
    });
  }

  async function buildCoverBackgroundSvg(imagePathOrNull) {
    const W = 1280, H = 720;
    const xTop = Math.round((CONFIG.cover.diagonalTopPct / 100) * W);
    const xBottom = Math.round((CONFIG.cover.diagonalBottomPct / 100) * W);
    const green = `#${CONFIG.cover.bgColor}`;
    const white = "#ffffff";
    
    const imgHref = imagePathOrNull ? await downloadImageAsDataUri(imagePathOrNull) : null;

    return `<?xml version="1.0" encoding="UTF-8"?>
    <svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">
      <defs>
        <clipPath id="leftArea">
          <polygon points="0,0 ${xTop},0 ${xBottom},${H} 0,${H}"/>
        </clipPath>
      </defs>
      ${
        imgHref
          ? `<image href="${imgHref}" x="0" y="0" width="${xTop}" height="${H}" preserveAspectRatio="xMidYMid slice" clip-path="url(#leftArea)"/>`
          : `<polygon points="0,0 ${xTop},0 ${xBottom},${H} 0,${H}" fill="${white}"/>`
      }
      <polygon points="${xTop},0 ${W},0 ${W},${H} ${xBottom},${H}" fill="${green}"/>
    </svg>`;
  }

  async function addCoverSlide(pptx, inputData, coverImageUrl) {
    const slide = pptx.addSlide();
    const SLIDE_W = 10;
    const SLIDE_H = 5.625;
    const BASE_W = 13.333;
    const S = SLIDE_W / BASE_W;
    const sx = (v) => Number((v * S).toFixed(4));

    const periodLong = periodRange(inputData.current_period_label,
      ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']);

    slide.addImage({
      data: svgToDataUriCover(await buildCoverBackgroundSvg(coverImageUrl)),
      x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
    });

    // Japanese title on one line: widened to the left (still inside the green panel) and never wrapped
    slide.addText(t("Monthly\nFinancials\nMetrics", "月次業績レポート"), {
      x: sx(CONFIG.cover.titleX) - (JA ? 0.8 : 0), y: sx(CONFIG.cover.titleY), w: sx(5.85) + (JA ? 0.8 : 0), h: sx(3.40),
      ...(JA ? { wrap: false } : {}),
      fontFace: CONFIG.fonts.primary,
      fontSize: even(CONFIG.cover.titleFontSize * S),
      color: CONFIG.colors.white,
      bold: true,
      align: "right",
      valign: "mid",
    });

    slide.addShape(pptx.ShapeType.rect, {
      x: sx(8.10), y: sx(5.72), w: sx(4.70), h: sx(0.03),
      fill: { color: CONFIG.colors.white, transparency: 55 },
      line: { color: CONFIG.colors.white, transparency: 100 },
    });

    slide.addText(inputData.company_name || t("Your Company Name", "会社名"), {
      x: sx(CONFIG.cover.titleX), y: sx(5.78), w: sx(5.85), h: sx(0.40),
      fontFace: CONFIG.fonts.primary,
      fontSize: even(18 * S),
      color: CONFIG.colors.white,
      align: "right",
      valign: "top",
    });

    if (periodLong) {
      slide.addText(periodLong, {
        x: sx(CONFIG.cover.titleX), y: sx(6.08), w: sx(5.85), h: sx(0.35),
        fontFace: CONFIG.fonts.primary,
        fontSize: even(16 * S),
        color: CONFIG.colors.white,
        align: "right",
        valign: "top",
      });
    }
  }

  // ============================================================================
  // V2 SLIDES (only when financialData.three_way is sent): executive summary,
  // three-way P&L and OPEX deep-dive. Even font sizes only (10/12/14/16/20/24).
  // ============================================================================

  const TW = inputData.three_way && Array.isArray(inputData.three_way.lines) ? inputData.three_way : null;
  const V2 = { body: 12, small: 10, point: 14, headline: 16, kpi: 24 };
  const isNum = (n) => typeof n === 'number' && Number.isFinite(n);
  const pctChange = (a, b) => (isNum(a) && isNum(b) && b !== 0 ? ((a - b) / Math.abs(b)) * 100 : null);
  const fmtPct = (p) => (p === null ? '–' : `${p >= 0 ? '+' : '−'}${Math.abs(p).toFixed(1)}%`);
  const fmtAmt = (n) => (isNum(n) ? formatCurrency(n) : '–');
  const fmtDiff = (n) => (isNum(n) ? `${n >= 0 ? '+' : '−'}${formatCurrency(Math.abs(n))}` : '–');
  // Cost lines are favourable when they come in below the comparison.
  const varColor = (p, cost) => (p === null || p === 0 ? CONFIG.colors.gray700 : ((p > 0) !== !!cost ? CONFIG.colors.positive : CONFIG.colors.negative));
  const twLine = (key) => (TW ? TW.lines.find(l => l.key === key) : null) || { actual: 0, budget: null, prior: null };
  const twPeriod = extractPeriodLabel(inputData.current_period_label) || t('Current', '当期');
  const twPriorPeriod = extractPeriodLabel(inputData.comparison_period_label) || t('Prior Year', '前年同期');
  const TW_LABELS_JA = { revenue: '売上高', cogs: '売上原価', gross_profit: '売上総利益', operating_expenses: '営業費用', operating_profit: '営業利益', other_income: 'その他収益', finance_costs: '金融費用', profit_after_tax: '当期純利益' };
  const twLabel = (l) => (JA && TW_LABELS_JA[l.key]) || l.label;

  function addV2Takeaways(slide, pptx, texts) {
    addKeyTakeaways(slide, pptx, texts.slice(0, 3));
  }

  // Computed fallback when the AI summary is missing or unparseable.
  function computedSummary() {
    const rev = twLine('revenue'), op = twLine('operating_profit'), pat = twLine('profit_after_tax'), opex = twLine('operating_expenses');
    const items = TW.opex_items || [];
    const top = [...items].sort((a, b) => b.actual - a.actual)[0];
    const revGrowth = pctChange(rev.actual, rev.prior);
    const fastCosts = items
      .filter(it => opex.actual && it.actual / opex.actual >= 0.05)
      .map(it => ({ ...it, g: pctChange(it.actual, it.prior) }))
      .filter(it => it.g !== null && revGrowth !== null && it.g > revGrowth)
      .sort((a, b) => b.actual - a.actual);
    return {
      headline: t(`Revenue ${fmtAmt(rev.actual)}: ${fmtPct(pctChange(rev.actual, rev.budget))} vs budget, ${fmtPct(revGrowth)} vs prior year.`,
        `売上高${fmtAmt(rev.actual)}：予算比${fmtPct(pctChange(rev.actual, rev.budget))}、前年同期比${fmtPct(revGrowth)}。`),
      points: [
        { title: t('Against budget', '予算との比較'), text: t(`Revenue ${fmtPct(pctChange(rev.actual, rev.budget))}, operating profit ${fmtPct(pctChange(op.actual, op.budget))} and profit after tax ${fmtPct(pctChange(pat.actual, pat.budget))} vs budget.`,
          `予算比で売上高${fmtPct(pctChange(rev.actual, rev.budget))}、営業利益${fmtPct(pctChange(op.actual, op.budget))}、当期純利益${fmtPct(pctChange(pat.actual, pat.budget))}。`) },
        { title: t('Against prior year', '前年同期との比較'), text: t(`Revenue ${fmtPct(revGrowth)} and profit after tax ${fmtPct(pctChange(pat.actual, pat.prior))} vs ${twPriorPeriod}.`,
          `${twPriorPeriod}比で売上高${fmtPct(revGrowth)}、当期純利益${fmtPct(pctChange(pat.actual, pat.prior))}。`) },
        { title: t('Costs and margin', '費用と利益率'), text: t(`Operating expenses ${fmtAmt(opex.actual)} (${fmtPct(pctChange(opex.actual, opex.budget))} vs budget)${top && opex.actual ? `; ${top.name} is ${Math.round((top.actual / opex.actual) * 100)}% of the total` : ''}.`,
          `営業費用${fmtAmt(opex.actual)}（予算比${fmtPct(pctChange(opex.actual, opex.budget))}）${top && opex.actual ? `。${acct(top.name)}が営業費用の${Math.round((top.actual / opex.actual) * 100)}%を占める` : ''}。`) },
      ],
      watch: [
        ...items
          .map(it => ({ ...it, vb: pctChange(it.actual, it.budget) }))
          .filter(it => it.vb !== null && it.vb >= 10 && opex.actual && it.actual / opex.actual >= 0.02)
          .sort((a, b) => (b.actual - b.budget) - (a.actual - a.budget))
          .map(it => t(`${it.name} is ${fmtPct(it.vb)} over budget (${fmtAmt(it.actual)} vs ${fmtAmt(it.budget)}).`, `${acct(it.name)}が予算を${fmtPct(it.vb)}超過（実績${fmtAmt(it.actual)}、予算${fmtAmt(it.budget)}）。`)),
        ...fastCosts.map(it => t(`${it.name} grew ${fmtPct(it.g)} vs prior year, faster than revenue (${fmtPct(revGrowth)}).`, `${acct(it.name)}は前年同期比${fmtPct(it.g)}と、売上高の伸び（${fmtPct(revGrowth)}）を上回る。`)),
      ].slice(0, 3),
      computed: true,
    };
  }

  function addExecutiveSummarySlide(pptx) {
    const ai = executiveSummary && Array.isArray(executiveSummary.points) && executiveSummary.points.length ? executiveSummary : null;
    const s = ai || computedSummary();
    const slide = pptx.addSlide();
    addSlideTitle(slide, t('Executive summary', 'エグゼクティブサマリー'));
    slide.addText(s.headline || '', {
      x: CONFIG.layout.marginLeft, y: 0.6, w: 9, h: 0.6,
      fontSize: V2.headline, bold: true, color: CONFIG.colors.primaryDark, fontFace: CONFIG.fonts.primary, valign: 'top',
    });

    // KPI tiles: one shape + one multi-run text box each
    const tiles = [[t('Revenue', '売上高'), twLine('revenue')], [t('Operating profit', '営業利益'), twLine('operating_profit')], [t('Profit after tax', '当期純利益'), twLine('profit_after_tax')]];
    tiles.forEach(([label, l], i) => {
      const x = CONFIG.layout.marginLeft + i * 3.05;
      const vb = pctChange(l.actual, l.budget), vp = pctChange(l.actual, l.prior);
      slide.addShape(pptx.ShapeType.rect, { x, y: 1.3, w: 2.9, h: 0.95, fill: { color: CONFIG.colors.gray50 }, line: { color: 'D1D5DB', width: 1 } });
      slide.addText([
        { text: t(`${label} (${twPeriod})`, `${label}（${twPeriod}）`), options: { fontSize: V2.small, color: CONFIG.colors.gray500, breakLine: true } },
        { text: fmtAmt(l.actual), options: { fontSize: V2.kpi, bold: true, color: CONFIG.colors.gray900, breakLine: true } },
        { text: `${t('vs budget', '予算比')} ${fmtPct(vb)}`, options: { fontSize: V2.small, bold: true, color: varColor(vb, false) } },
        { text: '   ', options: { fontSize: V2.small } },
        { text: `${t('vs prior year', '前年同期比')} ${fmtPct(vp)}`, options: { fontSize: V2.small, bold: true, color: varColor(vp, false) } },
      ], { x: x + 0.12, y: 1.32, w: 2.7, h: 0.9, fontFace: CONFIG.fonts.primary, valign: 'middle', margin: 0 });
    });

    // Key messages (left) and watch items (right)
    s.points.slice(0, 3).forEach((p, i) => {
      slide.addText([
        { text: `${i + 1}. ${p.title}`, options: { fontSize: V2.point, bold: true, color: CONFIG.colors.gray900, breakLine: true } },
        { text: p.text, options: { fontSize: V2.body, color: CONFIG.colors.gray700 } },
      ], { x: CONFIG.layout.marginLeft, y: 2.45 + i * 0.82, w: 5.9, h: 0.78, fontFace: CONFIG.fonts.primary, valign: 'top', margin: 0 });
    });
    const watch = (s.watch || []).slice(0, 3);
    slide.addShape(pptx.ShapeType.rect, { x: 6.65, y: 2.45, w: 2.85, h: 2.5, fill: { color: CONFIG.colors.takeawayBg }, line: { type: 'none' } });
    slide.addText([
      { text: t('Watch items', '注視事項'), options: { fontSize: V2.point, bold: true, color: CONFIG.colors.gray900, breakLine: true } },
      ...(watch.length
        ? watch.map((w, i) => ({ text: w, options: { fontSize: V2.body, color: CONFIG.colors.gray700, bullet: { indent: 12 }, breakLine: i < watch.length - 1, paraSpaceBefore: 6 } }))
        : [{ text: t('No items flagged.', '該当なし'), options: { fontSize: V2.body, color: CONFIG.colors.gray500 } }]),
    ], { x: 6.8, y: 2.55, w: 2.6, h: 2.3, fontFace: CONFIG.fonts.primary, valign: 'top', margin: 0 });

    slide.addText(ai ? t('AI-drafted from the reported figures. Review before distribution.', 'AIが報告数値をもとに作成。配布前に内容を確認してください。')
      : t('Computed from the reported figures (AI summary unavailable for this run).', '報告数値から自動算出（今回はAIサマリーを利用できませんでした）。'), {
      x: CONFIG.layout.marginLeft, y: 5.15, w: 9, h: 0.3, fontSize: V2.small, color: CONFIG.colors.gray500, fontFace: CONFIG.fonts.primary,
    });
  }

  function addThreeWaySlide(pptx) {
    const slide = pptx.addSlide();
    addSlideTitle(slide, t('P&L: actual vs budget vs prior year', '損益：実績・予算・前年同期の比較'));
    addSlideSubtitle(slide, t(`${twPeriod} actual against budget and against ${twPriorPeriod}. Green = favourable, red = unfavourable (lower costs count as favourable).`,
      `${twPeriod}の実績を予算および${twPriorPeriod}と比較。緑＝有利、赤＝不利（費用は減少を有利とする）。`));

    const head = JA ? ['', '実績', '予算', '予算差', '予算比', '前年同期', '前年比'] : ['', 'Actual', 'Budget', 'vs Bud', 'vs Bud %', 'Prior Year', 'vs PY %'];
    const rows = [head.map((h, i) => ({ text: h, options: { bold: true, fontSize: V2.small, color: CONFIG.colors.white, fill: CONFIG.colors.tableHeaderBg, align: i ? 'right' : 'left' } }))];
    TW.lines.forEach(l => {
      const vb = pctChange(l.actual, l.budget), vp = pctChange(l.actual, l.prior);
      const fill = l.subtotal ? CONFIG.colors.gray100 : CONFIG.colors.white;
      const base = { fontSize: V2.small, fill, bold: !!l.subtotal, align: 'right' };
      rows.push([
        { text: twLabel(l), options: { ...base, align: 'left', color: CONFIG.colors.gray800 } },
        { text: fmtAmt(l.actual), options: { ...base, color: CONFIG.colors.gray900 } },
        { text: fmtAmt(l.budget), options: { ...base, color: CONFIG.colors.gray500 } },
        { text: isNum(l.budget) ? fmtDiff(l.actual - l.budget) : '–', options: { ...base, color: varColor(vb, l.cost) } },
        { text: fmtPct(vb), options: { ...base, color: varColor(vb, l.cost) } },
        { text: fmtAmt(l.prior), options: { ...base, color: CONFIG.colors.gray500 } },
        { text: fmtPct(vp), options: { ...base, color: varColor(vp, l.cost) } },
      ]);
    });
    slide.addTable(rows, {
      x: CONFIG.layout.marginLeft, y: 1.05, w: 6.0, colW: [1.36, 0.78, 0.78, 0.72, 0.76, 0.82, 0.78], rowH: 0.3,
      border: { pt: CONFIG.tables.borderWidth, color: CONFIG.tables.borderColor },
      margin: [0.02, 0.06, 0.02, 0.06], fontFace: CONFIG.fonts.primary, valign: 'middle', autoPage: false,
    });

    const shops = TW.revenue_items || [];
    slide.addChart(pptx.ChartType.bar, [
      { name: t('Actual', '実績'), labels: shops.map(s => s.name), values: shops.map(s => s.actual || 0) },
      { name: t('Budget', '予算'), labels: shops.map(s => s.name), values: shops.map(s => s.budget || 0) },
      { name: t('Prior Year', '前年同期'), labels: shops.map(s => s.name), values: shops.map(s => s.prior || 0) },
    ], {
      x: 6.65, y: 1.0, w: 3.0, h: 3.45, barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 60,
      chartColors: [CONFIG.colors.primaryDark, '9CA3AF', 'D1D5DB'],
      showTitle: true, title: t('Revenue by shop', '店舗別売上高'), titleFontSize: V2.small, titleColor: CONFIG.colors.gray700, titleFontFace: CONFIG.fonts.primary,
      showLegend: true, legendPos: 'b', legendFontSize: V2.small, legendFontFace: CONFIG.fonts.primary,
      catAxisLabelFontSize: V2.small, catAxisLabelFontFace: CONFIG.fonts.primary,
      valAxisLabelFontSize: V2.small, valAxisLabelFormatCode: '#,##0,"k"',
      valGridLine: { style: 'solid', color: CONFIG.charts.gridColor, size: 1 },
    });

    const rev = twLine('revenue'), pat = twLine('profit_after_tax');
    const best = [...shops].map(s => ({ ...s, vb: pctChange(s.actual, s.budget) })).filter(s => s.vb !== null).sort((a, b) => b.vb - a.vb)[0];
    addV2Takeaways(slide, pptx, [
      t(`Revenue ${fmtPct(pctChange(rev.actual, rev.budget))} vs budget, ${fmtPct(pctChange(rev.actual, rev.prior))} vs prior year.`,
        `売上高：予算比${fmtPct(pctChange(rev.actual, rev.budget))}、前年同期比${fmtPct(pctChange(rev.actual, rev.prior))}`),
      t(`Profit after tax ${fmtPct(pctChange(pat.actual, pat.budget))} vs budget.`, `当期純利益：予算比${fmtPct(pctChange(pat.actual, pat.budget))}`),
      best ? t(`${best.name}: largest beat vs budget (${fmtPct(best.vb)}).`, `${best.name}：予算超過幅が最大（${fmtPct(best.vb)}）`) : t('No shop-level budget available.', '店舗別予算なし'),
    ]);
  }

  function addOpexDeepDiveSlide(pptx) {
    const slide = pptx.addSlide();
    addSlideTitle(slide, t('Operating expenses deep-dive', '営業費用の詳細分析'));
    addSlideSubtitle(slide, t(`Operating expense lines for ${twPeriod}, largest first (PY = ${twPriorPeriod}). Highlighted: 10% or more off budget and at least 2% of total OPEX.`,
      `${twPeriod}の営業費用（金額順、前年＝${twPriorPeriod}）。黄色：予算差10%以上かつ営業費用合計の2%以上の科目。`));

    const opex = twLine('operating_expenses'), rev = twLine('revenue');
    const items = [...(TW.opex_items || [])].sort((a, b) => b.actual - a.actual);
    const head = JA ? ['', '実績', '予算', '予算比', '前年', '前年比', '売上比'] : ['', 'Actual', 'Budget', 'vs Bud', 'PY', 'vs PY', '% rev'];
    const rows = [head.map((h, i) => ({ text: h, options: { bold: true, fontSize: V2.small, color: CONFIG.colors.white, fill: CONFIG.colors.tableHeaderBg, align: i ? 'right' : 'left' } }))];
    const row = (name, l, opts = {}) => {
      const vb = pctChange(l.actual, l.budget), vp = pctChange(l.actual, l.prior);
      const flagged = !opts.total && vb !== null && Math.abs(vb) >= 10 && opex.actual && l.actual / opex.actual >= 0.02;
      const fill = opts.total ? CONFIG.colors.gray100 : (flagged ? 'FEF3C7' : CONFIG.colors.white);
      const base = { fontSize: V2.small, fill, bold: !!opts.total || flagged, align: 'right' };
      return [
        { text: name, options: { ...base, align: 'left', color: CONFIG.colors.gray800 } },
        { text: fmtAmt(l.actual), options: { ...base, color: CONFIG.colors.gray900 } },
        { text: fmtAmt(l.budget), options: { ...base, color: CONFIG.colors.gray500 } },
        { text: fmtPct(vb), options: { ...base, color: varColor(vb, true) } },
        { text: fmtAmt(l.prior), options: { ...base, color: CONFIG.colors.gray500 } },
        { text: fmtPct(vp), options: { ...base, color: varColor(vp, true) } },
        { text: rev.actual ? `${((l.actual / rev.actual) * 100).toFixed(1)}%` : '–', options: { ...base, color: CONFIG.colors.gray700 } },
      ];
    };
    // Largest 9 lines, the rest folded into one row so the table never runs into the takeaways
    const shown = items.length > 10 ? items.slice(0, 9) : items;
    const rest = items.slice(shown.length);
    shown.forEach(it => rows.push(row(acct(it.name), it)));
    if (rest.length) {
      const sum = (k) => (rest.every(it => isNum(it[k])) ? rest.reduce((t, it) => t + it[k], 0) : null);
      rows.push(row(t(`Other (${rest.length} lines)`, `その他（${rest.length}科目）`), { actual: sum('actual'), budget: sum('budget'), prior: sum('prior') }));
    }
    rows.push(row(t('Total operating expenses', '営業費用合計'), opex, { total: true }));
    const rowH = 0.26;
    slide.addTable(rows, {
      x: CONFIG.layout.marginLeft, y: 1.05, w: 6.06, colW: [2.1, 0.7, 0.7, 0.64, 0.7, 0.66, 0.56], rowH,
      border: { pt: CONFIG.tables.borderWidth, color: CONFIG.tables.borderColor },
      margin: [0.01, 0.06, 0.01, 0.06], fontFace: CONFIG.fonts.primary, valign: 'middle', autoPage: false,
    });

    const top = items.slice(0, 5).reverse();
    slide.addChart(pptx.ChartType.bar, [
      { name: t('Actual', '実績'), labels: top.map(it => acct(it.name)), values: top.map(it => it.actual || 0) },
      { name: t('Budget', '予算'), labels: top.map(it => acct(it.name)), values: top.map(it => it.budget || 0) },
    ], {
      x: 6.65, y: 1.0, w: 3.0, h: 3.45, barDir: 'bar', barGrouping: 'clustered', barGapWidthPct: 50,
      chartColors: [CONFIG.colors.primaryDark, '9CA3AF'],
      showTitle: true, title: t('Top 5 lines: actual vs budget', '上位5科目：実績と予算'), titleFontSize: V2.small, titleColor: CONFIG.colors.gray700, titleFontFace: CONFIG.fonts.primary,
      showLegend: true, legendPos: 'b', legendFontSize: V2.small, legendFontFace: CONFIG.fonts.primary,
      catAxisLabelFontSize: V2.small, catAxisLabelFontFace: CONFIG.fonts.primary,
      valAxisLabelFontSize: V2.small, valAxisLabelFormatCode: '#,##0,"k"',
      valGridLine: { style: 'solid', color: CONFIG.charts.gridColor, size: 1 },
    });

    const biggest = items[0];
    const ratio = (l) => (l.actual && rev.actual ? (l.actual / rev.actual) * 100 : null);
    const ratioPrior = opex.prior && rev.prior ? (opex.prior / rev.prior) * 100 : null;
    addV2Takeaways(slide, pptx, [
      biggest && opex.actual
        ? t(`${biggest.name}: ${Math.round((biggest.actual / opex.actual) * 100)}% of OPEX.`, `${acct(biggest.name)}：営業費用の${Math.round((biggest.actual / opex.actual) * 100)}%`)
        : t('No OPEX lines found.', '営業費用の科目なし'),
      t(`OPEX ${fmtPct(pctChange(opex.actual, opex.budget))} vs budget (${fmtDiff(isNum(opex.budget) ? opex.actual - opex.budget : null)}).`,
        `営業費用：予算比${fmtPct(pctChange(opex.actual, opex.budget))}（${fmtDiff(isNum(opex.budget) ? opex.actual - opex.budget : null)}）`),
      t(`OPEX ratio ${ratio(opex) === null ? '–' : ratio(opex).toFixed(1) + '%'} of revenue (prior year ${ratioPrior === null ? '–' : ratioPrior.toFixed(1) + '%'}).`,
        `営業費用率：売上高の${ratio(opex) === null ? '–' : ratio(opex).toFixed(1) + '%'}（前年同期${ratioPrior === null ? '–' : ratioPrior.toFixed(1) + '%'}）`),
    ]);
  }

  // ============================================================================
  // MAIN PRESENTATION GENERATION (Slides 2-4 continue...)
  // I'll include the complete generation code in the next section
  // ============================================================================

  try {
    const pptx = new PptxGenJS();
    pptx.layout = CONFIG.metadata.layout;
    pptx.author = CONFIG.metadata.author;
    pptx.company = inputData.company_name || '';
    pptx.title = `Financial Metrics - ${inputData.company_name || ''}`;

    // SLIDE 1: COVER
    await addCoverSlide(pptx, inputData, COVER_IMAGE_URL);
    if (TW) addExecutiveSummarySlide(pptx);

    // SLIDE 2: KEY FINANCIALS
    const slide2 = pptx.addSlide();
    let usedNativeWaterfall = null;
    addSlideTitle(slide2, CONFIG.text.slide2Title);
    
    const currentLabel = extractPeriodLabel(inputData.current_period_label);
    const comparisonLabel = extractPeriodLabel(inputData.comparison_period_label);
    addSlideSubtitle(slide2, t(`Income statement and P&L bridge, ${currentLabel} vs ${comparisonLabel}.`, `損益計算書と利益ブリッジ（${currentLabel} 対 ${comparisonLabel}）`));

    const chartImageDataUri = CHART_IMAGE_URL ? await downloadImageAsDataUri(CHART_IMAGE_URL) : null;
    const waterfall = chartImageDataUri ? null : buildWaterfallSeries(inputData, [
      { key: 'current', label: extractPeriodLabel(inputData.current_period_label) || t('Current', '当期') },
      { key: 'comparison', label: extractPeriodLabel(inputData.comparison_period_label) || t('Prior Year', '前年同期') },
    ], JA ? ['売上高', '売上原価', '売上総利益', '営業費用', 'その他（純額）', '当期純利益'] : undefined);

    if (waterfall && waterfall.valid) {
      const currencyLabel = pptxConfig?.currency_label || 'JPY';
      const labels = [waterfall.inner, waterfall.outer];
      slide2.addChart(pptx.ChartType.bar, [
        { name: WATERFALL_SPACER, labels, values: waterfall.spacer },
        { name: 'Total', labels, values: waterfall.total },
        { name: 'Increase', labels, values: waterfall.increase },
        { name: 'Decrease', labels, values: waterfall.decrease },
      ], {
        x: CONFIG.slide2.graphX,
        y: CONFIG.layout.contentStartY,
        w: CONFIG.slide2.graphW,
        h: CONFIG.slide2.graphH,
        barDir: 'bar',
        barGrouping: 'stacked',
        barGapWidthPct: 30,
        catAxisOrientation: 'maxMin',
        chartColors: ['FFFFFF', CONFIG.colors.primaryDark, '6EE7B7', 'A7F3D0'],
        catAxisMultiLevelLabels: true,
        catAxisLabelFontSize: CONFIG.fontSizes.chartLabel,
        catAxisLabelColor: CONFIG.colors.gray700,
        catAxisLabelFontFace: CONFIG.fonts.primary,
        valAxisHidden: true,
        valGridLine: { style: 'none' },
        showLegend: false,
        showValue: true,
        dataLabelPosition: 'inEnd',
        dataLabelFontSize: CONFIG.fontSizes.chartLabel,
        dataLabelFontFace: CONFIG.fonts.primary,
        dataLabelFormatCode: '#,##0,',
        dataLabelColor: '1F2937',
        showTitle: true,
        title: t(`P&L bridge (${currencyLabel} thousands)`, `利益ブリッジ（単位：千${currencyLabel === 'JPY' ? '円' : currencyLabel}）`),
        titleFontSize: CONFIG.fontSizes.chartLabel,
        titleColor: CONFIG.colors.gray500,
        titleFontFace: CONFIG.fonts.primary,
      });
      usedNativeWaterfall = waterfall;
    } else if (chartImageDataUri) {
      slide2.addImage({
        data: chartImageDataUri,
        x: CONFIG.slide2.graphX,
        y: CONFIG.layout.contentStartY,
        w: CONFIG.slide2.graphW,
        h: CONFIG.slide2.graphH
      });
    } else if (chartData && chartData.current && chartData.current.values) {
      const cd = chartData;
      const catLabels = ['Revenue', 'COGS', 'Gross Profit', 'OPEX', 'Net Income'];
      const catKeys = ['revenue', 'cogs', 'gross_profit', 'operating_expenses', 'net_income'];
      const seriesVals = (ser) => catKeys.map(k => Number(ser?.values?.[k]) || 0);
      slide2.addChart(pptx.ChartType.bar, [
        { name: String(cd.current.label || 'Actual'), labels: catLabels, values: seriesVals(cd.current) },
        { name: String((cd.comparison && cd.comparison.label) || 'Budget'), labels: catLabels, values: seriesVals(cd.comparison) },
      ], {
        x: CONFIG.slide2.graphX,
        y: CONFIG.layout.contentStartY,
        w: CONFIG.slide2.graphW,
        h: CONFIG.slide2.graphH,
        barDir: 'col',
        barGrouping: 'clustered',
        chartColors: [CONFIG.colors.chartBar1, CONFIG.colors.primary],
        showLegend: true,
        legendPos: 't',
        legendFontSize: CONFIG.fontSizes.chartLegend,
        valGridLine: { style: CONFIG.charts.gridStyle, color: CONFIG.charts.gridColor, size: 1 },
        catAxisLabelFontSize: CONFIG.fontSizes.small,
        valAxisLabelFontSize: CONFIG.fontSizes.small,
        showTitle: false,
      });
    } else {
      slide2.addText('Graph Placeholder\nVisualization of Financial Trends', {
        x: CONFIG.slide2.graphX,
        y: CONFIG.layout.contentStartY,
        w: CONFIG.slide2.graphW,
        h: CONFIG.slide2.graphH,
        fontSize: CONFIG.fontSizes.table,
        color: CONFIG.colors.gray500,
        align: 'center',
        valign: 'middle',
        fill: { color: CONFIG.colors.gray50 },
        line: { color: 'D1D5DB', width: 2, dashType: 'dash' },
        fontFace: CONFIG.fonts.primary
      });
    }

    
// Process operating expenses (detailed breakdown)
const opexItems = Array.isArray(inputData?.operating_expenses?.items) ? inputData.operating_expenses.items : [];

const sumItems = (items) => ({
  current: items.reduce((s, it) => s + (Number(it.current) || 0), 0),
  comparison: items.reduce((s, it) => s + (Number(it.comparison) || 0), 0),
});

const pickBy = (pred) => opexItems.filter(it => pred(String(it?.name || '').toLowerCase()));
const remainingExcluding = (picked) => {
  const pickedSet = new Set(picked);
  return opexItems.filter(it => !pickedSet.has(it));
};

// 1) Employee Benefits (salaries/wages/provident/welfare/employee benefits)
const employeeBenefitsItems = pickBy(n =>
  n.includes('employee') ||
  n.includes('benefit') ||
  n.includes('salary') ||
  n.includes('salaries') ||
  n.includes('wage') ||
  n.includes('wages') ||
  n.includes('provident') ||
  n.includes('welfare')
);

// 2) Advertisement & Publicity (advertisement/publicity/promo/marketing)
const advertisementItems = pickBy(n =>
  n.includes('advert') ||
  n.includes('publicity') ||
  n.includes('promotion') ||
  n.includes('promo') ||
  n.includes('marketing')
);

// 3) Legal & Professional Charges
const legalProfItems = pickBy(n =>
  n.includes('legal') ||
  n.includes('professional') ||
  n.includes('consult') ||
  n.includes('advis') ||
  n.includes('audit')
);

// 4) Depreciation & Amortization
const depAmortItems = pickBy(n =>
  n.includes('depreciation') ||
  n.includes('amortization') ||
  n.includes('amortisation')
);

// Remaining -> Other
const pickedAll = [...employeeBenefitsItems, ...advertisementItems, ...legalProfItems, ...depAmortItems];
const otherItems = remainingExcluding(pickedAll);

const employeeBenefitsTotal = { name: t('Employee Benefits Expenses', '人件費'), ...sumItems(employeeBenefitsItems) };
const advertisementTotal    = { name: t('Advertisement and Publicity', '広告宣伝費'), ...sumItems(advertisementItems) };
const legalProfTotal        = { name: t('Legal and Professional Charges', '支払報酬'), ...sumItems(legalProfItems) };
const depAmortTotal         = { name: t('Depreciation and Amortization', '減価償却費'), ...sumItems(depAmortItems) };
const otherExpensesTotal    = { name: t('Other', 'その他'), ...sumItems(otherItems) };

// Build in the exact order you expect (skip empty lines unless you want them shown)
const expenseBreakdown = [];
const pushIfAny = (o) => { if ((o.current || 0) !== 0 || (o.comparison || 0) !== 0) expenseBreakdown.push(o); };

pushIfAny(employeeBenefitsTotal);
pushIfAny(advertisementTotal);
pushIfAny(legalProfTotal);
pushIfAny(depAmortTotal);
// "Other" should be shown if there is any remainder OR if you want it always visible
if ((otherExpensesTotal.current || 0) !== 0 || (otherExpensesTotal.comparison || 0) !== 0) {
  expenseBreakdown.push(otherExpensesTotal);
}
const currentPeriodShort = extractPeriodLabel(inputData.current_period_label);
    const comparisonPeriodShort = extractPeriodLabel(inputData.comparison_period_label);

    // Build table data (continuing with all rows from your original code...)
    const tableData = [];
    tableData.push(createTableHeader([t('Metric', '項目'), currentPeriodShort, comparisonPeriodShort, t('Var %', '増減率')]));

    // Revenue
    tableData.push([
      { text: t('Revenue', '売上高'), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray800, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.revenue.current), options: { fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.revenue.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.revenue.current, inputData.revenue.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, bold: true, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // COGS
    tableData.push([
      { text: t('Cost of Goods Sold', '売上原価'), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray700, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.cogs.current), options: { fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.cogs.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.cogs.current, inputData.cogs.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, bold: true, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Gross Profit
    tableData.push([
      { text: t('Gross Profit', '売上総利益'), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray900, fill: CONFIG.colors.highlight, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.gross_profit.current), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray900, fill: CONFIG.colors.highlight, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.gross_profit.comparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray700, fill: CONFIG.colors.highlight, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.gross_profit.current, inputData.gross_profit.comparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.primaryDark, fill: CONFIG.colors.highlight, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Operating Expenses header
    tableData.push([
      { text: t('Operating Expenses', '営業費用'), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray700, fill: CONFIG.colors.gray50, align: 'left', colspan: 4, fontFace: CONFIG.fonts.primary } },
      { text: '', options: { fill: CONFIG.colors.gray50 } },
      { text: '', options: { fill: CONFIG.colors.gray50 } },
      { text: '', options: { fill: CONFIG.colors.gray50 } }
    ]);

    // Expense items
    for (const expense of expenseBreakdown) {
      tableData.push([
        { text: '  ' + expense.name, options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray700, align: 'left', fontFace: CONFIG.fonts.primary } },
        { text: formatCurrency(expense.current), options: { fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
        { text: formatCurrency(expense.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
        { text: calculateVariance(expense.current, expense.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, align: 'right', fontFace: CONFIG.fonts.primary } }
      ]);
    }

    // Total Operating Expenses
    tableData.push([
      { text: t('Total Operating Expenses', '営業費用合計'), options: { fontSize: CONFIG.fontSizes.table, bold: true, color: CONFIG.colors.gray800, fill: CONFIG.colors.gray100, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.operating_expenses.current), options: { fontSize: CONFIG.fontSizes.table, bold: true, fill: CONFIG.colors.gray100, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.operating_expenses.comparison), options: { fontSize: CONFIG.fontSizes.table, bold: true, color: CONFIG.colors.gray500, fill: CONFIG.colors.gray100, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.operating_expenses.current, inputData.operating_expenses.comparison), options: { fontSize: CONFIG.fontSizes.table, bold: true, color: CONFIG.colors.positive, fill: CONFIG.colors.gray100, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Profit from Operations
    const profitFromOps = inputData.gross_profit.current - inputData.operating_expenses.current;
    const profitFromOpsComparison = inputData.gross_profit.comparison - inputData.operating_expenses.comparison;
    tableData.push([
      { text: t('Profit from Operations', '営業利益'), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray800, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(profitFromOps), options: { bold: true, fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(profitFromOpsComparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(profitFromOps, profitFromOpsComparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Other Income
    tableData.push([
      { text: t('Other Income', 'その他収益'), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray700, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.other_income.current), options: { fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.other_income.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.other_income.current, inputData.other_income.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Profit Before Finance Costs
    const profitBeforeFinance = profitFromOps + inputData.other_income.current;
    const profitBeforeFinanceComparison = profitFromOpsComparison + inputData.other_income.comparison;
    tableData.push([
      { text: t('Profit Before Finance Costs', '金融費用控除前利益'), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray800, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(profitBeforeFinance), options: { bold: true, fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(profitBeforeFinanceComparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(profitBeforeFinance, profitBeforeFinanceComparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Finance Costs
    tableData.push([
      { text: t('Finance Costs', '金融費用'), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray700, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.finance_costs.current), options: { fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.finance_costs.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.finance_costs.current, inputData.finance_costs.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Profit Before Tax
    tableData.push([
      { text: t('Profit Before Tax', '税引前利益'), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray800, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.profit_before_tax.current), options: { bold: true, fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.profit_before_tax.comparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.profit_before_tax.current, inputData.profit_before_tax.comparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Tax Expense
    tableData.push([
      { text: t('Tax Expense', '法人税等'), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray700, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.tax_expense.current), options: { fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.tax_expense.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.tax_expense.current, inputData.tax_expense.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Profit After Tax
    tableData.push([
      { text: t('Profit After Tax', '当期純利益'), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.white, fill: CONFIG.colors.primary, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.profit_after_tax.current), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.white, fill: CONFIG.colors.primary, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.profit_after_tax.comparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.white, fill: CONFIG.colors.primary, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.profit_after_tax.current, inputData.profit_after_tax.comparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.white, fill: CONFIG.colors.primary, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Add table
    slide2.addTable(tableData, {
      x: CONFIG.slide2.tableX,
      y: CONFIG.layout.contentStartY,
      w: CONFIG.slide2.tableW,
      border: { pt: CONFIG.tables.borderWidth, color: CONFIG.tables.borderColor },
      colW: [CONFIG.slide2.tableColMetric, CONFIG.slide2.tableColValue, CONFIG.slide2.tableColValue, CONFIG.slide2.tableColValue],
      rowH: CONFIG.slide2.tableRowHeight,
      margin: [CONFIG.tables.cellPaddingTop, CONFIG.tables.cellPaddingSide, CONFIG.tables.cellPaddingTop, CONFIG.tables.cellPaddingSide],
      fontFace: CONFIG.fonts.primary,
      fontSize: CONFIG.fontSizes.table,
      valign: 'middle',
      autoPage: false
    });

    // Key Takeaways
    const revenueVariance = ((inputData.revenue.current - inputData.revenue.comparison) / inputData.revenue.comparison * 100).toFixed(0);
    const grossMarginCurrent = ((inputData.gross_profit.current / inputData.revenue.current) * 100).toFixed(0);
    const netProfitIncrease = (inputData.profit_after_tax.current - inputData.profit_after_tax.comparison) / 1000000;

    addKeyTakeaways(slide2, pptx, [
      CONFIG.text.takeawayTemplate1.replace('{var}', revenueVariance),
      CONFIG.text.takeawayTemplate2.replace('{var}', grossMarginCurrent),
      CONFIG.text.takeawayTemplate3.replace('{var}', netProfitIncrease.toFixed(1))
    ]);


    if (TW) addThreeWaySlide(pptx);

	    // ==================== SLIDE 3: DASHBOARD ====================
	    const slide3 = pptx.addSlide();
	    addSlideTitle(slide3, CONFIG.text.slide3Title);
	    addSlideSubtitle(slide3, t('Key margins, income statement and year-on-year performance for the period.', '主要利益率、損益計算書、前年同期との比較。'));

// Calculate metrics
	    const safeRevenue = Number(inputData.revenue.current) || 0;
	    const safeGrossProfit = Number(inputData.gross_profit.current) || 0;
	    const safeOpex = Number(inputData.operating_expenses.current) || 0;
	    const safeProfitAfterTax = Number(inputData.profit_after_tax.current) || 0;

	    const grossProfitMargin = safeRevenue ? Math.round((safeGrossProfit / safeRevenue) * 100) : 0;
	    const opexRatio = safeRevenue ? Math.round((safeOpex / safeRevenue) * 100) : 0;
	    const operatingProfit = safeGrossProfit - safeOpex;
	    const operatingProfitMargin = safeRevenue ? Math.round((operatingProfit / safeRevenue) * 100) : 0;
	    const netProfitMargin = safeRevenue ? Math.round((safeProfitAfterTax / safeRevenue) * 100) : 0;

	    // Period labels from the data (this is a year-on-year comparison)
    const curLabel = extractPeriodLabel(inputData.current_period_label) || t('Current', '当期');
    const priorLabel = extractPeriodLabel(inputData.comparison_period_label) || t('Prior Year', '前年同期');

    // Gauges — pushed below the subtitle to avoid overlap
    const gaugeRowY = 1.3;
    addGaugeChart(slide3, pptx, { x: 0.5, y: gaugeRowY, value: grossProfitMargin, title: t('Gross Profit Margin', '売上総利益率'), color: CONFIG.colors.primary, periodLabel: curLabel });
    addGaugeChart(slide3, pptx, { x: 0.5 + CONFIG.slide3.gaugeSpacing, y: gaugeRowY, value: opexRatio, title: t('OPEX Ratio', '営業費用率'), color: CONFIG.colors.primaryDark, periodLabel: curLabel });
    addGaugeChart(slide3, pptx, { x: 0.5 + CONFIG.slide3.gaugeSpacing * 2, y: gaugeRowY, value: operatingProfitMargin, title: t('Operating Profit Margin', '営業利益率'), color: CONFIG.colors.primary, periodLabel: curLabel });
    addGaugeChart(slide3, pptx, { x: 0.5 + CONFIG.slide3.gaugeSpacing * 3, y: gaugeRowY, value: netProfitMargin, title: t('Net Profit Margin', '純利益率'), color: CONFIG.colors.primaryDark, periodLabel: curLabel });

    // ---- Income Statement table (right) — real figures, real OPEX items ----
    const opexItemsRaw = Array.isArray(inputData.operating_expenses.items) ? inputData.operating_expenses.items : [];
    const sortedOpex = opexItemsRaw.map(it => ({ name: acct(String(it.name || '').trim()), current: Number(it.current) || 0 }))
      .filter(it => it.name && it.current)
      .sort((a, b) => b.current - a.current);
    const topOpex = sortedOpex.slice(0, 4);
    const otherOpex = sortedOpex.slice(4).reduce((sum, it) => sum + it.current, 0);

    slide3.addText(curLabel, {
      x: CONFIG.slide3.tableX, y: 0.9, w: CONFIG.slide3.tableW, h: 0.25,
      fontSize: 12, bold: true, color: CONFIG.colors.gray800, fontFace: CONFIG.fonts.primary, align: 'center',
    });

    const tableData3 = [];
    tableData3.push([
      { text: t('Income Statement', '損益計算書'), options: { bold: true, fontSize: 10, fill: CONFIG.colors.gray800, color: CONFIG.colors.white, fontFace: CONFIG.fonts.primary } },
      { text: '', options: { fill: CONFIG.colors.gray800 } },
    ]);
    const isRows = [
      [t('Revenue', '売上高'), formatCurrency(inputData.revenue.current), false],
      [t('COGS', '売上原価'), formatCurrency(inputData.cogs.current), false],
      [t('Gross Profit', '売上総利益'), formatCurrency(inputData.gross_profit.current), true],
      [t('Operating Expenses', '営業費用'), formatCurrency(inputData.operating_expenses.current), true],
    ];
    topOpex.forEach(it => isRows.push(['  ' + it.name, formatCurrency(it.current), false]));
    if (otherOpex > 0) isRows.push([t('  Other', '  その他'), formatCurrency(otherOpex), false]);
    isRows.push([t('Other Income', 'その他収益'), formatCurrency(inputData.other_income.current), false]);
    isRows.push([t('Finance Costs', '金融費用'), formatCurrency(inputData.finance_costs.current), false]);
    isRows.push([t('Profit After Tax', '当期純利益'), formatCurrency(inputData.profit_after_tax.current), true]);

    isRows.forEach(([label, val, bold]) => {
      const bg = bold ? CONFIG.colors.gray100 : CONFIG.colors.white;
      tableData3.push([
        { text: label, options: { bold, fontSize: 8, fill: bg, color: CONFIG.colors.gray800, fontFace: CONFIG.fonts.primary, align: 'left' } },
        { text: val, options: { bold, fontSize: 8, fill: bg, color: CONFIG.colors.gray800, fontFace: CONFIG.fonts.primary, align: 'right' } },
      ]);
    });
    slide3.addTable(tableData3, {
      x: CONFIG.slide3.tableX, y: 1.2, w: CONFIG.slide3.tableW,
      colW: [1.6, 1.1], rowH: 0.255,
      border: { pt: CONFIG.tables.borderWidth, color: 'D1D5DB' },
      fontSize: 8, fontFace: CONFIG.fonts.primary,
    });

    // ---- Chart A: P&L current vs prior year (honest YoY, real values) ----
    slide3.addText(t('Performance vs prior year', '前年同期との比較'), {
      x: 0.5, y: 2.65, w: 3.3, h: 0.25, fontSize: 10, bold: true,
      color: CONFIG.colors.gray800, fontFace: CONFIG.fonts.primary,
    });
    slide3.addChart(pptx.ChartType.bar, [
      { name: curLabel, labels: [t('Revenue', '売上高'), t('Gross Profit', '売上総利益'), t('Net Profit', '当期純利益')],
        values: [safeRevenue, safeGrossProfit, safeProfitAfterTax] },
      { name: priorLabel, labels: [t('Revenue', '売上高'), t('Gross Profit', '売上総利益'), t('Net Profit', '当期純利益')],
        values: [Number(inputData.revenue.comparison) || 0, Number(inputData.gross_profit.comparison) || 0, Number(inputData.profit_after_tax.comparison) || 0] },
    ], {
      x: 0.4, y: 2.95, w: 3.35, h: 2.2,
      barDir: 'col', barGrouping: 'clustered',
      chartColors: [CONFIG.colors.primary, CONFIG.colors.gray500],
      showLegend: true, legendPos: 'b', legendFontSize: 8,
      valGridLine: { style: 'solid', color: CONFIG.charts.gridColor, size: 1 },
      catAxisLabelFontSize: 8, valAxisLabelFontSize: 8, valAxisLabelFormatCode: '#,##0,"k"',
    });

    // ---- Chart B: operating expense breakdown (real line items) ----
    slide3.addText(t('Operating expense breakdown', '営業費用の内訳'), {
      x: 4.0, y: 2.65, w: 3.05, h: 0.25, fontSize: 10, bold: true,
      color: CONFIG.colors.gray800, fontFace: CONFIG.fonts.primary,
    });
    const clip = (nm) => nm.length > 20 ? nm.slice(0, 18) + '\u2026' : nm;
    const opexBars = topOpex.slice().reverse().map(it => ({ name: clip(it.name), current: it.current }));
    if (otherOpex > 0) opexBars.unshift({ name: t('Other', 'その他'), current: otherOpex });
    slide3.addChart(pptx.ChartType.bar, [
      { name: curLabel, labels: opexBars.map(b => b.name), values: opexBars.map(b => b.current) },
    ], {
      x: 4.0, y: 2.95, w: 3.05, h: 2.2,
      barDir: 'bar',
      chartColors: [CONFIG.colors.primary],
      showLegend: false,
      valGridLine: { style: 'solid', color: CONFIG.charts.gridColor, size: 1 },
      catAxisLabelFontSize: 8, valAxisLabelFontSize: 8, valAxisLabelFormatCode: '#,##0,"k"',
    });

    if (TW) addOpexDeepDiveSlide(pptx);

    // ==================== SLIDE 4: BUDGET VS ACTUAL ====================
	    const slide4 = pptx.addSlide();
	    addSlideTitle(slide4, CONFIG.text.slide4Title);
	    addSlideSubtitle(slide4, t(`Revenue by shop against budget for ${extractPeriodLabel(inputData.current_period_label)}, with variance commentary.`, `${extractPeriodLabel(inputData.current_period_label)}の店舗別売上高と予算差異、および差異の説明。`));

slide4.addText(t('Amount in JPY', '単位：円'), {
	      x: CONFIG.layout.marginLeft,
	      y: 0.9,
	      w: 2,
	      h: 0.25,
	      fontSize: CONFIG.fontSizes.table,
	      bold: true,
	      color: CONFIG.colors.gray800,
	      fontFace: CONFIG.fonts.primary,
	    });

	    // Variance table
	    const budgetDataLocal = inputData.budget_data || {};
	    const tableData4 = [];

	    // Header
	    tableData4.push([
	      { text: t('Account No', '勘定科目番号'), options: { fontSize: 8, bold: true, fill: 'E4D4F4', fontFace: CONFIG.fonts.primary } },
	      { text: t('Shop', '店舗'), options: { fontSize: 8, bold: true, fill: 'E4D4F4', fontFace: CONFIG.fonts.primary } },
	      { text: t('Actual', '実績'), options: { fontSize: 8, bold: true, fill: 'E4D4F4', fontFace: CONFIG.fonts.primary, align: 'right' } },
	      { text: t('Budget', '予算'), options: { fontSize: 8, bold: true, fill: 'E4D4F4', fontFace: CONFIG.fonts.primary, align: 'right' } },
	      { text: t('Var', '差異'), options: { fontSize: 8, bold: true, fill: 'E4D4F4', fontFace: CONFIG.fonts.primary, align: 'right' } },
	      { text: t('Var%', '差異率'), options: { fontSize: 8, bold: true, fill: 'E4D4F4', fontFace: CONFIG.fonts.primary, align: 'right' } },
	      { text: t('Explanation', '説明'), options: { fontSize: 8, bold: true, fill: 'E4D4F4', fontFace: CONFIG.fonts.primary } },
	    ]);

	    // Revenue section header
	    tableData4.push([
	      { text: t('Revenue', '売上高'), options: { fontSize: 10, bold: true, fill: CONFIG.colors.primary, color: CONFIG.colors.white, fontFace: CONFIG.fonts.primary, colspan: 7 } },
	      { text: '', options: { fill: CONFIG.colors.primary } },
	      { text: '', options: { fill: CONFIG.colors.primary } },
	      { text: '', options: { fill: CONFIG.colors.primary } },
	      { text: '', options: { fill: CONFIG.colors.primary } },
	      { text: '', options: { fill: CONFIG.colors.primary } },
	      { text: '', options: { fill: CONFIG.colors.primary } },
	    ]);

	    // Map any LLM-provided explanations by shop/line name
	    const explByShop = {};
	    (Array.isArray(shopCommentary) ? shopCommentary : []).forEach((r) => {
	      if (r && r.shop) explByShop[String(r.shop).trim().toLowerCase()] = String(r.explanation || '').trim();
	    });

	    // Revenue line items (actual shops vs budget) with real commentary
	    const revenueItems = (inputData.revenue && inputData.revenue.items) ? inputData.revenue.items : [];
	    const budgetRevenueItems = (budgetDataLocal.revenue && budgetDataLocal.revenue.items) ? budgetDataLocal.revenue.items : [];
	    const glAccounts = ['GL 123001', 'GL 123002', 'GL 123003', 'GL 123004'];

	    revenueItems.forEach((item, idx) => {
	      const budgetItem = budgetRevenueItems[idx] || { current: 0 };
	      const actual = Number(item.current) || 0;
	      const budget = Number(budgetItem.current) || 0;
	      const variance = actual - budget;
	      const variancePct = budget !== 0 ? ((variance / budget) * 100).toFixed(0) : '0';
	      const varColor = variance >= 0 ? CONFIG.colors.positive : CONFIG.colors.negative;
	      const varArrow = variance >= 0 ? '\u25b2' : '\u25bc';
	      const nm = String(item.name || '').trim();
	      const explanation = explByShop[nm.toLowerCase()] ||
	        (budget === 0 ? t('No budget baseline for this line.', '予算の設定なし。')
	          : (Math.abs(Number(variancePct)) < 10
	              ? t('Within \u00b110% of budget; no commentary required.', '予算比±10%以内のため説明不要。')
	              : t(nm + ' came in ' + (variance >= 0 ? 'above' : 'below') + ' budget by ' + Math.abs(Number(variancePct)) + '% (' + formatCurrency(actual) + ' vs ' + formatCurrency(budget) + ').',
	                  `${nm}は予算を${Math.abs(Number(variancePct))}%${variance >= 0 ? '上回った' : '下回った'}（実績${formatCurrency(actual)}、予算${formatCurrency(budget)}）。`)));
	      tableData4.push([
	        { text: glAccounts[idx] || '', options: { fontSize: 8, fontFace: CONFIG.fonts.primary } },
	        { text: nm, options: { fontSize: 8, fontFace: CONFIG.fonts.primary } },
	        { text: formatCurrency(actual), options: { fontSize: 8, fontFace: CONFIG.fonts.primary, align: 'right' } },
	        { text: formatCurrency(budget), options: { fontSize: 8, fontFace: CONFIG.fonts.primary, align: 'right' } },
	        { text: formatCurrency(variance), options: { fontSize: 8, fontFace: CONFIG.fonts.primary, align: 'right', color: varColor } },
	        { text: varArrow + ' ' + variancePct + '%', options: { fontSize: 8, fontFace: CONFIG.fonts.primary, align: 'right', color: varColor } },
	        { text: explanation, options: { fontSize: 8, fontFace: CONFIG.fonts.primary, color: CONFIG.colors.gray500 } },
	      ]);
	    });

	    // Total Revenue row
	    const budgetRevTotal = (budgetDataLocal.revenue && typeof budgetDataLocal.revenue.current === 'number') ? budgetDataLocal.revenue.current : (Number(budgetDataLocal.revenue?.current) || 0);
	    const revVar = (Number(inputData.revenue.current) || 0) - (Number(budgetRevTotal) || 0);
	    const revVarPct = (Number(budgetRevTotal) || 0) !== 0 ? ((revVar / Number(budgetRevTotal)) * 100).toFixed(0) : '0';
	    const revVarColor = revVar >= 0 ? CONFIG.colors.positive : CONFIG.colors.negative;
	    const revArrow = revVar >= 0 ? '▲' : '▼';

	    tableData4.push([
	      { text: t('Total Revenue', '売上高合計'), options: { fontSize: 8, bold: true, fill: CONFIG.colors.gray100, fontFace: CONFIG.fonts.primary, colspan: 2 } },
	      { text: formatCurrency(inputData.revenue.current), options: { fontSize: 8, bold: true, fill: CONFIG.colors.gray100, fontFace: CONFIG.fonts.primary, align: 'right' } },
	      { text: formatCurrency(budgetRevTotal), options: { fontSize: 8, bold: true, fill: CONFIG.colors.gray100, fontFace: CONFIG.fonts.primary, align: 'right' } },
	      { text: formatCurrency(revVar), options: { fontSize: 8, bold: true, fill: CONFIG.colors.gray100, fontFace: CONFIG.fonts.primary, align: 'right', color: revVarColor } },
	      { text: `${revArrow} ${revVarPct}%`, options: { fontSize: 8, bold: true, fill: CONFIG.colors.gray100, fontFace: CONFIG.fonts.primary, align: 'right', color: revVarColor } },
	      { text: '', options: { fill: CONFIG.colors.gray100 } },
	    ]);

	    slide4.addTable(tableData4, {
	      x: CONFIG.slide4.tableX,
	      y: CONFIG.slide4.tableY,
	      w: CONFIG.slide4.tableW,
	      colW: [0.98, 1.18, 0.98, 0.98, 0.98, 0.98, 2.95],
	      rowH: CONFIG.slide4.tableRowH,
	      border: { pt: CONFIG.tables.borderWidth, color: 'D1D5DB' },
	      fontSize: 8,
	      fontFace: CONFIG.fonts.primary,
	      valign: 'top',
	    });

	    // Budget comparison charts
	    const budgetRev = Number(budgetRevTotal) || 0;
	    const budgetGP = (budgetDataLocal.gross_profit && budgetDataLocal.gross_profit.current) ? Number(budgetDataLocal.gross_profit.current) : 0;
	    const budgetProfit = (budgetDataLocal.profit_after_tax && budgetDataLocal.profit_after_tax.current) ? Number(budgetDataLocal.profit_after_tax.current) : 0;

	    const revPct = budgetRev !== 0 ? Math.round(((Number(inputData.revenue.current) || 0) - budgetRev) / budgetRev * 100) : 0;
	    const gpPct = budgetGP !== 0 ? Math.round(((Number(inputData.gross_profit.current) || 0) - budgetGP) / budgetGP * 100) : 0;
	    const profitPct = budgetProfit !== 0 ? Math.round(((Number(inputData.profit_after_tax.current) || 0) - budgetProfit) / budgetProfit * 100) : 0;

	    addBudgetChart(slide4, pptx, {
	      x: 0.5,
	      title: t('Revenue vs Budget', '売上高（対予算）'),
	      subtitle: `$${Math.round((Number(inputData.revenue.current) || 0) / 1000000)}M`,
	      actual: Number(inputData.revenue.current) || 0,
	      budget: budgetRev,
	      varPct: revPct,
	    });

	    addBudgetChart(slide4, pptx, {
	      x: 3.83,
	      title: t('Gross Profit vs Budget', '売上総利益（対予算）'),
	      subtitle: `$${Math.round((Number(inputData.gross_profit.current) || 0) / 1000000)}M`,
	      actual: Number(inputData.gross_profit.current) || 0,
	      budget: budgetGP,
	      varPct: gpPct,
	    });

	    addBudgetChart(slide4, pptx, {
	      x: 7.15,
	      title: t('Profit After Tax vs Budget', '当期純利益（対予算）'),
	      subtitle: `$${Math.round((Number(inputData.profit_after_tax.current) || 0) / 1000000)}M`,
	      actual: Number(inputData.profit_after_tax.current) || 0,
	      budget: budgetProfit,
	      varPct: profitPct,
	    });

	    // Footer
	    slide4.addText(CONFIG.text.footer, {
	      x: 1.1,
	      y: CONFIG.layout.footerY,
	      w: 11.1,
	      h: 0.24,
	      fontSize: CONFIG.fontSizes.small,
	      color: CONFIG.colors.gray500,
	      fontFace: CONFIG.fonts.primary,
	    });

    let pptxBase64 = await pptx.write('base64');
    if (usedNativeWaterfall) pptxBase64 = await patchWaterfallCharts(pptxBase64, usedNativeWaterfall.deltaLabels, CONFIG.fonts.primary);
    pptxBase64 = await applyEastAsianFont(pptxBase64, JA);
    const filename = `${CONFIG.metadata.filenamePrefix}_${safeFilePart(inputData.company_name)}.pptx`;
    
    return { filename, pptx_base64: pptxBase64 };

  } catch (err) {
    console.error('PPTX Generation Error:', err);
    throw err;
  }
}

// ============================================================================
// EXPRESS API ENDPOINTS
// ============================================================================

app.post('/generate-pptx', async (req, res) => {
  try {
    const body = req.body || {};

    // Helpful server-side logging (kept short to avoid printing huge payloads)
    const topKeys = Object.keys(body);
    const fdCandidate =
      body.financialData ?? body.financial_data ?? body.fd_json ?? body.fdJson ?? body.data ?? null;
    console.log(`[REQ] /generate-pptx keys=${topKeys.join(',')} financialDataType=${typeof fdCandidate} len=${(typeof fdCandidate === 'string') ? fdCandidate.length : ''}`);

    // Dify workflow often sends financialData as a JSON string (fd_json).
    const normalizedFinancialData = normalizeFinancialData(fdCandidate);

    // pptxConfig: keep existing behavior (falls back to root body if not nested)
    const pptxConfig = body.pptxConfig || body.pptx_config || body.config || body || {};

    // Optional images
    const coverImageUrl = body.coverImageUrl || body.cover_image_url || null;
    const chartImageUrl = body.chartImageUrl || body.chart_image_url || null;
    const chartData = safeJsonParse(body.chart_data ?? body.chartData ?? null, 2);
    const shopCommentary = safeJsonParse(body.shop_commentary ?? body.shopCommentary ?? null, 2);
    const executiveSummary = safeJsonParse(body.executive_summary ?? body.executiveSummary ?? null, 2);

    const result = await generatePresentation({
      pptxConfig,
      financialData: normalizedFinancialData,
      coverImageUrl,
      chartImageUrl,
      chartData,
      shopCommentary,
      executiveSummary,
      language: body.language,
    });

    // Persist the pptx to disk so you can see it in the output folder
    if (!result || !result.pptx_base64 || !result.filename) {
      throw new Error('Generator returned empty result (missing filename or pptx_base64).');
    }

    const buf = Buffer.from(result.pptx_base64, 'base64');
    const publicBase = process.env.PUBLIC_BASE_URL
      || `${req.get('x-forwarded-proto') || req.protocol}://${req.get('x-forwarded-host') || req.get('host')}`;

    // ?link=1 — save under a random id and return an expiring download link
    // that works without the render token (used by the n8n workflow).
    const wantLink = req.query.link === '1' || req.query.link === 'true';
    if (wantLink) {
      const id = crypto.randomBytes(16).toString('hex');
      const dir = path.join(DOWNLOAD_DIR, id);
      fs.mkdirSync(dir, { recursive: true });
      const filename = path.basename(result.filename);
      fs.writeFileSync(path.join(dir, filename), buf);
      console.log(`[OK] wrote dl/${id}/${filename} (${buf.length} bytes)`);
      return res.json({
        success: true,
        filename,
        file_size_bytes: buf.length,
        download_url: `${publicBase.replace(/\/$/, '')}/d/${id}/${encodeURIComponent(filename)}`,
        expires_at: new Date(Date.now() + DOWNLOAD_TTL_MS).toISOString(),
      });
    }

    const outPath = path.join(OUTPUT_DIR, result.filename);
    fs.writeFileSync(outPath, buf);

    console.log(`[OK] wrote ${outPath} (${buf.length} bytes)`);

    // ?return_file=1 — stream the .pptx binary back directly. Dify's HTTP
    // Request node captures binary responses into its `files` output, so the
    // deck can be surfaced as a downloadable file on the Dify run page itself
    // (needed in environments where this server's domain is browser-blocked).
    const wantFile = req.query.return_file === '1' || req.query.return_file === 'true';
    if (wantFile) {
      const asciiName = result.filename.replace(/[^\x20-\x7e]/g, '_').replace(/"/g, '');
      res.set({
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `attachment; filename="${asciiName}"; filename*=UTF-8''${encodeURIComponent(result.filename)}`,
        'Content-Length': String(buf.length),
      });
      return res.send(buf);
    }

    // NOTE: do NOT return pptx_base64 in the response body — a full deck with
    // embedded images is several MB and exceeds Dify's 1 MB HTTP-node response
    // limit. The file is saved to disk and fetchable via download_url; callers
    // that want the bytes can pass ?include_base64=1.
    const wantBase64 = req.query.include_base64 === '1' || req.query.include_base64 === 'true';
    // Absolute download URL (behind Easypanel's proxy the public scheme/host
    // arrive in X-Forwarded-*). Include the token so the link works when
    // clicked as-is — the caller already holds the token to reach this route.
    const tokenQs = RENDER_TOKEN ? `?token=${encodeURIComponent(RENDER_TOKEN)}` : '';
    res.json({
      success: true,
      filename: result.filename,
      filepath: outPath,
      file_size_bytes: buf.length,
      download_url: `${publicBase}/download/${encodeURIComponent(result.filename)}${tokenQs}`,
      ...(wantBase64 ? { pptx_base64: result.pptx_base64 } : {}),
    });
  } catch (error) {
    console.error('Error generating PPTX:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get('/download/:filename', (req, res) => {
  const filepath = path.join(OUTPUT_DIR, path.basename(req.params.filename));
  if (fs.existsSync(filepath)) {
    res.download(filepath);
  } else {
    res.status(404).json({ error: 'File not found' });
  }
});

app.get('/d/:id/:filename', (req, res) => {
  if (!/^[a-f0-9]{32}$/.test(req.params.id)) return res.status(404).json({ error: 'File not found' });
  const dir = path.join(DOWNLOAD_DIR, req.params.id);
  const filepath = path.join(dir, path.basename(req.params.filename));
  if (!fs.existsSync(filepath) || Date.now() - fs.statSync(dir).mtimeMs > DOWNLOAD_TTL_MS) {
    return res.status(404).json({ error: 'File not found or link expired' });
  }
  res.download(filepath, path.basename(filepath));
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok', version: 'bilingual-3', timestamp: new Date().toISOString() });
});

// Web front end (config + run-workflow UI), served at the root.
app.get('/', (req, res) => {
  const page = path.join(__dirname, 'public', 'index.html');
  if (fs.existsSync(page)) return res.sendFile(page);
  res.status(404).send('frontend not bundled');
});

app.listen(PORT, () => {
  console.log(`
╔════════════════════════════════════════════════════════════╗
║   PPTX Generator API Server                                ║
║   Running on: http://localhost:${PORT}                        ║
║   Output directory: ${OUTPUT_DIR}                  ║
╚════════════════════════════════════════════════════════════╝
  `);
});

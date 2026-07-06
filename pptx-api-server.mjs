import express from 'express';
import cors from 'cors';
import PptxGenJS from 'pptxgenjs';
import axios from 'axios';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { Buffer } from 'buffer';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// Optional bearer-token auth. If RENDER_TOKEN is set (recommended when the
// server is publicly reachable, e.g. on Easypanel), /generate-pptx and
// /download require it — via `Authorization: Bearer <token>` or `?token=`.
// /health stays open. When RENDER_TOKEN is unset (local use), auth is off.
const RENDER_TOKEN = process.env.RENDER_TOKEN || '';
app.use((req, res, next) => {
  if (!RENDER_TOKEN || req.path === '/health') return next();
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
// PPTX GENERATION FUNCTION (From your n8n code)
// ============================================================================

async function generatePresentation(config) {
  const { pptxConfig, financialData, coverImageUrl, chartImageUrl, chartData, shopCommentary } = config;
  
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
      table: globalStyles.font_size_table || 9,
      small: globalStyles.font_size_small || 7,
      chartLabel: globalStyles.font_size_chart_label || 8,
      chartLegend: globalStyles.font_size_chart_legend || 7,
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
      slide2Title: textContent.text_slide2_title || 'Key monthly financials metrics with gross and net profit',
      slide3Title: textContent.text_slide3_title || 'Monthly financial metrics dashboard with income statement',
      slide4Title: textContent.text_slide4_title || 'Monthly financials metrics with revenue and gross profit',
      footer: textContent.text_footer || 'This graph/chart is linked to excel, and changes automatically based on data. Just left click on it and select "edit data"',
      takeawayTemplate1: textContent.text_takeaway_template_1 || 'Year-on-year revenue change of {var}%.',
      takeawayTemplate2: textContent.text_takeaway_template_2 || 'Gross profit margin of {var}% of revenue.',
      takeawayTemplate3: textContent.text_takeaway_template_3 || 'Year-on-year net profit change of {var}M (local currency).'
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

  function extractPeriodLabel(label) {
    const s = String(label ?? '');
    const match = s.match(/to\s+(\d{3})\/(\d{4})/);
    if (match) {
      const period = parseInt(match[1]);
      const year = match[2];
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      return `${months[period - 1]} ${year}`;
    }
    return s;
  }

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

    slide.addText('Key Takeaways', {
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
    
    slide.addText(options.periodLabel || 'Mar. 2025', {
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
      fontSize: 15, bold: true, color: CONFIG.colors.gray900, fontFace: CONFIG.fonts.primary, valign: 'middle',
    });
    slide.addText(`${varPct >= 0 ? '\u25b2' : '\u25bc'} ${Math.abs(varPct)}%`, {
      x: x + 1.75, y: 4.78, w: 0.95, h: 0.34,
      fontSize: 11, bold: true, align: 'right', valign: 'middle',
      color: varPct >= 0 ? CONFIG.colors.positive : CONFIG.colors.negative, fontFace: CONFIG.fonts.primary,
    });
    slide.addText(`vs Budget ${formatCurrency(budget)}`, {
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

    const periodLong = (() => {
      const label = String(inputData.current_period_label || '').trim();
      const m = label.match(/to\s*(\d{3})\/(\d{4})/i) || label.match(/(\d{3})\/(\d{4})/);
      if (!m) return '';
      const period = parseInt(m[1], 10);
      const year = m[2];
      const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
      if (!period || period < 1 || period > 12) return '';
      return `${months[period - 1]} ${year}`;
    })();

    slide.addImage({
      data: svgToDataUriCover(await buildCoverBackgroundSvg(coverImageUrl)),
      x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
    });

    slide.addText("Monthly\nFinancials\nMetrics", {
      x: sx(CONFIG.cover.titleX), y: sx(CONFIG.cover.titleY), w: sx(5.85), h: sx(3.40),
      fontFace: CONFIG.fonts.primary,
      fontSize: Math.round(CONFIG.cover.titleFontSize * S),
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

    slide.addText(inputData.company_name || "Your Company Name", {
      x: sx(CONFIG.cover.titleX), y: sx(5.78), w: sx(5.85), h: sx(0.40),
      fontFace: CONFIG.fonts.primary,
      fontSize: Math.round(18 * S),
      color: CONFIG.colors.white,
      align: "right",
      valign: "top",
    });

    if (periodLong) {
      slide.addText(periodLong, {
        x: sx(CONFIG.cover.titleX), y: sx(6.08), w: sx(5.85), h: sx(0.35),
        fontFace: CONFIG.fonts.primary,
        fontSize: Math.round(16 * S),
        color: CONFIG.colors.white,
        align: "right",
        valign: "top",
      });
    }
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

    // SLIDE 2: KEY FINANCIALS
    const slide2 = pptx.addSlide();
    addSlideTitle(slide2, CONFIG.text.slide2Title);
    
    const currentLabel = String(inputData.current_period_label || '').replace('Posting periods ', '').replace(' to ', '-');
    const comparisonLabel = String(inputData.comparison_period_label || '').replace('Posting periods ', '').replace(' to ', '-');
    addSlideSubtitle(slide2, `Financial overview comparing ${currentLabel} vs ${comparisonLabel}. This slide showcases income statement to analyze overall growth of business and plan expenses for upcoming periods.`);

    const chartImageDataUri = CHART_IMAGE_URL ? await downloadImageAsDataUri(CHART_IMAGE_URL) : null;
    
    if (chartImageDataUri) {
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

const employeeBenefitsTotal = { name: 'Employee Benefits Expenses', ...sumItems(employeeBenefitsItems) };
const advertisementTotal    = { name: 'Advertisement and Publicity', ...sumItems(advertisementItems) };
const legalProfTotal        = { name: 'Legal and Professional Charges', ...sumItems(legalProfItems) };
const depAmortTotal         = { name: 'Depreciation and Amortization', ...sumItems(depAmortItems) };
const otherExpensesTotal    = { name: 'Other', ...sumItems(otherItems) };

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
    tableData.push(createTableHeader(['Metric', currentPeriodShort, comparisonPeriodShort, 'Var %']));

    // Revenue
    tableData.push([
      { text: 'Revenue', options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray800, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.revenue.current), options: { fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.revenue.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.revenue.current, inputData.revenue.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, bold: true, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // COGS
    tableData.push([
      { text: 'Cost of Goods Sold', options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray700, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.cogs.current), options: { fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.cogs.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.cogs.current, inputData.cogs.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, bold: true, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Gross Profit
    tableData.push([
      { text: 'Gross Profit', options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray900, fill: CONFIG.colors.highlight, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.gross_profit.current), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray900, fill: CONFIG.colors.highlight, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.gross_profit.comparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray700, fill: CONFIG.colors.highlight, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.gross_profit.current, inputData.gross_profit.comparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.primaryDark, fill: CONFIG.colors.highlight, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Operating Expenses header
    tableData.push([
      { text: 'Operating Expenses', options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray700, fill: CONFIG.colors.gray50, align: 'left', colspan: 4, fontFace: CONFIG.fonts.primary } },
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
      { text: 'Total Operating Expenses', options: { fontSize: CONFIG.fontSizes.table, bold: true, color: CONFIG.colors.gray800, fill: CONFIG.colors.gray100, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.operating_expenses.current), options: { fontSize: CONFIG.fontSizes.table, bold: true, fill: CONFIG.colors.gray100, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.operating_expenses.comparison), options: { fontSize: CONFIG.fontSizes.table, bold: true, color: CONFIG.colors.gray500, fill: CONFIG.colors.gray100, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.operating_expenses.current, inputData.operating_expenses.comparison), options: { fontSize: CONFIG.fontSizes.table, bold: true, color: CONFIG.colors.positive, fill: CONFIG.colors.gray100, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Profit from Operations
    const profitFromOps = inputData.gross_profit.current - inputData.operating_expenses.current;
    const profitFromOpsComparison = inputData.gross_profit.comparison - inputData.operating_expenses.comparison;
    tableData.push([
      { text: 'Profit from Operations', options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray800, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(profitFromOps), options: { bold: true, fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(profitFromOpsComparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(profitFromOps, profitFromOpsComparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Other Income
    tableData.push([
      { text: 'Other Income', options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray700, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.other_income.current), options: { fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.other_income.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.other_income.current, inputData.other_income.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Profit Before Finance Costs
    const profitBeforeFinance = profitFromOps + inputData.other_income.current;
    const profitBeforeFinanceComparison = profitFromOpsComparison + inputData.other_income.comparison;
    tableData.push([
      { text: 'Profit Before Finance Costs', options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray800, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(profitBeforeFinance), options: { bold: true, fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(profitBeforeFinanceComparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(profitBeforeFinance, profitBeforeFinanceComparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Finance Costs
    tableData.push([
      { text: 'Finance Costs', options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray700, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.finance_costs.current), options: { fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.finance_costs.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.finance_costs.current, inputData.finance_costs.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Profit Before Tax
    tableData.push([
      { text: 'Profit Before Tax', options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray800, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.profit_before_tax.current), options: { bold: true, fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.profit_before_tax.comparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.profit_before_tax.current, inputData.profit_before_tax.comparison), options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Tax Expense
    tableData.push([
      { text: 'Tax Expense', options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray700, align: 'left', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.tax_expense.current), options: { fontSize: CONFIG.fontSizes.table, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: formatCurrency(inputData.tax_expense.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.gray500, align: 'right', fontFace: CONFIG.fonts.primary } },
      { text: calculateVariance(inputData.tax_expense.current, inputData.tax_expense.comparison), options: { fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.positive, align: 'right', fontFace: CONFIG.fonts.primary } }
    ]);

    // Profit After Tax
    tableData.push([
      { text: 'Profit After Tax', options: { bold: true, fontSize: CONFIG.fontSizes.table, color: CONFIG.colors.white, fill: CONFIG.colors.primary, align: 'left', fontFace: CONFIG.fonts.primary } },
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


	    // ==================== SLIDE 3: DASHBOARD ====================
	    const slide3 = pptx.addSlide();
	    addSlideTitle(slide3, CONFIG.text.slide3Title);
	    addSlideSubtitle(slide3, 'Key margins, income statement and year-on-year performance for the period.');

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
    const curLabel = extractPeriodLabel(inputData.current_period_label) || 'Current';
    const priorLabel = extractPeriodLabel(inputData.comparison_period_label) || 'Prior Year';

    // Gauges — pushed below the subtitle to avoid overlap
    const gaugeRowY = 1.3;
    addGaugeChart(slide3, pptx, { x: 0.5, y: gaugeRowY, value: grossProfitMargin, title: 'Gross Profit Margin', color: CONFIG.colors.primary, periodLabel: curLabel });
    addGaugeChart(slide3, pptx, { x: 0.5 + CONFIG.slide3.gaugeSpacing, y: gaugeRowY, value: opexRatio, title: 'OPEX Ratio', color: CONFIG.colors.primaryDark, periodLabel: curLabel });
    addGaugeChart(slide3, pptx, { x: 0.5 + CONFIG.slide3.gaugeSpacing * 2, y: gaugeRowY, value: operatingProfitMargin, title: 'Operating Profit Margin', color: CONFIG.colors.primary, periodLabel: curLabel });
    addGaugeChart(slide3, pptx, { x: 0.5 + CONFIG.slide3.gaugeSpacing * 3, y: gaugeRowY, value: netProfitMargin, title: 'Net Profit Margin', color: CONFIG.colors.primaryDark, periodLabel: curLabel });

    // ---- Income Statement table (right) — real figures, real OPEX items ----
    const opexItemsRaw = Array.isArray(inputData.operating_expenses.items) ? inputData.operating_expenses.items : [];
    const sortedOpex = opexItemsRaw.map(it => ({ name: String(it.name || '').trim(), current: Number(it.current) || 0 }))
      .filter(it => it.name && it.current)
      .sort((a, b) => b.current - a.current);
    const topOpex = sortedOpex.slice(0, 4);
    const otherOpex = sortedOpex.slice(4).reduce((sum, it) => sum + it.current, 0);

    slide3.addText(curLabel, {
      x: CONFIG.slide3.tableX, y: 0.9, w: CONFIG.slide3.tableW, h: 0.25,
      fontSize: 11, bold: true, color: CONFIG.colors.gray800, fontFace: CONFIG.fonts.primary, align: 'center',
    });

    const tableData3 = [];
    tableData3.push([
      { text: 'Income Statement', options: { bold: true, fontSize: 10, fill: CONFIG.colors.gray800, color: CONFIG.colors.white, fontFace: CONFIG.fonts.primary } },
      { text: '', options: { fill: CONFIG.colors.gray800 } },
    ]);
    const isRows = [
      ['Revenue', formatCurrency(inputData.revenue.current), false],
      ['COGS', formatCurrency(inputData.cogs.current), false],
      ['Gross Profit', formatCurrency(inputData.gross_profit.current), true],
      ['Operating Expenses', formatCurrency(inputData.operating_expenses.current), true],
    ];
    topOpex.forEach(it => isRows.push(['  ' + it.name, formatCurrency(it.current), false]));
    if (otherOpex > 0) isRows.push(['  Other', formatCurrency(otherOpex), false]);
    isRows.push(['Other Income', formatCurrency(inputData.other_income.current), false]);
    isRows.push(['Finance Costs', formatCurrency(inputData.finance_costs.current), false]);
    isRows.push(['Profit After Tax', formatCurrency(inputData.profit_after_tax.current), true]);

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
    slide3.addText('Performance vs prior year (' + curLabel + ' vs ' + priorLabel + ')', {
      x: 0.5, y: 2.65, w: 3.3, h: 0.25, fontSize: 10, bold: true,
      color: CONFIG.colors.gray800, fontFace: CONFIG.fonts.primary,
    });
    slide3.addChart(pptx.ChartType.bar, [
      { name: curLabel, labels: ['Revenue', 'Gross Profit', 'Net Profit'],
        values: [safeRevenue, safeGrossProfit, safeProfitAfterTax] },
      { name: priorLabel, labels: ['Revenue', 'Gross Profit', 'Net Profit'],
        values: [Number(inputData.revenue.comparison) || 0, Number(inputData.gross_profit.comparison) || 0, Number(inputData.profit_after_tax.comparison) || 0] },
    ], {
      x: 0.4, y: 2.95, w: 3.35, h: 2.2,
      barDir: 'col', barGrouping: 'clustered',
      chartColors: [CONFIG.colors.primary, CONFIG.colors.gray500],
      showLegend: true, legendPos: 'b', legendFontSize: 7,
      valGridLine: { style: 'solid', color: CONFIG.charts.gridColor, size: 1 },
      catAxisLabelFontSize: 8, valAxisLabelFontSize: 7,
    });

    // ---- Chart B: operating expense breakdown (real line items) ----
    slide3.addText('Operating expense breakdown (' + curLabel + ')', {
      x: 4.0, y: 2.65, w: 3.05, h: 0.25, fontSize: 10, bold: true,
      color: CONFIG.colors.gray800, fontFace: CONFIG.fonts.primary,
    });
    const clip = (nm) => nm.length > 20 ? nm.slice(0, 18) + '\u2026' : nm;
    const opexBars = topOpex.slice().reverse().map(it => ({ name: clip(it.name), current: it.current }));
    if (otherOpex > 0) opexBars.unshift({ name: 'Other', current: otherOpex });
    slide3.addChart(pptx.ChartType.bar, [
      { name: curLabel, labels: opexBars.map(b => b.name), values: opexBars.map(b => b.current) },
    ], {
      x: 4.0, y: 2.95, w: 3.05, h: 2.2,
      barDir: 'bar',
      chartColors: [CONFIG.colors.primary],
      showLegend: false,
      valGridLine: { style: 'solid', color: CONFIG.charts.gridColor, size: 1 },
      catAxisLabelFontSize: 7, valAxisLabelFontSize: 6,
    });

    // ==================== SLIDE 4: BUDGET VS ACTUAL ====================
	    const slide4 = pptx.addSlide();
	    addSlideTitle(slide4, CONFIG.text.slide4Title);
	    addSlideSubtitle(slide4, 'This slide displays P&L variance analysis to closely monitor key financial parameters of the company. It includes details about cost, discounts allowed, budget meeting, overtime finance charges, etc.');

slide4.addText('Amount in JPY', {
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
	      { text: 'Account No', options: { fontSize: 8, bold: true, fill: 'E4D4F4', fontFace: CONFIG.fonts.primary } },
	      { text: 'Shop', options: { fontSize: 8, bold: true, fill: 'E4D4F4', fontFace: CONFIG.fonts.primary } },
	      { text: 'Act MTh', options: { fontSize: 8, bold: true, fill: 'E4D4F4', fontFace: CONFIG.fonts.primary, align: 'right' } },
	      { text: 'Budget', options: { fontSize: 8, bold: true, fill: 'E4D4F4', fontFace: CONFIG.fonts.primary, align: 'right' } },
	      { text: 'Var', options: { fontSize: 8, bold: true, fill: 'E4D4F4', fontFace: CONFIG.fonts.primary, align: 'right' } },
	      { text: 'Var%', options: { fontSize: 8, bold: true, fill: 'E4D4F4', fontFace: CONFIG.fonts.primary, align: 'right' } },
	      { text: 'Explanation', options: { fontSize: 8, bold: true, fill: 'E4D4F4', fontFace: CONFIG.fonts.primary } },
	    ]);

	    // Revenue section header
	    tableData4.push([
	      { text: 'Revenue', options: { fontSize: 9, bold: true, fill: CONFIG.colors.primary, color: CONFIG.colors.white, fontFace: CONFIG.fonts.primary, colspan: 7 } },
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
	        (budget === 0 ? 'No budget baseline for this line.'
	          : (Math.abs(Number(variancePct)) < 10
	              ? 'Within \u00b110% of budget; no commentary required.'
	              : nm + ' came in ' + (variance >= 0 ? 'above' : 'below') + ' budget by ' + Math.abs(Number(variancePct)) + '% (' + formatCurrency(actual) + ' vs ' + formatCurrency(budget) + ').'));
	      tableData4.push([
	        { text: glAccounts[idx] || '', options: { fontSize: 8, fontFace: CONFIG.fonts.primary } },
	        { text: nm, options: { fontSize: 8, fontFace: CONFIG.fonts.primary } },
	        { text: formatCurrency(actual), options: { fontSize: 8, fontFace: CONFIG.fonts.primary, align: 'right' } },
	        { text: formatCurrency(budget), options: { fontSize: 8, fontFace: CONFIG.fonts.primary, align: 'right' } },
	        { text: formatCurrency(variance), options: { fontSize: 8, fontFace: CONFIG.fonts.primary, align: 'right', color: varColor } },
	        { text: varArrow + ' ' + variancePct + '%', options: { fontSize: 8, fontFace: CONFIG.fonts.primary, align: 'right', color: varColor } },
	        { text: explanation, options: { fontSize: 7, fontFace: CONFIG.fonts.primary, color: CONFIG.colors.gray500 } },
	      ]);
	    });

	    // Total Revenue row
	    const budgetRevTotal = (budgetDataLocal.revenue && typeof budgetDataLocal.revenue.current === 'number') ? budgetDataLocal.revenue.current : (Number(budgetDataLocal.revenue?.current) || 0);
	    const revVar = (Number(inputData.revenue.current) || 0) - (Number(budgetRevTotal) || 0);
	    const revVarPct = (Number(budgetRevTotal) || 0) !== 0 ? ((revVar / Number(budgetRevTotal)) * 100).toFixed(0) : '0';
	    const revVarColor = revVar >= 0 ? CONFIG.colors.positive : CONFIG.colors.negative;
	    const revArrow = revVar >= 0 ? '▲' : '▼';

	    tableData4.push([
	      { text: 'Total Revenue', options: { fontSize: 8, bold: true, fill: CONFIG.colors.gray100, fontFace: CONFIG.fonts.primary, colspan: 2 } },
	      { text: '', options: { fill: CONFIG.colors.gray100 } },
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
	      title: 'Revenue vs Budget',
	      subtitle: `$${Math.round((Number(inputData.revenue.current) || 0) / 1000000)}M`,
	      actual: Number(inputData.revenue.current) || 0,
	      budget: budgetRev,
	      varPct: revPct,
	    });

	    addBudgetChart(slide4, pptx, {
	      x: 3.83,
	      title: 'Gross Profit vs Budget',
	      subtitle: `$${Math.round((Number(inputData.gross_profit.current) || 0) / 1000000)}M`,
	      actual: Number(inputData.gross_profit.current) || 0,
	      budget: budgetGP,
	      varPct: gpPct,
	    });

	    addBudgetChart(slide4, pptx, {
	      x: 7.15,
	      title: 'Profit After Tax vs Budget',
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

    const pptxBase64 = await pptx.write('base64');
    const safeCompany = String(inputData.company_name || 'Company');
    const filename = `${CONFIG.metadata.filenamePrefix}_${safeCompany.replace(/[^a-z0-9]/gi, '_')}.pptx`;
    
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

    const result = await generatePresentation({
      pptxConfig,
      financialData: normalizedFinancialData,
      coverImageUrl,
      chartImageUrl,
      chartData,
      shopCommentary,
    });

    // Persist the pptx to disk so you can see it in the output folder
    if (!result || !result.pptx_base64 || !result.filename) {
      throw new Error('Generator returned empty result (missing filename or pptx_base64).');
    }

    const outPath = path.join(OUTPUT_DIR, result.filename);
    const buf = Buffer.from(result.pptx_base64, 'base64');
    fs.writeFileSync(outPath, buf);

    console.log(`[OK] wrote ${outPath} (${buf.length} bytes)`);

    // NOTE: do NOT return pptx_base64 in the response body — a full deck with
    // embedded images is several MB and exceeds Dify's 1 MB HTTP-node response
    // limit. The file is saved to disk and fetchable via download_url; callers
    // that want the bytes can pass ?include_base64=1.
    const wantBase64 = req.query.include_base64 === '1' || req.query.include_base64 === 'true';
    res.json({
      success: true,
      filename: result.filename,
      filepath: outPath,
      file_size_bytes: buf.length,
      download_url: `/download/${encodeURIComponent(result.filename)}`,
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

app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
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

<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Namma BI — Enterprise Analytics</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/PptxGenJS/3.12.0/pptxgen.bundle.js"></script>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  /* Core palette — deep navy + electric cyan + warm amber */
  --bg-base:       #0a0e1a;
  --bg-panel:      #111827;
  --bg-card:       #1a2236;
  --bg-card-hover: #1f2a42;
  --bg-input:      #0f1624;
  --border:        rgba(99,126,219,0.15);
  --border-bright: rgba(99,126,219,0.35);
  --glow:          rgba(56,189,248,0.08);

  --text-primary:  #f0f4ff;
  --text-secondary:#94a3c8;
  --text-muted:    #4a5578;

  --accent:        #38bdf8;       /* electric cyan */
  --accent-dim:    rgba(56,189,248,0.12);
  --accent-glow:   rgba(56,189,248,0.25);
  --amber:         #fbbf24;
  --amber-dim:     rgba(251,191,36,0.12);
  --green:         #34d399;
  --green-dim:     rgba(52,211,153,0.12);
  --rose:          #f87171;
  --rose-dim:      rgba(248,113,113,0.12);
  --purple:        #a78bfa;
  --purple-dim:    rgba(167,139,250,0.12);

  --radius-sm:  6px;
  --radius-md:  10px;
  --radius-lg:  16px;
  --sidebar-w:  240px;
  --topbar-h:   56px;
  --font:       'DM Sans', sans-serif;
  --mono:       'DM Mono', monospace;

  /* Chart palette */
  --c1: #38bdf8; --c2: #34d399; --c3: #fbbf24; --c4: #f87171;
  --c5: #a78bfa; --c6: #fb923c; --c7: #e879f9; --c8: #4ade80;
}

html { scroll-behavior: smooth; }

body {
  font-family: var(--font);
  background: var(--bg-base);
  color: var(--text-primary);
  font-size: 13px;
  line-height: 1.5;
  min-height: 100vh;
  overflow-x: hidden;
}

/* ── BACKGROUND MESH ── */
body::before {
  content: '';
  position: fixed; inset: 0; z-index: 0; pointer-events: none;
  background:
    radial-gradient(ellipse 80% 60% at 10% 0%, rgba(56,189,248,0.06) 0%, transparent 60%),
    radial-gradient(ellipse 60% 50% at 90% 100%, rgba(167,139,250,0.05) 0%, transparent 60%),
    radial-gradient(ellipse 40% 40% at 50% 50%, rgba(251,191,36,0.03) 0%, transparent 70%);
}

.app { display: flex; flex-direction: column; min-height: 100vh; position: relative; z-index: 1; }

/* ══ TOPBAR ══════════════════════════════════════════ */
.topbar {
  height: var(--topbar-h);
  background: rgba(17,24,39,0.92);
  backdrop-filter: blur(20px);
  border-bottom: 1px solid var(--border);
  display: flex; align-items: center; gap: 12px;
  padding: 0 20px;
  position: sticky; top: 0; z-index: 300;
  flex-shrink: 0;
}

.logo {
  display: flex; align-items: center; gap: 10px;
  font-weight: 700; font-size: 16px; letter-spacing: -0.5px;
  flex-shrink: 0;
}
.logo-icon {
  width: 32px; height: 32px; border-radius: 8px;
  background: linear-gradient(135deg, #38bdf8, #6366f1);
  display: flex; align-items: center; justify-content: center;
  font-size: 16px; box-shadow: 0 0 20px rgba(56,189,248,0.3);
}
.logo-text { color: var(--text-primary); }
.logo-text span { color: var(--accent); }
.logo-badge {
  font-size: 9px; font-weight: 600; letter-spacing: 1px;
  background: linear-gradient(90deg, var(--accent), var(--purple));
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
  padding: 2px 0;
}

.topbar-center { flex: 1; display: flex; align-items: center; gap: 10px; overflow: hidden; }

.file-pill {
  display: none; align-items: center; gap: 6px;
  background: var(--accent-dim); border: 1px solid rgba(56,189,248,0.2);
  color: var(--accent); border-radius: 20px; padding: 4px 12px;
  font-size: 11px; font-weight: 500; max-width: 280px; overflow: hidden;
  white-space: nowrap; text-overflow: ellipsis;
}

.date-filter-inline {
  display: none; align-items: center; gap: 8px; flex-wrap: wrap;
}
.date-filter-inline label { font-size: 11px; color: var(--text-muted); }

.topbar-right { display: flex; align-items: center; gap: 8px; flex-shrink: 0; }

/* Theme swatches */
.swatch-row { display: flex; gap: 5px; }
.swatch {
  width: 16px; height: 16px; border-radius: 50%; cursor: pointer;
  border: 2px solid transparent; transition: all 0.2s; flex-shrink: 0;
}
.swatch.on, .swatch:hover { border-color: var(--text-primary); transform: scale(1.2); }

/* ══ BUTTONS ══════════════════════════════════════════ */
.btn {
  display: inline-flex; align-items: center; gap: 6px;
  padding: 7px 14px; border-radius: var(--radius-sm);
  border: 1px solid var(--border-bright);
  background: rgba(255,255,255,0.04);
  color: var(--text-secondary);
  font-size: 12px; font-family: var(--font); font-weight: 500;
  cursor: pointer; transition: all 0.18s; white-space: nowrap;
  letter-spacing: 0.01em;
}
.btn:hover { background: rgba(255,255,255,0.08); color: var(--text-primary); border-color: var(--border-bright); }
.btn.accent {
  background: linear-gradient(135deg, rgba(56,189,248,0.2), rgba(99,102,241,0.2));
  border-color: var(--accent); color: var(--accent);
}
.btn.accent:hover { background: linear-gradient(135deg, rgba(56,189,248,0.3), rgba(99,102,241,0.3)); box-shadow: 0 0 20px rgba(56,189,248,0.2); }
.btn.solid {
  background: linear-gradient(135deg, #38bdf8, #6366f1);
  border-color: transparent; color: #fff; font-weight: 600;
}
.btn.solid:hover { opacity: 0.9; box-shadow: 0 4px 20px rgba(56,189,248,0.35); }
.btn.danger { border-color: rgba(248,113,113,0.3); color: var(--rose); }
.btn.danger:hover { background: var(--rose-dim); }
.btn.sm { padding: 5px 10px; font-size: 11px; }
.btn.icon-only { padding: 7px 9px; }

/* ══ INPUTS ══════════════════════════════════════════ */
input, select, textarea {
  font-family: var(--font); font-size: 12px;
  color: var(--text-primary); background: var(--bg-input);
  border: 1px solid var(--border); border-radius: var(--radius-sm);
  padding: 7px 10px; outline: none; transition: border-color 0.18s;
}
input:focus, select:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(56,189,248,0.1); }
select option { background: var(--bg-panel); }
input[type=date]::-webkit-calendar-picker-indicator { filter: invert(0.7) hue-rotate(180deg); cursor: pointer; }

/* ══ LAYOUT ══════════════════════════════════════════ */
.main { display: flex; flex: 1; }

.sidebar {
  width: var(--sidebar-w); background: var(--bg-panel);
  border-right: 1px solid var(--border);
  overflow-y: auto; flex-shrink: 0;
  transition: width 0.25s cubic-bezier(.4,0,.2,1), opacity 0.25s;
  display: none;
}
.sidebar.visible { display: block; }
.sidebar.collapsed { width: 0; opacity: 0; overflow: hidden; }
.sidebar::-webkit-scrollbar { width: 4px; }
.sidebar::-webkit-scrollbar-track { background: transparent; }
.sidebar::-webkit-scrollbar-thumb { background: var(--border-bright); border-radius: 4px; }

.content { flex: 1; overflow-y: auto; padding: 20px; min-width: 0; }
.content::-webkit-scrollbar { width: 6px; }
.content::-webkit-scrollbar-track { background: transparent; }
.content::-webkit-scrollbar-thumb { background: var(--border-bright); border-radius: 4px; }

/* ══ SIDEBAR INTERNALS ══════════════════════════════ */
.sb-header {
  padding: 16px 14px 8px;
  font-size: 10px; font-weight: 600; letter-spacing: 1px;
  text-transform: uppercase; color: var(--text-muted);
  border-bottom: 1px solid var(--border);
}
.sb-section { padding: 8px 0; }
.sb-label {
  font-size: 9px; font-weight: 600; letter-spacing: 1.2px;
  text-transform: uppercase; color: var(--text-muted);
  padding: 6px 14px 4px;
}
.sb-field {
  display: flex; align-items: center; gap: 8px;
  padding: 7px 14px; cursor: pointer; font-size: 12px;
  border-radius: var(--radius-sm); margin: 1px 6px;
  transition: all 0.15s; color: var(--text-secondary);
}
.sb-field:hover { background: var(--glow); color: var(--text-primary); }
.sb-field-icon { font-size: 13px; flex-shrink: 0; }
.sb-field-name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.type-pill {
  font-size: 9px; font-weight: 700; padding: 2px 6px;
  border-radius: 10px; flex-shrink: 0; letter-spacing: 0.3px;
}
.pill-num  { background: rgba(56,189,248,0.15);  color: var(--accent); }
.pill-cat  { background: rgba(52,211,153,0.15);  color: var(--green); }
.pill-date { background: rgba(251,191,36,0.15);  color: var(--amber); }

/* ══ TABS ══════════════════════════════════════════ */
.tabs {
  display: flex; gap: 2px; margin-bottom: 20px;
  background: var(--bg-panel); border: 1px solid var(--border);
  border-radius: var(--radius-md); padding: 4px; width: fit-content;
}
.tab {
  padding: 7px 18px; cursor: pointer; font-size: 12px; font-weight: 500;
  border-radius: 7px; color: var(--text-muted); transition: all 0.2s;
  white-space: nowrap;
}
.tab:hover { color: var(--text-primary); background: rgba(255,255,255,0.05); }
.tab.on { background: linear-gradient(135deg, rgba(56,189,248,0.2), rgba(99,102,241,0.2)); color: var(--accent); }

/* ══ UPLOAD PAGE ══════════════════════════════════ */
.upload-page {
  display: flex; flex-direction: column; align-items: center;
  justify-content: center; min-height: 80vh; padding: 24px; gap: 28px;
}

.upload-hero {
  text-align: center; max-width: 580px; animation: fadeUp 0.6s ease both;
}
@keyframes fadeUp { from { opacity: 0; transform: translateY(24px); } to { opacity: 1; transform: translateY(0); } }

.upload-badge {
  display: inline-flex; align-items: center; gap: 6px;
  background: var(--accent-dim); border: 1px solid rgba(56,189,248,0.2);
  color: var(--accent); border-radius: 20px; padding: 5px 14px;
  font-size: 11px; font-weight: 500; margin-bottom: 20px; letter-spacing: 0.3px;
}
.upload-badge::before { content: '●'; animation: pulse 2s ease infinite; }
@keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.3} }

.upload-title {
  font-size: 36px; font-weight: 700; line-height: 1.15;
  letter-spacing: -1px; margin-bottom: 12px;
  background: linear-gradient(135deg, #f0f4ff 0%, #94a3c8 100%);
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
}
.upload-title em { font-style: italic; color: var(--accent); -webkit-text-fill-color: var(--accent); }
.upload-sub { color: var(--text-secondary); font-size: 15px; line-height: 1.6; margin-bottom: 28px; }

.dropzone {
  border: 1.5px dashed var(--border-bright); border-radius: var(--radius-lg);
  padding: 44px 32px; text-align: center; cursor: pointer;
  transition: all 0.25s; position: relative; overflow: hidden;
  background: rgba(56,189,248,0.03);
}
.dropzone::after {
  content: ''; position: absolute; inset: 0;
  background: radial-gradient(ellipse at 50% 0%, rgba(56,189,248,0.06) 0%, transparent 70%);
  pointer-events: none;
}
.dropzone:hover, .dropzone.drag {
  border-color: var(--accent); background: var(--accent-dim);
  box-shadow: 0 0 40px rgba(56,189,248,0.1);
}
.dropzone-icon { font-size: 40px; margin-bottom: 14px; display: block; }
.dropzone h3 { font-size: 16px; font-weight: 600; margin-bottom: 6px; }
.dropzone p { color: var(--text-secondary); font-size: 13px; margin-bottom: 18px; }
.drop-divider { color: var(--text-muted); font-size: 11px; margin: 12px 0; }

.feat-grid {
  display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
  gap: 10px; width: 100%; max-width: 640px;
  animation: fadeUp 0.6s 0.2s ease both;
}
.feat-card {
  background: var(--bg-card); border: 1px solid var(--border);
  border-radius: var(--radius-md); padding: 14px 12px;
  transition: all 0.2s;
}
.feat-card:hover { border-color: var(--border-bright); background: var(--bg-card-hover); transform: translateY(-2px); }
.feat-card-icon { font-size: 20px; margin-bottom: 8px; }
.feat-card-title { font-weight: 600; font-size: 12px; margin-bottom: 3px; }
.feat-card-desc { color: var(--text-muted); font-size: 11px; line-height: 1.5; }

/* ══ METRICS ══════════════════════════════════════ */
.kpi-strip {
  display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
  gap: 12px; margin-bottom: 20px;
}
.kpi {
  background: var(--bg-card); border: 1px solid var(--border);
  border-radius: var(--radius-md); padding: 16px;
  position: relative; overflow: hidden; transition: all 0.2s;
}
.kpi:hover { border-color: var(--border-bright); transform: translateY(-1px); box-shadow: 0 8px 32px rgba(0,0,0,0.3); }
.kpi::before {
  content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px;
  background: linear-gradient(90deg, var(--accent), var(--purple));
}
.kpi-label { font-size: 10px; font-weight: 600; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 8px; }
.kpi-value { font-size: 24px; font-weight: 700; letter-spacing: -0.5px; line-height: 1; }
.kpi-sub { font-size: 10px; color: var(--text-muted); margin-top: 5px; }
.kpi-trend { font-size: 11px; font-weight: 600; margin-top: 4px; }
.kpi-icon { position: absolute; top: 14px; right: 14px; font-size: 18px; opacity: 0.3; }

/* ══ CHART CARDS ══════════════════════════════════ */
.chart-grid {
  display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
  gap: 14px;
}
.chart-card {
  background: var(--bg-card); border: 1px solid var(--border);
  border-radius: var(--radius-md); padding: 16px;
  position: relative; overflow: hidden; transition: all 0.2s;
}
.chart-card:hover { border-color: var(--border-bright); box-shadow: 0 8px 32px rgba(0,0,0,0.25); }
.chart-card-header { display: flex; align-items: center; gap: 8px; margin-bottom: 14px; }
.chart-card-title { font-size: 12px; font-weight: 600; flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--text-primary); }
.chart-card-controls { display: flex; gap: 4px; opacity: 0; transition: opacity 0.2s; }
.chart-card:hover .chart-card-controls { opacity: 1; }
.chart-wrap { position: relative; }

.chart-type-bar {
  display: flex; gap: 3px; background: var(--bg-input);
  border-radius: var(--radius-sm); padding: 3px; border: 1px solid var(--border);
}
.ct-btn {
  padding: 3px 7px; border-radius: 4px; font-size: 10px;
  cursor: pointer; color: var(--text-muted); transition: all 0.15s;
  background: none; border: none; font-family: var(--font); white-space: nowrap;
}
.ct-btn:hover { color: var(--text-primary); background: rgba(255,255,255,0.06); }
.ct-btn.on { background: rgba(56,189,248,0.2); color: var(--accent); }

/* ══ FILTER STRIP ══════════════════════════════════ */
.filter-strip {
  display: flex; flex-wrap: wrap; gap: 8px; align-items: center;
  padding: 10px 14px; background: var(--bg-panel);
  border: 1px solid var(--border); border-radius: var(--radius-md);
  margin-bottom: 16px; display: none;
}
.filter-strip-label { font-size: 10px; color: var(--text-muted); font-weight: 600; letter-spacing: 0.5px; text-transform: uppercase; }
.filter-chip {
  display: flex; align-items: center; gap: 5px;
  padding: 3px 10px; background: var(--accent-dim);
  border: 1px solid rgba(56,189,248,0.2); color: var(--accent);
  border-radius: 20px; font-size: 11px; font-weight: 500;
}
.filter-chip-close {
  background: none; border: none; cursor: pointer; color: var(--accent);
  font-size: 14px; line-height: 1; padding: 0; margin-left: 2px; opacity: 0.7;
}
.filter-chip-close:hover { opacity: 1; }

/* ══ DATA TABLE ══════════════════════════════════ */
.tbl-container { overflow-x: auto; border-radius: var(--radius-md); border: 1px solid var(--border); }
.data-tbl { width: 100%; border-collapse: collapse; font-size: 12px; }
.data-tbl th {
  background: var(--bg-panel); padding: 9px 12px;
  text-align: left; font-size: 10px; font-weight: 600;
  color: var(--text-muted); border-bottom: 1px solid var(--border);
  white-space: nowrap; cursor: pointer; letter-spacing: 0.3px;
  text-transform: uppercase; user-select: none;
}
.data-tbl th:hover { color: var(--accent); }
.data-tbl td {
  padding: 8px 12px; border-bottom: 1px solid rgba(99,126,219,0.07);
  max-width: 180px; overflow: hidden; text-overflow: ellipsis;
  white-space: nowrap; color: var(--text-secondary);
}
.data-tbl tr:hover td { background: var(--glow); color: var(--text-primary); }
.data-tbl td:first-child { color: var(--text-primary); font-weight: 500; }

/* ══ AI PANEL ══════════════════════════════════════ */
.ai-panel {
  background: var(--bg-card); border: 1px solid var(--border);
  border-radius: var(--radius-md); padding: 18px; margin-bottom: 16px;
}
.ai-panel-header { display: flex; align-items: center; gap: 10px; margin-bottom: 14px; }
.ai-label {
  display: inline-flex; align-items: center; gap: 5px;
  background: linear-gradient(135deg, rgba(99,102,241,0.2), rgba(167,139,250,0.2));
  border: 1px solid rgba(167,139,250,0.3); color: var(--purple);
  border-radius: 20px; padding: 3px 12px; font-size: 10px; font-weight: 700;
  letter-spacing: 0.5px; text-transform: uppercase;
}
.ai-title { font-size: 13px; font-weight: 600; color: var(--text-primary); }
.ai-subtitle { font-size: 11px; color: var(--text-muted); }
.ai-input-row { display: flex; gap: 8px; }
.ai-input-row input { flex: 1; }
.ai-result {
  margin-top: 14px; padding: 14px 16px;
  background: rgba(167,139,250,0.06); border-left: 2px solid var(--purple);
  border-radius: 0 var(--radius-sm) var(--radius-sm) 0;
  font-size: 13px; line-height: 1.75; color: var(--text-secondary);
}
.ai-loading {
  display: flex; align-items: center; gap: 10px;
  color: var(--text-muted); font-size: 12px; padding: 10px 0;
}
.dots { display: flex; gap: 4px; }
.dots span {
  width: 5px; height: 5px; border-radius: 50%; background: var(--purple);
  animation: blink 1.2s ease infinite;
}
.dots span:nth-child(2) { animation-delay: 0.2s; }
.dots span:nth-child(3) { animation-delay: 0.4s; }
@keyframes blink { 0%,80%,100%{opacity:0.15;transform:scale(0.8)} 40%{opacity:1;transform:scale(1)} }
.suggestion-row { display: flex; gap: 6px; flex-wrap: wrap; margin-top: 12px; }
.suggestion-chip {
  padding: 4px 12px; background: var(--bg-input); border: 1px solid var(--border);
  color: var(--text-secondary); border-radius: 20px; font-size: 11px;
  cursor: pointer; transition: all 0.15s; font-family: var(--font);
}
.suggestion-chip:hover { border-color: var(--purple); color: var(--purple); background: var(--purple-dim); }

/* ══ BUILDER ══════════════════════════════════════ */
.builder-panel {
  background: var(--bg-card); border: 1px solid var(--border);
  border-radius: var(--radius-md); padding: 18px; margin-bottom: 18px;
}
.builder-title { font-size: 13px; font-weight: 600; margin-bottom: 14px; display: flex; align-items: center; gap: 8px; }
.builder-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; }
.form-group { display: flex; flex-direction: column; gap: 5px; }
.form-label { font-size: 10px; font-weight: 600; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; }
.form-group select, .form-group input { width: 100%; }

/* ══ EXPORT OVERLAY ══════════════════════════════ */
.overlay {
  display: none; position: fixed; inset: 0; z-index: 400;
  background: rgba(0,0,0,0.7); backdrop-filter: blur(8px);
  align-items: center; justify-content: center; padding: 20px;
}
.overlay.show { display: flex; }
.modal {
  background: var(--bg-panel); border: 1px solid var(--border-bright);
  border-radius: var(--radius-lg); padding: 24px;
  width: 100%; max-width: 500px; max-height: 90vh; overflow-y: auto;
  box-shadow: 0 40px 80px rgba(0,0,0,0.5);
  animation: modalIn 0.2s ease;
}
@keyframes modalIn { from { opacity: 0; transform: scale(0.97) translateY(8px); } to { opacity: 1; transform: scale(1) translateY(0); } }
.modal-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px; }
.modal-title { font-size: 15px; font-weight: 700; }
.export-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.export-card {
  background: var(--bg-card); border: 1px solid var(--border);
  border-radius: var(--radius-md); padding: 18px; text-align: center;
  cursor: pointer; transition: all 0.2s;
}
.export-card:hover { border-color: var(--accent); background: var(--accent-dim); transform: translateY(-2px); box-shadow: 0 8px 24px rgba(56,189,248,0.1); }
.export-card .ec-icon { font-size: 32px; margin-bottom: 8px; }
.export-card .ec-title { font-weight: 600; font-size: 13px; margin-bottom: 4px; }
.export-card .ec-desc { font-size: 11px; color: var(--text-muted); line-height: 1.5; }
.export-msg { font-size: 12px; color: var(--green); margin-top: 12px; display: none; }

/* ══ TOOLBAR ══════════════════════════════════════ */
.section-toolbar {
  display: flex; align-items: center; gap: 10px;
  margin-bottom: 18px; flex-wrap: wrap;
}
.section-title { font-size: 16px; font-weight: 700; flex: 1; letter-spacing: -0.3px; }
.section-sub { font-size: 11px; color: var(--text-muted); font-weight: 400; margin-top: 1px; }

/* ══ SLICERS / QUICK FILTERS ══════════════════════ */
.slicer-row {
  display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 16px;
  display: none;
}
.slicer-group { display: flex; flex-direction: column; gap: 4px; }
.slicer-label { font-size: 9px; font-weight: 600; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; }
.slicer-select { min-width: 130px; }

/* ══ RESPONSIVE ══════════════════════════════════ */
@media (max-width: 768px) {
  .sidebar { display: none !important; }
  .chart-grid { grid-template-columns: 1fr; }
  .kpi-strip { grid-template-columns: repeat(2, 1fr); }
  .content { padding: 12px; }
  .topbar { padding: 0 12px; }
  .date-filter-inline { display: none !important; }
  .export-grid { grid-template-columns: 1fr; }
  .builder-grid { grid-template-columns: 1fr 1fr; }
  .upload-title { font-size: 24px; }
}

@media print {
  .topbar, .sidebar, .tabs, .filter-strip, .ai-panel, .builder-panel,
  .chart-card-controls, .chart-type-bar, .slicer-row, .export-overlay { display: none !important; }
  body { background: white; color: black; }
  .chart-card { break-inside: avoid; }
}

/* ══ SCROLLBAR (global) ══════════════════════════ */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: var(--border-bright); border-radius: 3px; }

/* ══ PROGRESS BAR ══════════════════════════════ */
.progress-bar { height: 2px; background: var(--border); border-radius: 1px; overflow: hidden; margin-top: 6px; }
.progress-fill { height: 100%; background: linear-gradient(90deg, var(--accent), var(--purple)); border-radius: 1px; transition: width 0.4s ease; }

/* ══ TOOLTIP ══════════════════════════════════ */
[title] { position: relative; }

/* ══ EMPTY STATE ══════════════════════════════ */
.empty-state {
  text-align: center; padding: 48px 20px; color: var(--text-muted);
  border: 1px dashed var(--border); border-radius: var(--radius-md);
}
.empty-state-icon { font-size: 32px; margin-bottom: 10px; }
.empty-state-text { font-size: 13px; }
</style>
</head>
<body>

<div class="app" id="app">

  <!-- ══ TOPBAR ══ -->
  <div class="topbar">
    <div class="logo">
      <div class="logo-icon">📊</div>
      <div>
        <div class="logo-text">Namma<span> BI</span></div>
        <div class="logo-badge">ENTERPRISE</div>
      </div>
    </div>

    <div class="topbar-center">
      <div class="file-pill" id="file-pill">📄 <span id="file-pill-name"></span></div>
      <div class="date-filter-inline" id="date-filter-inline">
        <label>Date range</label>
        <select id="date-col-sel" onchange="onDateColChange()" style="min-width:110px"></select>
        <input type="date" id="date-from" onchange="applyDateFilter()">
        <span style="color:var(--text-muted);font-size:11px">–</span>
        <input type="date" id="date-to" onchange="applyDateFilter()">
        <button class="btn sm" onclick="clearDateFilter()" style="padding:5px 8px;color:var(--rose)">✕</button>
      </div>
    </div>

    <div class="topbar-right">
      <div class="swatch-row" title="Color accent">
        <div class="swatch on" style="background:linear-gradient(135deg,#38bdf8,#6366f1)" onclick="setAccent('default',this)"></div>
        <div class="swatch" style="background:linear-gradient(135deg,#34d399,#0891b2)" onclick="setAccent('teal',this)"></div>
        <div class="swatch" style="background:linear-gradient(135deg,#fbbf24,#f87171)" onclick="setAccent('amber',this)"></div>
        <div class="swatch" style="background:linear-gradient(135deg,#a78bfa,#ec4899)" onclick="setAccent('violet',this)"></div>
        <div class="swatch" style="background:linear-gradient(135deg,#f0f4ff,#94a3c8)" onclick="setAccent('mono',this)"></div>
      </div>
      <button class="btn sm" id="sb-btn" onclick="toggleSidebar()" style="display:none">☰ Fields</button>
      <button class="btn accent" id="export-btn" style="display:none" onclick="showExport()">⬇ Export</button>
      <button class="btn sm" id="new-btn" style="display:none" onclick="resetApp()">↑ New file</button>
    </div>
  </div>

  <div class="main">

    <!-- ══ SIDEBAR ══ -->
    <div class="sidebar" id="sidebar">
      <div class="sb-header">Data fields</div>
      <div class="sb-section">
        <div class="sb-label">Dimensions & Measures</div>
        <div id="field-list"></div>
      </div>
      <div class="sb-section" id="filter-section" style="display:none">
        <div class="sb-label" style="color:var(--amber)">Active filters</div>
        <div id="active-filter-list" style="padding:0 8px"></div>
      </div>
    </div>

    <!-- ══ CONTENT ══ -->
    <div class="content" id="content">

      <!-- UPLOAD PAGE -->
      <div id="upload-page" class="upload-page">
        <div class="upload-hero">
          <div class="upload-badge">Live analytics engine</div>
          <div class="upload-title">Transform data into<br><em>insights instantly</em></div>
          <div class="upload-sub">Upload any Excel or CSV file. Namma BI auto-generates<br>Power BI-style dashboards, KPIs, and AI-powered analysis.</div>

          <div class="dropzone" id="dropzone">
            <span class="dropzone-icon">☁</span>
            <h3>Drop your file here</h3>
            <p>Supports .xlsx, .xls, .csv — up to 100k rows</p>
            <div style="display:flex;gap:10px;justify-content:center;flex-wrap:wrap">
              <button class="btn solid" onclick="document.getElementById('file-inp').click()">📂 Browse files</button>
              <button class="btn" onclick="loadSample()">⚡ Load sample data</button>
            </div>
            <div class="drop-divider">or drag and drop</div>
            <input type="file" id="file-inp" accept=".xlsx,.xls,.csv" style="display:none">
          </div>
        </div>

        <div class="feat-grid">
          <div class="feat-card"><div class="feat-card-icon">🤖</div><div class="feat-card-title">AI Analysis</div><div class="feat-card-desc">Claude-powered natural language Q&A on your data</div></div>
          <div class="feat-card"><div class="feat-card-icon">📊</div><div class="feat-card-title">Auto Charts</div><div class="feat-card-desc">Smart chart detection — bar, line, pie, waterfall & more</div></div>
          <div class="feat-card"><div class="feat-card-icon">🎛</div><div class="feat-card-title">Slicers & Filters</div><div class="feat-card-desc">Interactive slicers update all visuals simultaneously</div></div>
          <div class="feat-card"><div class="feat-card-icon">⬇</div><div class="feat-card-title">PPT & JPG Export</div><div class="feat-card-desc">Download full dashboards as PowerPoint or JPG images</div></div>
          <div class="feat-card"><div class="feat-card-icon">📅</div><div class="feat-card-title">Date Intelligence</div><div class="feat-card-desc">Auto date detection with range filters across all charts</div></div>
          <div class="feat-card"><div class="feat-card-icon">🎨</div><div class="feat-card-title">5 Themes</div><div class="feat-card-desc">Enterprise-grade color palettes with one click</div></div>
        </div>
      </div>

      <!-- DASHBOARD PAGE -->
      <div id="dash-page" style="display:none">
        <div class="tabs">
          <div class="tab on" onclick="goTab('overview',this)">📊 Overview</div>
          <div class="tab" onclick="goTab('analysis',this)">📈 Analysis</div>
          <div class="tab" onclick="goTab('ai',this)">🤖 AI Insights</div>
          <div class="tab" onclick="goTab('data',this)">📋 Data</div>
          <div class="tab" onclick="goTab('build',this)">🔧 Custom</div>
        </div>

        <!-- OVERVIEW TAB -->
        <div id="tab-overview">
          <div class="section-toolbar">
            <div>
              <div class="section-title" id="report-title">Dashboard Overview</div>
              <div class="section-sub" id="report-sub"></div>
            </div>
            <button class="btn sm" onclick="refreshAll()">↺ Refresh</button>
            <button class="btn accent sm" onclick="exportJPG()">📷 JPG</button>
            <button class="btn solid sm" onclick="exportPPT()">📑 PPT</button>
          </div>
          <div class="slicer-row" id="slicer-row"></div>
          <div class="filter-strip" id="filter-strip">
            <span class="filter-strip-label">🔎 Filters</span>
            <div id="filter-chips"></div>
            <button class="btn sm danger" style="margin-left:auto" onclick="clearAllFilters()">Clear all</button>
          </div>
          <div class="kpi-strip" id="kpi-strip"></div>
          <div class="chart-grid" id="chart-grid-overview"></div>
        </div>

        <!-- ANALYSIS TAB -->
        <div id="tab-analysis" style="display:none">
          <div class="section-toolbar">
            <div>
              <div class="section-title">Deep Analysis</div>
              <div class="section-sub">Detailed breakdowns and trend analysis</div>
            </div>
            <button class="btn accent sm" onclick="exportJPG('analysis')">📷 JPG</button>
            <button class="btn solid sm" onclick="exportPPT('analysis')">📑 PPT</button>
          </div>
          <div class="kpi-strip" id="kpi-analysis"></div>
          <div class="chart-grid" id="chart-grid-analysis"></div>
        </div>

        <!-- AI TAB -->
        <div id="tab-ai" style="display:none">
          <div class="ai-panel">
            <div class="ai-panel-header">
              <span class="ai-label">✦ Claude AI</span>
              <div>
                <div class="ai-title">Ask anything about your data</div>
                <div class="ai-subtitle">Powered by Claude — understands your dataset structure and values</div>
              </div>
            </div>
            <div class="ai-input-row">
              <input type="text" id="ai-q" placeholder="e.g. Which category drives the most revenue? Any outliers?" onkeydown="if(event.key==='Enter')askAI()">
              <button class="btn solid" onclick="askAI()">✦ Ask</button>
            </div>
            <div id="ai-result"></div>
            <div class="suggestion-row" id="ai-suggestions"></div>
          </div>
          <div id="ai-history" style="display:flex;flex-direction:column;gap:12px"></div>
        </div>

        <!-- DATA TAB -->
        <div id="tab-data" style="display:none">
          <div class="section-toolbar" style="margin-bottom:12px">
            <div class="section-title">Data Table</div>
            <input type="text" id="search-inp" placeholder="🔍 Search…" oninput="doSearch()" style="max-width:220px;flex:1">
            <select id="sort-col-sel" onchange="doSearch()" style="min-width:120px"></select>
            <select id="sort-dir-sel" onchange="doSearch()">
              <option value="asc">↑ Asc</option><option value="desc">↓ Desc</option>
            </select>
            <span id="row-count" style="font-size:11px;color:var(--text-muted);white-space:nowrap"></span>
            <button class="btn sm" onclick="exportCSV()">⬇ CSV</button>
          </div>
          <div class="tbl-container">
            <table class="data-tbl"><thead id="tbl-head"></thead><tbody id="tbl-body"></tbody></table>
          </div>
        </div>

        <!-- CUSTOM BUILD TAB -->
        <div id="tab-build" style="display:none">
          <div class="builder-panel">
            <div class="builder-title">⚙ Chart Builder <span style="font-size:11px;font-weight:400;color:var(--text-muted)">— craft your own visualisation</span></div>
            <div class="builder-grid">
              <div class="form-group"><div class="form-label">Chart type</div>
                <select id="b-type">
                  <option value="bar">Bar</option><option value="line">Line</option>
                  <option value="pie">Pie</option><option value="doughnut">Donut</option>
                  <option value="radar">Radar</option><option value="polarArea">Polar area</option>
                  <option value="bar_h">Horizontal bar</option>
                </select>
              </div>
              <div class="form-group"><div class="form-label">X / Category</div><select id="b-x"></select></div>
              <div class="form-group"><div class="form-label">Y / Value</div><select id="b-y"></select></div>
              <div class="form-group"><div class="form-label">Aggregation</div>
                <select id="b-agg">
                  <option value="sum">Sum</option><option value="avg">Average</option>
                  <option value="count">Count</option><option value="max">Max</option>
                  <option value="min">Min</option><option value="distinct">Distinct count</option>
                </select>
              </div>
              <div class="form-group"><div class="form-label">Top N</div><input type="number" id="b-topn" value="10" min="1" max="200"></div>
              <div class="form-group"><div class="form-label">Chart title</div><input type="text" id="b-title" placeholder="Auto"></div>
            </div>
            <div style="margin-top:14px;display:flex;gap:8px">
              <button class="btn solid" onclick="buildCustomChart()">✨ Add chart</button>
              <button class="btn" onclick="clearCustomCharts()">🗑 Clear custom</button>
            </div>
          </div>
          <div class="chart-grid" id="chart-grid-custom">
            <div class="empty-state"><div class="empty-state-icon">📊</div><div class="empty-state-text">Use the builder above to add custom charts</div></div>
          </div>
        </div>

      </div><!-- /dash-page -->
    </div><!-- /content -->
  </div><!-- /main -->
</div><!-- /app -->

<!-- ══ EXPORT OVERLAY ══ -->
<div class="overlay" id="export-overlay" onclick="if(event.target===this)closeOverlay()">
  <div class="modal">
    <div class="modal-header">
      <span class="modal-title">⬇ Export dashboard</span>
      <button class="btn sm icon-only" onclick="closeOverlay()">✕</button>
    </div>
    <div class="export-grid">
      <div class="export-card" onclick="exportPPT('overview')">
        <div class="ec-icon">📑</div>
        <div class="ec-title">PowerPoint (.pptx)</div>
        <div class="ec-desc">Full dashboard as a professional slide deck with charts & KPIs</div>
      </div>
      <div class="export-card" onclick="exportJPG('overview')">
        <div class="ec-icon">📷</div>
        <div class="ec-title">JPG Images</div>
        <div class="ec-desc">Each chart exported as a high-quality JPG image</div>
      </div>
      <div class="export-card" onclick="exportCSV()">
        <div class="ec-icon">📄</div>
        <div class="ec-title">CSV data</div>
        <div class="ec-desc">Download filtered dataset as a CSV spreadsheet</div>
      </div>
      <div class="export-card" onclick="window.print()">
        <div class="ec-icon">🖨</div>
        <div class="ec-title">Print / PDF</div>
        <div class="ec-desc">Print dashboard or save as PDF via your browser</div>
      </div>
    </div>
    <div class="export-msg" id="export-msg"></div>
  </div>
</div>

<!-- ══ SCRIPTS ══ -->
<script>
/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   CONFIG — add your Anthropic API key here
   ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
const ANTHROPIC_API_KEY = '';   // paste your key: sk-ant-...

/* ━━━ STATE ━━━ */
let raw=[], filtered=[], headers=[], types={};
let chartInst={}, activeFilters={};
let dateFilter={col:null,from:null,to:null};
let sortState={col:'',dir:'asc'};
let customChartConfigs=[];
let currentAccent='default';

const PALETTES = {
  default: ['#38bdf8','#34d399','#fbbf24','#f87171','#a78bfa','#fb923c','#e879f9','#4ade80','#f472b6','#22d3ee'],
  teal:    ['#34d399','#0891b2','#6ee7b7','#38bdf8','#059669','#67e8f9','#10b981','#0284c7','#a7f3d0','#bae6fd'],
  amber:   ['#fbbf24','#f87171','#fb923c','#fde68a','#ef4444','#fed7aa','#fca5a5','#f59e0b','#fecdd3','#fef3c7'],
  violet:  ['#a78bfa','#ec4899','#818cf8','#f472b6','#7c3aed','#c084fc','#e879f9','#6366f1','#d946ef','#8b5cf6'],
  mono:    ['#f0f4ff','#94a3c8','#cbd5e1','#64748b','#e2e8f0','#475569','#f8fafc','#334155','#e2e8f0','#1e293b'],
};

const CHART_COLORS = () => PALETTES[currentAccent] || PALETTES.default;

/* ━━━ FILE HANDLING ━━━ */
const dz = document.getElementById('dropzone');
dz.addEventListener('dragover', e=>{ e.preventDefault(); dz.classList.add('drag'); });
dz.addEventListener('dragleave', ()=> dz.classList.remove('drag'));
dz.addEventListener('drop', e=>{ e.preventDefault(); dz.classList.remove('drag'); handleFile(e.dataTransfer.files[0]); });
document.getElementById('file-inp').addEventListener('change', e=> handleFile(e.target.files[0]));

function handleFile(f) {
  if(!f) return;
  document.getElementById('file-pill-name').textContent = f.name;
  document.getElementById('file-pill').style.display = 'flex';
  const rdr = new FileReader();
  if(f.name.toLowerCase().endsWith('.csv')) {
    rdr.onload = e => parseCSV(e.target.result);
    rdr.readAsText(f);
  } else {
    rdr.onload = e => {
      const wb = XLSX.read(e.target.result, {type:'binary', cellDates:true});
      ingestData(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval:'', raw:false}));
    };
    rdr.readAsBinaryString(f);
  }
}

function parseCSV(txt) {
  const lines = txt.trim().split('\n');
  const del = lines[0].includes('\t') ? '\t' : ',';
  const hdrs = lines[0].split(del).map(h=>h.trim().replace(/^"|"$/g,''));
  const data = lines.slice(1).filter(l=>l.trim()).map(l=>{
    const v = l.split(del).map(s=>s.trim().replace(/^"|"$/g,''));
    const o={}; hdrs.forEach((h,i)=>o[h]=v[i]||''); return o;
  });
  ingestData(data);
}

function ingestData(data) {
  raw = data.filter(r=>Object.values(r).some(v=>v!==''));
  if(!raw.length){ alert('No data found.'); return; }
  headers = Object.keys(raw[0]);
  detectTypes();
  filtered = [...raw];
  showDash();
  populateBuilder();
  populateSortSel();
  populateDateColSel();
  buildSlicers();
  renderSidebar();
  renderKPIs('kpi-strip');
  buildOverviewCharts();
  buildAnalysisCharts();
  renderTable(filtered);
  buildAISuggestions();
  updateSubtitle();
}

function detectTypes() {
  types = {};
  headers.forEach(h=>{
    const vals = raw.map(r=>r[h]).filter(v=>v!=='');
    const numR = vals.filter(v=>!isNaN(parseFloat(v))&&v.toString().trim()!=='').length/vals.length;
    const dateR = vals.filter(v=>/\d{4}[-\/]\d{1,2}[-\/]\d{1,2}|\d{1,2}[-\/]\d{1,2}[-\/]\d{4}/.test(v)).length/vals.length;
    types[h] = numR>0.7 ? 'number' : dateR>0.4 ? 'date' : 'category';
  });
}

/* ━━━ UI STATE ━━━ */
function showDash() {
  document.getElementById('upload-page').style.display = 'none';
  document.getElementById('dash-page').style.display   = '';
  document.getElementById('sidebar').classList.add('visible');
  ['sb-btn','export-btn','new-btn'].forEach(id=> document.getElementById(id).style.display='');
  if(headers.some(h=>types[h]==='date'))
    document.getElementById('date-filter-inline').style.display='flex';
}

function goTab(name, el) {
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));
  el.classList.add('on');
  ['overview','analysis','ai','data','build'].forEach(t=>{
    document.getElementById('tab-'+t).style.display = (t===name)?'':'none';
  });
  if(name==='data') renderTable(filtered);
}

function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('collapsed');
}

function resetApp() {
  raw=[]; filtered=[]; headers=[]; types={}; activeFilters={};
  dateFilter={col:null,from:null,to:null};
  Object.values(chartInst).forEach(c=>{ try{c.destroy()}catch(e){} });
  chartInst={};
  document.getElementById('upload-page').style.display='';
  document.getElementById('dash-page').style.display='none';
  document.getElementById('sidebar').classList.remove('visible');
  ['sb-btn','export-btn','new-btn'].forEach(id=> document.getElementById(id).style.display='none');
  document.getElementById('file-pill').style.display='none';
  document.getElementById('date-filter-inline').style.display='none';
  ['chart-grid-overview','chart-grid-analysis','chart-grid-custom','kpi-strip','kpi-analysis'].forEach(id=>{
    const el=document.getElementById(id); if(el) el.innerHTML='';
  });
  document.getElementById('file-inp').value='';
  customChartConfigs=[];
}

function updateSubtitle() {
  document.getElementById('report-title').textContent = 'Dashboard Overview';
  document.getElementById('report-sub').textContent =
    filtered.length.toLocaleString()+' rows · '+headers.length+' fields · '+
    headers.filter(h=>types[h]==='number').length+' measures · '+
    headers.filter(h=>types[h]==='category').length+' dimensions';
}

/* ━━━ KPIs ━━━ */
function renderKPIs(containerId) {
  const nums = headers.filter(h=>types[h]==='number');
  const cats = headers.filter(h=>types[h]==='category');
  const el = document.getElementById(containerId);
  const cards = [
    { label:'Total records', val:filtered.length.toLocaleString(), sub:'of '+raw.length.toLocaleString()+' total', icon:'📋' },
    { label:'Data fields', val:headers.length, sub:nums.length+' numeric, '+cats.length+' categorical', icon:'📐' },
  ];
  nums.slice(0,6).forEach(col=>{
    const vals = filtered.map(r=>parseFloat(r[col])).filter(v=>!isNaN(v));
    if(!vals.length) return;
    const s = vals.reduce((a,b)=>a+b,0);
    const pct = raw.length>0 ? ((filtered.length/raw.length-1)*100).toFixed(1) : 0;
    cards.push({ label:col, val:fmt(s), sub:'avg '+fmt(s/vals.length)+' · n='+vals.length, icon:'🔢', trend: pct });
  });
  el.innerHTML = cards.map(c=>`
    <div class="kpi">
      <div class="kpi-icon">${c.icon}</div>
      <div class="kpi-label">${c.label}</div>
      <div class="kpi-value" style="color:var(--accent)">${c.val}</div>
      <div class="kpi-sub">${c.sub}</div>
      ${c.trend!==undefined?`<div class="kpi-trend" style="color:${parseFloat(c.trend)>=0?'var(--green)':'var(--rose)'}">
        ${parseFloat(c.trend)>=0?'▲':'▼'} ${Math.abs(c.trend)}%
      </div>`:''}
      <div class="progress-bar"><div class="progress-fill" style="width:${Math.min(100,(filtered.length/Math.max(raw.length,1))*100)}%"></div></div>
    </div>`).join('');
}

function fmt(n) {
  if(isNaN(n)) return '—';
  const abs=Math.abs(n);
  if(abs>=1e9) return (n/1e9).toFixed(1)+'B';
  if(abs>=1e6) return (n/1e6).toFixed(1)+'M';
  if(abs>=1e3) return (n/1e3).toFixed(1)+'K';
  return n%1===0 ? n.toLocaleString() : n.toFixed(2);
}

/* ━━━ CHART SYSTEM ━━━ */
function buildOverviewCharts() {
  destroyGridCharts('overview');
  document.getElementById('chart-grid-overview').innerHTML='';
  const cats=headers.filter(h=>types[h]==='category');
  const nums=headers.filter(h=>types[h]==='number');
  const dates=headers.filter(h=>types[h]==='date');
  const cfgs=[];
  if(cats[0]&&nums[0]) cfgs.push({type:'bar',    x:cats[0], y:nums[0], agg:'sum',   topN:10, title:nums[0]+' by '+cats[0]});
  if(cats[0]&&nums[0]) cfgs.push({type:'doughnut',x:cats[0], y:nums[0], agg:'sum',   topN:8,  title:'Share by '+cats[0]});
  if(dates[0]&&nums[0])cfgs.push({type:'line',   x:dates[0],y:nums[0], agg:'sum',   topN:30, title:nums[0]+' over time'});
  if(cats[1]&&nums[0]) cfgs.push({type:'bar',    x:cats[1], y:nums[0], agg:'avg',   topN:10, title:'Avg '+nums[0]+' by '+cats[1]});
  cfgs.forEach(cfg=>addChartCard(cfg,'chart-grid-overview','overview'));
}

function buildAnalysisCharts() {
  destroyGridCharts('analysis');
  document.getElementById('chart-grid-analysis').innerHTML='';
  renderKPIs('kpi-analysis');
  const cats=headers.filter(h=>types[h]==='category');
  const nums=headers.filter(h=>types[h]==='number');
  const dates=headers.filter(h=>types[h]==='date');
  const cfgs=[];
  if(cats[0]&&nums[1]) cfgs.push({type:'bar_h',  x:cats[0], y:nums[1],  agg:'sum', topN:10, title:nums[1]+' by '+cats[0]});
  if(cats[0]&&nums[0]) cfgs.push({type:'radar',  x:cats[0], y:nums[0],  agg:'avg', topN:7,  title:'Radar — '+nums[0]});
  if(cats[1]&&nums[1]) cfgs.push({type:'polarArea',x:cats[1],y:nums[1], agg:'sum', topN:7,  title:'Distribution '+cats[1]});
  if(dates[0]&&nums[1])cfgs.push({type:'line',   x:dates[0],y:nums[1],  agg:'sum', topN:30, title:nums[1]+' trend'});
  if(nums[0]&&cats[0]) cfgs.push({type:'bar',    x:cats[0], y:nums[0],  agg:'count',topN:10,title:'Count by '+cats[0]});
  cfgs.forEach(cfg=>addChartCard(cfg,'chart-grid-analysis','analysis'));
}

function destroyGridCharts(tag) {
  Object.keys(chartInst).filter(k=>k.startsWith(tag+'_')).forEach(k=>{
    try{chartInst[k].destroy()}catch(e){}; delete chartInst[k];
  });
}

function addChartCard(cfg, gridId, tag) {
  const id = (tag||'c')+'_'+Math.random().toString(36).slice(2,8);
  const grid = document.getElementById(gridId);
  if(!grid) return;
  const isPolar = ['pie','doughnut','polarArea','radar'].includes(cfg.type);
  const h = isPolar ? 240 : 220;
  const card = document.createElement('div');
  card.className='chart-card'; card.id='card_'+id;
  const types_list = ['bar','line','pie','doughnut','radar','polarArea','bar_h'];
  const type_labels = {bar:'Bar',line:'Line',pie:'Pie',doughnut:'Donut',radar:'Radar',polarArea:'Polar',bar_h:'H-Bar'};
  card.innerHTML=`
    <div class="chart-card-header">
      <span class="chart-card-title">${cfg.title}</span>
      <div class="chart-card-controls">
        <div class="chart-type-bar">
          ${types_list.map(t=>`<button class="ct-btn${t===cfg.type?' on':''}" onclick="changeType('${id}','${t}',this)">${type_labels[t]}</button>`).join('')}
        </div>
        <button class="btn sm" onclick="exportOneJPG('${id}')" title="Export JPG">📷</button>
        <button class="btn sm danger" onclick="removeChart('${id}')">✕</button>
      </div>
    </div>
    <div class="chart-wrap" style="height:${h}px"><canvas id="${id}" role="img" aria-label="${cfg.title}"></canvas></div>`;
  grid.appendChild(card);
  setTimeout(()=>drawChart(id,cfg),40);
}

function drawChart(id, cfg) {
  if(chartInst[id]){ try{chartInst[id].destroy()}catch(e){} }
  const canvas = document.getElementById(id);
  if(!canvas) return;
  const realType = cfg.type==='bar_h' ? 'bar' : cfg.type;
  const {labels, values} = aggregateData(cfg.x, cfg.y, cfg.agg, cfg.topN);
  if(!labels.length) return;
  const isPolar = ['pie','doughnut','polarArea','radar'].includes(realType);
  const pal = CHART_COLORS();
  const tc='rgba(240,244,255,0.75)', gc='rgba(99,126,219,0.1)';
  chartInst[id] = new Chart(canvas, {
    type: realType,
    data: {
      labels,
      datasets:[{
        label: cfg.y||'Value',
        data: values,
        backgroundColor: isPolar ? pal.slice(0,labels.length).map(c=>c+'cc') : pal[0]+'bb',
        borderColor:     isPolar ? pal.slice(0,labels.length) : pal[0],
        borderWidth: 2.5,
        tension: 0.4, fill: false, pointRadius: 3, pointHoverRadius: 5,
        indexAxis: cfg.type==='bar_h' ? 'y' : undefined,
      }]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      animation:{duration:600, easing:'easeInOutQuart'},
      plugins:{
        legend:{
          display:isPolar, position:'bottom',
          labels:{color:tc, font:{size:10, family:'DM Sans'}, boxWidth:8, padding:8}
        },
        tooltip:{
          backgroundColor:'rgba(17,24,39,0.95)', borderColor:'rgba(99,126,219,0.3)', borderWidth:1,
          titleColor:tc, bodyColor:'rgba(148,163,200,0.9)', padding:10, cornerRadius:8,
          callbacks:{label:ctx=>(cfg.y||'Value')+': '+fmt(ctx.raw)}
        }
      },
      scales: isPolar ? {} : {
        x:{ticks:{color:tc, maxRotation:35, font:{size:10}}, grid:{color:gc}, border:{color:'transparent'}},
        y:{ticks:{color:tc, callback:v=>fmt(v), font:{size:10}}, grid:{color:gc}, border:{color:'transparent'}}
      }
    }
  });
  chartInst[id]._cfg = cfg;
}

function changeType(id, type, btn) {
  const cfg = chartInst[id]?._cfg;
  if(!cfg) return;
  cfg.type = type;
  btn.closest('.chart-type-bar').querySelectorAll('.ct-btn').forEach(b=>b.classList.remove('on'));
  btn.classList.add('on');
  drawChart(id, cfg);
}

function removeChart(id) {
  if(chartInst[id]){ try{chartInst[id].destroy()}catch(e){}; delete chartInst[id]; }
  const card = document.getElementById('card_'+id);
  if(card) card.remove();
}

function refreshAll() {
  buildOverviewCharts();
  buildAnalysisCharts();
  renderKPIs('kpi-strip');
  updateSubtitle();
}

/* ━━━ AGGREGATION ━━━ */
function aggregateData(xCol, yCol, method, topN) {
  const g={};
  filtered.forEach(row=>{
    const key = String(row[xCol]||'(blank)').slice(0,30);
    const val = parseFloat(row[yCol])||0;
    if(!g[key]) g[key]={sum:0,count:0,min:val,max:val,dist:new Set()};
    g[key].sum+=val; g[key].count++;
    g[key].min=Math.min(g[key].min,val); g[key].max=Math.max(g[key].max,val);
    g[key].dist.add(row[yCol]);
  });
  let entries = Object.entries(g).map(([k,v])=>{
    const val={sum:v.sum,avg:v.sum/v.count,count:v.count,max:v.max,min:v.min,distinct:v.dist.size}[method]||v.sum;
    return [k, parseFloat(val.toFixed(4))];
  });
  entries.sort((a,b)=>b[1]-a[1]);
  if(topN) entries=entries.slice(0,topN);
  return {labels:entries.map(e=>e[0]), values:entries.map(e=>e[1])};
}

/* ━━━ SLICERS ━━━ */
function buildSlicers() {
  const cats = headers.filter(h=>types[h]==='category');
  const row = document.getElementById('slicer-row');
  if(!cats.length){ row.style.display='none'; return; }
  row.style.display='flex';
  row.innerHTML = cats.slice(0,4).map(col=>{
    const vals=[...new Set(raw.map(r=>r[col]))].filter(Boolean).sort();
    return `<div class="slicer-group">
      <div class="slicer-label">${col}</div>
      <select class="slicer-select" onchange="slicerChange('${col}',this.value)">
        <option value="">All ${col}s</option>
        ${vals.map(v=>`<option>${v}</option>`).join('')}
      </select>
    </div>`;
  }).join('');
}

function slicerChange(col, val) {
  if(!val) { delete activeFilters[col]; }
  else { activeFilters[col] = val; }
  applyAllFilters();
}

/* ━━━ FILTERS ━━━ */
function renderSidebar() {
  document.getElementById('field-list').innerHTML = headers.map(h=>{
    const cls = types[h]==='number' ? 'pill-num' : types[h]==='date' ? 'pill-date' : 'pill-cat';
    const ico = types[h]==='number' ? '🔢' : types[h]==='date' ? '📅' : '🏷';
    const lbl = types[h]==='number' ? '#' : types[h]==='date' ? 'Date' : 'Cat';
    return `<div class="sb-field" onclick="quickFilter('${h}')">
      <span class="sb-field-icon">${ico}</span>
      <span class="sb-field-name">${h}</span>
      <span class="type-pill ${cls}">${lbl}</span>
    </div>`;
  }).join('');
}

function quickFilter(col) {
  if(types[col]==='category') {
    const vals=[...new Set(raw.map(r=>r[col]))].filter(Boolean).sort().slice(0,25);
    const v = prompt('Filter "'+col+'" (partial match):\n\nTop values: '+vals.slice(0,10).join(', ')+(vals.length>10?'...':'')+'\n\nEnter value:');
    if(v===null) return;
    activeFilters[col]=v; applyAllFilters(); renderFilterStrip();
  } else if(types[col]==='number') {
    const vals=raw.map(r=>parseFloat(r[col])).filter(v=>!isNaN(v));
    const v=prompt(`Filter "${col}" (${fmt(Math.min(...vals))} – ${fmt(Math.max(...vals))}):\n\nEnter min,max:`);
    if(v===null) return;
    const p=v.split(',').map(x=>parseFloat(x.trim()));
    if(p.length===2&&!isNaN(p[0])&&!isNaN(p[1])){ activeFilters[col]={min:p[0],max:p[1]}; applyAllFilters(); renderFilterStrip(); }
  }
}

function applyAllFilters() {
  filtered = raw.filter(row=>{
    for(const [col,val] of Object.entries(activeFilters)){
      if(typeof val==='string'&&!String(row[col]).toLowerCase().includes(val.toLowerCase())) return false;
      if(typeof val==='object'&&val.min!==undefined){ const v=parseFloat(row[col]); if(isNaN(v)||v<val.min||v>val.max) return false; }
    }
    if(dateFilter.col&&(dateFilter.from||dateFilter.to)){
      const d=new Date(row[dateFilter.col]);
      if(!isNaN(d)){
        if(dateFilter.from&&d<dateFilter.from) return false;
        if(dateFilter.to&&d>dateFilter.to)   return false;
      }
    }
    return true;
  });
  renderKPIs('kpi-strip'); renderKPIs('kpi-analysis');
  Object.keys(chartInst).forEach(id=>{ const cfg=chartInst[id]?._cfg; if(cfg) drawChart(id,cfg); });
  buildOverviewCharts(); buildAnalysisCharts();
  renderFilterStrip(); updateSubtitle();
  doSearch();
}

function renderFilterStrip() {
  const chips=Object.entries(activeFilters);
  const fs=document.getElementById('filter-strip');
  if(!chips.length){ fs.style.display='none'; return; }
  fs.style.display='flex';
  document.getElementById('filter-chips').innerHTML = chips.map(([col,val])=>
    `<div class="filter-chip">${col}: ${typeof val==='string'?val:fmt(val.min)+'–'+fmt(val.max)}
      <button class="filter-chip-close" onclick="removeFilter('${col}')">×</button>
    </div>`).join('');
}

function removeFilter(col){ delete activeFilters[col]; applyAllFilters(); }
function clearAllFilters(){ activeFilters={}; clearDateFilter(); applyAllFilters(); }

/* ━━━ DATE FILTER ━━━ */
function populateDateColSel() {
  const dc=headers.filter(h=>types[h]==='date');
  document.getElementById('date-col-sel').innerHTML=dc.map(h=>`<option>${h}</option>`).join('');
  if(dc.length) dateFilter.col=dc[0];
}
function onDateColChange(){ dateFilter.col=document.getElementById('date-col-sel').value; applyDateFilter(); }
function applyDateFilter(){
  const f=document.getElementById('date-from').value, t=document.getElementById('date-to').value;
  dateFilter.from=f?new Date(f):null; dateFilter.to=t?new Date(t):null;
  applyAllFilters();
}
function clearDateFilter(){
  document.getElementById('date-from').value=''; document.getElementById('date-to').value='';
  dateFilter.from=null; dateFilter.to=null; applyAllFilters();
}

/* ━━━ TABLE ━━━ */
function populateSortSel(){
  document.getElementById('sort-col-sel').innerHTML='<option value="">Sort by…</option>'+headers.map(h=>`<option>${h}</option>`).join('');
}
function renderTable(data) {
  document.getElementById('tbl-head').innerHTML='<tr>'+headers.map(h=>
    `<th onclick="sortBy('${h}')">${h} ${types[h]==='number'?'#':types[h]==='date'?'📅':'Aa'}</th>`).join('')+'</tr>';
  const rows=data.slice(0,1000);
  document.getElementById('tbl-body').innerHTML=rows.map(r=>'<tr>'+headers.map(h=>`<td title="${r[h]||''}">${r[h]||''}</td>`).join('')+'</tr>').join('');
  document.getElementById('row-count').textContent=data.length.toLocaleString()+' rows'+(data.length>1000?' (1k shown)':'');
}
function sortBy(col){
  if(sortState.col===col) sortState.dir=sortState.dir==='asc'?'desc':'asc';
  else { sortState.col=col; sortState.dir='asc'; }
  document.getElementById('sort-col-sel').value=col;
  document.getElementById('sort-dir-sel').value=sortState.dir;
  doSearch();
}
function doSearch(){
  const q=document.getElementById('search-inp').value.toLowerCase();
  const sc=document.getElementById('sort-col-sel').value;
  const sd=document.getElementById('sort-dir-sel').value;
  let d=filtered.filter(r=>!q||Object.values(r).some(v=>String(v).toLowerCase().includes(q)));
  if(sc){ d=[...d].sort((a,b)=>{ const av=types[sc]==='number'?parseFloat(a[sc]):String(a[sc]); const bv=types[sc]==='number'?parseFloat(b[sc]):String(b[sc]); return sd==='asc'?(av>bv?1:-1):(av<bv?1:-1); }); }
  renderTable(d);
}

/* ━━━ CUSTOM BUILDER ━━━ */
function populateBuilder(){
  document.getElementById('b-x').innerHTML=headers.map(h=>`<option>${h}</option>`).join('');
  document.getElementById('b-y').innerHTML=headers.filter(h=>types[h]==='number').map(h=>`<option>${h}</option>`).join('');
}
function buildCustomChart(){
  const cfg={type:document.getElementById('b-type').value,x:document.getElementById('b-x').value,y:document.getElementById('b-y').value,agg:document.getElementById('b-agg').value,topN:parseInt(document.getElementById('b-topn').value)||10,title:document.getElementById('b-title').value||''};
  cfg.title=cfg.title||cfg.y+' by '+cfg.x;
  const grid=document.getElementById('chart-grid-custom');
  if(grid.querySelector('.empty-state')) grid.innerHTML='';
  addChartCard(cfg,'chart-grid-custom','cust');
  customChartConfigs.push(cfg);
  goTab('build',document.querySelectorAll('.tab')[4]);
}
function clearCustomCharts(){
  Object.keys(chartInst).filter(k=>k.startsWith('cust_')).forEach(k=>{ try{chartInst[k].destroy()}catch(e){}; delete chartInst[k]; });
  document.getElementById('chart-grid-custom').innerHTML='<div class="empty-state"><div class="empty-state-icon">📊</div><div class="empty-state-text">Use the builder above to add custom charts</div></div>';
  customChartConfigs=[];
}

/* ━━━ AI INSIGHTS ━━━ */
function buildAISuggestions(){
  const cats=headers.filter(h=>types[h]==='category'), nums=headers.filter(h=>types[h]==='number');
  const sugs=['Summarise the key trends in this dataset'];
  if(cats[0]&&nums[0]) sugs.push(`Top ${cats[0]} by ${nums[0]}?`);
  if(nums.length>=2)   sugs.push(`Correlation between ${nums[0]} and ${nums[1]}?`);
  if(cats[0])          sugs.push(`Any outliers in ${cats[0]}?`);
  sugs.push('What actionable insights can I extract?');
  document.getElementById('ai-suggestions').innerHTML=sugs.map(s=>
    `<div class="suggestion-chip" onclick="document.getElementById('ai-q').value='${s.replace(/'/g,"\\'")}';askAI()">${s}</div>`).join('');
}

async function askAI(){
  const q=document.getElementById('ai-q').value.trim();
  if(!q) return;
  if(!ANTHROPIC_API_KEY){
    document.getElementById('ai-result').innerHTML=`<div class="ai-result" style="border-color:var(--amber)">⚠ No API key set. Add your Anthropic API key to the ANTHROPIC_API_KEY constant at the top of the script.</div>`;
    return;
  }
  document.getElementById('ai-result').innerHTML=`<div class="ai-loading"><div class="dots"><span></span><span></span><span></span></div><span>Analysing your dataset…</span></div>`;
  const nums=headers.filter(h=>types[h]==='number'), cats=headers.filter(h=>types[h]==='category');
  const numStats=nums.map(col=>{ const v=filtered.map(r=>parseFloat(r[col])).filter(v=>!isNaN(v)); if(!v.length) return ''; const s=v.reduce((a,b)=>a+b,0); return `${col}: sum=${fmt(s)},avg=${fmt(s/v.length)},min=${fmt(Math.min(...v))},max=${fmt(Math.max(...v))}`; }).filter(Boolean).join('\n');
  const catStats=cats.map(col=>{ const vc={}; filtered.forEach(r=>{const v=r[col]||'?';vc[v]=(vc[v]||0)+1;}); return `${col}: `+Object.entries(vc).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([k,v])=>`${k}(${v})`).join(', '); }).join('\n');
  const sample=filtered.slice(0,5).map(r=>headers.map(h=>r[h]).join(',')).join('\n');
  try {
    const resp=await fetch('https://api.anthropic.com/v1/messages',{method:'POST',headers:{'Content-Type':'application/json','x-api-key':ANTHROPIC_API_KEY,'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},
      body:JSON.stringify({model:'claude-sonnet-4-20250514',max_tokens:1000,system:'You are a senior data analyst. Respond in 4-6 sentences with specific numbers. Be precise and actionable.',messages:[{role:'user',content:`Dataset: ${filtered.length} rows, cols: ${headers.join(', ')}\nNumeric:\n${numStats}\nCategories:\n${catStats}\nSample:\n${sample}\n\nQuestion: ${q}`}]})});
    const data=await resp.json();
    const ans=data.content?.map(b=>b.text||'').join('')||'No response.';
    document.getElementById('ai-result').innerHTML=`<div class="ai-result">${ans.replace(/\n/g,'<br>')}</div>`;
    const hist=document.getElementById('ai-history');
    const entry=document.createElement('div');
    entry.className='ai-panel';
    entry.innerHTML=`<div style="font-size:11px;color:var(--text-muted);margin-bottom:6px">Q: ${q}</div><div class="ai-result" style="margin:0">${ans.replace(/\n/g,'<br>')}</div>`;
    hist.prepend(entry);
    document.getElementById('ai-q').value='';
  } catch(e) {
    document.getElementById('ai-result').innerHTML=`<div class="ai-result" style="border-color:var(--rose);color:var(--rose)">Error: ${e.message}</div>`;
  }
}

/* ━━━ THEME ━━━ */
const ACCENT_VARS = {
  default: {a:'#38bdf8',al:'rgba(56,189,248,0.12)',ag:'rgba(56,189,248,0.25)'},
  teal:    {a:'#34d399',al:'rgba(52,211,153,0.12)', ag:'rgba(52,211,153,0.25)'},
  amber:   {a:'#fbbf24',al:'rgba(251,191,36,0.12)', ag:'rgba(251,191,36,0.25)'},
  violet:  {a:'#a78bfa',al:'rgba(167,139,250,0.12)',ag:'rgba(167,139,250,0.25)'},
  mono:    {a:'#f0f4ff',al:'rgba(240,244,255,0.08)',ag:'rgba(240,244,255,0.15)'},
};
function setAccent(name, dotEl) {
  currentAccent=name;
  const v=ACCENT_VARS[name]||ACCENT_VARS.default;
  const root=document.documentElement;
  root.style.setProperty('--accent', v.a);
  root.style.setProperty('--accent-dim', v.al);
  root.style.setProperty('--accent-glow', v.ag);
  document.querySelectorAll('.swatch').forEach(s=>s.classList.remove('on'));
  dotEl.classList.add('on');
  Object.values(chartInst).forEach(c=>{ try{c.destroy()}catch(e){}; });
  chartInst={};
  setTimeout(()=>{ buildOverviewCharts(); buildAnalysisCharts(); },80);
}

/* ━━━ EXPORT — JPG ━━━ */
function exportOneJPG(id) {
  const canvas=document.getElementById(id);
  if(!canvas) return;
  const out=document.createElement('canvas');
  out.width=canvas.width; out.height=canvas.height;
  const ctx=out.getContext('2d');
  ctx.fillStyle='#111827'; ctx.fillRect(0,0,out.width,out.height);
  ctx.drawImage(canvas,0,0);
  const a=document.createElement('a');
  a.href=out.toDataURL('image/jpeg',0.95); a.download='chart_'+id+'.jpg'; a.click();
}

function exportJPG(tab) {
  const ids=Object.keys(chartInst).filter(k=>!tab||k.startsWith(tab==='overview'?'overview':tab==='analysis'?'analysis':''));
  if(!ids.length){ showExportMsg('No charts to export','amber'); return; }
  ids.forEach((id,i)=>setTimeout(()=>exportOneJPG(id), i*300));
  showExportMsg(ids.length+' JPG files downloading…');
}

/* ━━━ EXPORT — PPTX ━━━ */
async function exportPPT(tab) {
  showExportMsg('Building PowerPoint…');
  try {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.title  = 'Namma BI Report';
    pptx.author = 'Namma BI';

    /* ─ SLIDE 1: Title ─ */
    const s1 = pptx.addSlide();
    s1.background = {color:'0A0E1A'};
    s1.addShape(pptx.ShapeType.rect, {x:0,y:0,w:'100%',h:0.06,fill:{color:'38bdf8'}});
    s1.addText('Namma BI', {x:0.5,y:1.0,w:12,h:0.8,fontSize:44,bold:true,color:'F0F4FF',fontFace:'Calibri'});
    s1.addText('Enterprise Analytics Report', {x:0.5,y:1.9,w:12,h:0.5,fontSize:22,color:'94A3C8',fontFace:'Calibri'});
    const now=new Date();
    s1.addText(`Generated: ${now.toLocaleDateString()} · ${filtered.length.toLocaleString()} records · ${headers.length} fields`,
      {x:0.5,y:2.5,w:12,h:0.4,fontSize:13,color:'4A5578',fontFace:'Calibri'});

    // KPI boxes on title slide
    const nums=headers.filter(h=>types[h]==='number').slice(0,4);
    const kpiColors=['38bdf8','34d399','fbbf24','f87171'];
    nums.forEach((col,i)=>{
      const vals=filtered.map(r=>parseFloat(r[col])).filter(v=>!isNaN(v));
      if(!vals.length) return;
      const s=vals.reduce((a,b)=>a+b,0);
      const x=0.5+i*3.1, y=3.3;
      s1.addShape(pptx.ShapeType.rect,{x,y,w:2.8,h:1.5,fill:{color:'111827'},line:{color:kpiColors[i],width:1.5}});
      s1.addText(col.toUpperCase(),{x,y:y+0.12,w:2.8,h:0.3,fontSize:9,bold:true,color:'4A5578',align:'center',fontFace:'Calibri'});
      s1.addText(fmt(s),{x,y:y+0.38,w:2.8,h:0.6,fontSize:26,bold:true,color:kpiColors[i],align:'center',fontFace:'Calibri'});
      s1.addText('avg '+fmt(s/vals.length),{x,y:y+1.0,w:2.8,h:0.3,fontSize:10,color:'94A3C8',align:'center',fontFace:'Calibri'});
    });
    s1.addShape(pptx.ShapeType.rect,{x:0,y:'95%',w:'100%',h:0.06,fill:{color:'38bdf8',alpha:50}});

    /* ─ SLIDE 2: KPI Summary ─ */
    const s2=pptx.addSlide();
    s2.background={color:'0A0E1A'};
    s2.addShape(pptx.ShapeType.rect,{x:0,y:0,w:'100%',h:0.06,fill:{color:'38bdf8'}});
    s2.addText('Key Performance Indicators',{x:0.4,y:0.2,w:12,h:0.5,fontSize:22,bold:true,color:'F0F4FF',fontFace:'Calibri'});
    s2.addText(`${filtered.length.toLocaleString()} records analysed`,{x:0.4,y:0.68,w:6,h:0.3,fontSize:12,color:'4A5578',fontFace:'Calibri'});
    const allNums=headers.filter(h=>types[h]==='number').slice(0,8);
    allNums.forEach((col,i)=>{
      const vals=filtered.map(r=>parseFloat(r[col])).filter(v=>!isNaN(v));
      if(!vals.length) return;
      const s=vals.reduce((a,b)=>a+b,0);
      const row=Math.floor(i/4), col_i=i%4;
      const x=0.3+col_i*3.2, y=1.1+row*2.0;
      s2.addShape(pptx.ShapeType.rect,{x,y,w:3.0,h:1.7,fill:{color:'111827'},line:{color:'1A2236',width:1}});
      s2.addShape(pptx.ShapeType.rect,{x,y,w:3.0,h:0.04,fill:{color:kpiColors[i%4]}});
      s2.addText(col,{x,y:y+0.1,w:3.0,h:0.3,fontSize:9,bold:true,color:'4A5578',align:'center',fontFace:'Calibri'});
      s2.addText(fmt(s),{x,y:y+0.38,w:3.0,h:0.7,fontSize:28,bold:true,color:'F0F4FF',align:'center',fontFace:'Calibri'});
      s2.addText(`avg ${fmt(s/vals.length)} · max ${fmt(Math.max(...vals))}`,{x,y:y+1.1,w:3.0,h:0.35,fontSize:9,color:'94A3C8',align:'center',fontFace:'Calibri'});
    });

    /* ─ SLIDES 3+: Charts ─ */
    const chartIds=Object.keys(chartInst).filter(k=>{
      if(tab==='overview') return k.startsWith('overview_');
      if(tab==='analysis') return k.startsWith('analysis_');
      return true;
    });

    for(let i=0;i<chartIds.length;i+=2){
      const s=pptx.addSlide();
      s.background={color:'0A0E1A'};
      s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:'100%',h:0.06,fill:{color:'38bdf8'}});

      for(let j=0;j<2&&i+j<chartIds.length;j++){
        const cid=chartIds[i+j];
        const cfg=chartInst[cid]?._cfg;
        if(!cfg) continue;
        const canvas=document.getElementById(cid);
        if(!canvas) continue;

        // Render chart as image
        const tmp=document.createElement('canvas');
        tmp.width=canvas.width; tmp.height=canvas.height;
        const ctx=tmp.getContext('2d');
        ctx.fillStyle='#111827'; ctx.fillRect(0,0,tmp.width,tmp.height);
        ctx.drawImage(canvas,0,0);
        const imgData=tmp.toDataURL('image/png');

        const col_x=j===0?0.3:6.7;
        // Chart title bar
        s.addShape(pptx.ShapeType.rect,{x:col_x,y:0.2,w:5.9,h:0.4,fill:{color:'111827'},line:{color:'1A2236',width:0.5}});
        s.addText(cfg.title,{x:col_x+0.1,y:0.22,w:5.7,h:0.36,fontSize:11,bold:true,color:'F0F4FF',fontFace:'Calibri',valign:'middle'});
        // Chart image
        s.addImage({data:imgData,x:col_x,y:0.65,w:5.9,h:3.5});

        // Small aggregation legend below
        const {labels,values}=aggregateData(cfg.x,cfg.y,cfg.agg,5);
        if(labels.length){
          const tblData=[labels.map(l=>({text:l,options:{fontSize:8,bold:false,color:'94A3C8',fontFace:'Calibri'}})),
                         values.map(v=>({text:fmt(v),options:{fontSize:8,bold:true,color:'F0F4FF',fontFace:'Calibri'}}))];
          try {
            s.addTable([tblData[0].map((c,ci)=>[c,tblData[1][ci]])].flat(1).reduce((acc,_,ci)=>{
              if(ci%2===0) acc.push([tblData[0][ci/2],tblData[1][ci/2]]); return acc;
            },[]), {x:col_x,y:4.2,w:5.9,h:0.5,border:{color:'1A2236'},fill:{color:'111827'}});
          } catch(e){}
        }
      }
      s.addText(`Namma BI · Slide ${3+Math.floor(i/2)} · ${new Date().toLocaleDateString()}`,
        {x:0.3,y:5.15,w:12.5,h:0.25,fontSize:8,color:'4A5578',fontFace:'Calibri'});
    }

    /* ─ FINAL SLIDE ─ */
    const sf=pptx.addSlide();
    sf.background={color:'0A0E1A'};
    sf.addShape(pptx.ShapeType.rect,{x:0,y:0,w:'100%',h:0.06,fill:{color:'38bdf8'}});
    sf.addText('Thank you',{x:1,y:1.8,w:11,h:1.0,fontSize:40,bold:true,color:'F0F4FF',align:'center',fontFace:'Calibri'});
    sf.addText('Generated by Namma BI',{x:1,y:2.9,w:11,h:0.5,fontSize:16,color:'38bdf8',align:'center',fontFace:'Calibri'});
    sf.addText(`${filtered.length.toLocaleString()} records · ${headers.length} fields · ${new Date().toLocaleDateString()}`,
      {x:1,y:3.5,w:11,h:0.4,fontSize:12,color:'4A5578',align:'center',fontFace:'Calibri'});

    await pptx.writeFile({fileName:'Namma_BI_Report.pptx'});
    showExportMsg('PowerPoint downloaded successfully!');
  } catch(e) {
    showExportMsg('PPT Error: '+e.message,'rose');
    console.error(e);
  }
}

/* ━━━ CSV EXPORT ━━━ */
function exportCSV(){
  const rows=[headers.join(','),...filtered.map(r=>headers.map(h=>`"${String(r[h]||'').replace(/"/g,'""')}"`).join(','))];
  const blob=new Blob([rows.join('\n')],{type:'text/csv'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='Namma_data.csv'; a.click();
}

/* ━━━ EXPORT MODAL ━━━ */
function showExport(){ document.getElementById('export-overlay').classList.add('show'); }
function closeOverlay(){ document.getElementById('export-overlay').classList.remove('show'); }
function showExportMsg(msg, color='green'){
  const el=document.getElementById('export-msg');
  el.textContent=msg; el.style.color=`var(--${color})`; el.style.display='';
  setTimeout(()=>el.style.display='none',4000);
}

/* ━━━ SAMPLE DATA ━━━ */
function loadSample(){
  const cats=['Electronics','Clothing','Food & Bev','Sports','Books','Home & Garden','Toys','Beauty'];
  const regions=['North','South','East','West','Central'];
  const mgrs=['Alice','Bob','Carol','David','Eva'];
  const months=['2024-01','2024-02','2024-03','2024-04','2024-05','2024-06','2024-07','2024-08','2024-09','2024-10','2024-11','2024-12'];
  const data=[];
  for(let i=0;i<400;i++){
    const cat=cats[i%cats.length];
    const region=regions[Math.floor(Math.random()*regions.length)];
    const month=months[Math.floor(Math.random()*months.length)];
    data.push({
      'Date':month+'-'+String(Math.floor(Math.random()*28)+1).padStart(2,'0'),
      'Category':cat,'Region':region,'Manager':mgrs[Math.floor(Math.random()*mgrs.length)],
      'Revenue':Math.round(1000+Math.random()*99000),
      'Units':Math.round(1+Math.random()*999),
      'Profit':Math.round(-2000+Math.random()*30000),
      'Cost':Math.round(500+Math.random()*20000),
      'Returns':Math.round(Math.random()*80),
      'Discount%':parseFloat((Math.random()*40).toFixed(1)),
    });
  }
  document.getElementById('file-pill-name').textContent='Sample Sales Data (400 rows)';
  document.getElementById('file-pill').style.display='flex';
  ingestData(data);
}
</script>
</body>
</html>

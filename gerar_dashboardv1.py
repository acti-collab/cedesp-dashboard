#!/usr/bin/env python3
"""
gerar_dashboard.py
==================
Script semanal de atualização do Dashboard CEDESP Dom Bosco Itaquera.

Uso:
    python gerar_dashboard.py                          # usa planilha padrão
    python gerar_dashboard.py "caminho/planilha.xlsx"  # especifica arquivo

Saída:
    dashboard_cursos_frequencia.html   (pronto para publicar no Google Sites)
    dashboard_resumo.html              (visão geral por unidade)
"""

import sys
import json
import os
from datetime import datetime, timezone, timedelta
import pandas as pd

# ── CONFIGURAÇÃO ──────────────────────────────────────────────────────────────
PLANILHA_PADRAO = "1º_Sem__-__2026.xlsx"   # altere se o nome mudar
SAIDA_CURSOS    = "dashboard_cursos_frequencia.html"
SAIDA_RESUMO    = "dashboard_resumo.html"
# ─────────────────────────────────────────────────────────────────────────────


def carregar_planilha(caminho):
    print(f"📂  Lendo planilha: {caminho}")
    if not os.path.exists(caminho):
        print(f"❌  Arquivo não encontrado: {caminho}")
        sys.exit(1)
    return pd.read_excel(caminho, sheet_name=None, header=None)


def extrair_totais(sheets):
    """Extrai resumo por unidade da aba TOTAIS."""
    df = sheets.get("TOTAIS")
    if df is None:
        return []

    unidades = []
    for idx in range(len(df)):
        row = df.iloc[idx]
        val0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if "CEDESP" in val0 and "TOTAL" not in val0.upper():
            try:
                meta  = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0
                matr  = float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
                freq  = float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0
                saldo = float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0
                unidades.append({
                    "unit": val0.strip(),
                    "meta": meta,
                    "matr": matr,
                    "freq": freq,
                    "saldo": saldo,
                })
            except (ValueError, TypeError):
                continue
    return unidades


def extrair_cursos(sheets):
    """Extrai dados de frequência por curso de cada aba CEDESP."""
    all_courses = []

    for i in range(1, 9):
        sheet_name = f"CEDESP {i}"
        df = sheets.get(sheet_name)
        if df is None:
            print(f"  ⚠️  Aba '{sheet_name}' não encontrada, pulando.")
            continue

        print(f"  📊  Processando {sheet_name}...", end=" ")

        # Identifica linhas de cabeçalho de período
        period_rows = {}
        for idx in range(len(df)):
            row = df.iloc[idx]
            for col_idx in range(len(row)):
                val = row.iloc[col_idx]
                if pd.notna(val) and isinstance(val, str):
                    v = val.strip().upper()
                    if "MANHÃ" in v or "MANHA" in v:
                        period_rows[idx] = "Manhã"; break
                    elif "TARDE" in v:
                        period_rows[idx] = "Tarde"; break
                    elif "NOITE" in v:
                        period_rows[idx] = "Noite"; break

        # Para cada período, extrai layout de colunas
        period_col_layouts = {}
        for pidx, period in period_rows.items():
            row = df.iloc[pidx]
            layout = {}
            for c in range(len(row)):
                val = row.iloc[c]
                if pd.notna(val) and isinstance(val, str):
                    v = val.strip().upper()
                    if "CURSO" in v and any(p in v for p in ["MANHÃ","TARDE","NOITE","MANHA"]):
                        layout["name_col"] = c
                    elif "EIXO" in v:
                        layout["eixo_col"] = c
                    elif "META" in v and "CONV" in v:
                        layout["meta_col"] = c
                    elif "INSERIDO" in v:
                        layout["ins_col"] = c
                    elif "MATR" in v:
                        layout["matr_col"] = c
                    elif "VAGA" in v:
                        layout["vagas_col"] = c
            period_col_layouts[pidx] = (period, layout)

        # Identifica colunas FREQ por período
        # Todas as colunas com cabeçalho FREQ, excluindo 30/01 (dia opcional)
        SKIP_DATES = {"30/01", "30/1"}
        freq_cols_by_period = {}
        for pidx in period_col_layouts:
            header   = df.iloc[pidx]
            date_row = df.iloc[pidx + 1] if pidx + 1 < len(df) else None
            cols = []
            for c in range(len(header)):
                if pd.notna(header.iloc[c]) and str(header.iloc[c]).strip().upper() == "FREQ":
                    date_val = str(date_row.iloc[c]).strip() if date_row is not None and pd.notna(date_row.iloc[c]) else ""
                    if date_val in SKIP_DATES:
                        continue  # ignora dia opcional 30/01
                    # Convert DD/MM → MM/DD for sortable string comparison
                    import re as _re
                    m_dv = _re.match(r'^(\d{1,2})/(\d{1,2})$', date_val)
                    sort_date = f"{int(m_dv.group(2)):02d}/{int(m_dv.group(1)):02d}" if m_dv else date_val
                    display_date = date_val  # keep original DD/MM for display
                    cols.append((c, sort_date, display_date))  # store (col_idx, sortable_MM/DD, display_DD/MM)
            freq_cols_by_period[pidx] = cols

        # Extrai linhas de curso
        period_list = sorted(period_col_layouts.keys())
        cursos_encontrados = 0

        for p_i, pidx in enumerate(period_list):
            period, layout = period_col_layouts[pidx]
            freq_cols = freq_cols_by_period[pidx]
            next_pidx = period_list[p_i + 1] if p_i + 1 < len(period_list) else len(df)

            name_col = layout.get("name_col", 3)
            ins_col  = layout.get("ins_col",  7)
            matr_col = layout.get("matr_col", 8)
            meta_col = layout.get("meta_col", 5)

            for ridx in range(pidx + 2, next_pidx):
                row = df.iloc[ridx]

                # Nome do curso
                course_name = None
                val = row.iloc[name_col] if name_col < len(row) else None
                if pd.notna(val) and isinstance(val, str):
                    v = val.strip()
                    skip_words = ["TOTAL","SALDO","PLANEJAMENTO","NÃO FEZ","ANTES DA","ELETROTÉCNICA",
                                  "GUIA PRONATEC","PARADA PEDAGÓGICA","CURSOS ABAIXO"]
                    if v and not any(s in v.upper() for s in skip_words) and len(v) > 2:
                        course_name = v

                if course_name is None:
                    continue

                # Métricas numéricas
                def safe_float(c):
                    if c >= len(row): return None
                    v = row.iloc[c]
                    return float(v) if pd.notna(v) and isinstance(v, (int, float)) else None

                meta      = safe_float(meta_col)
                matr      = safe_float(matr_col)
                inseridos = safe_float(ins_col)

                # Frequências diárias — toda coluna FREQ com valor numérico = 1 aula
                day_freqs = []
                last_date = None
                last_date_display = None
                daily = []  # list of (sort_date, display_date, freq)
                for col_idx, sort_date, display_date in freq_cols:
                    if col_idx < len(row):
                        v = row.iloc[col_idx]
                        if pd.notna(v) and isinstance(v, (int, float)):
                            day_freqs.append(float(v))
                            daily.append((sort_date, display_date, float(v)))
                            last_date = sort_date
                            last_date_display = display_date

                freq_avg    = round(sum(day_freqs) / len(day_freqs), 1) if day_freqs else 0
                n_classes   = len(day_freqs)
                last_freq   = int(day_freqs[-1]) if day_freqs else None

                # Weekly aggregates (for trend chart) — compact format
                from datetime import datetime, timedelta
                week_agg = {}
                for (sd, dd, fr) in daily:
                    try:
                        parts = sd.split('/')
                        month, day_n = int(parts[0]), int(parts[1])
                        dt = datetime(2026, month, day_n)
                        dow = dt.weekday()
                        week_start = dt - timedelta(days=dow)
                        wk = week_start.strftime('%d/%m')
                        if wk not in week_agg:
                            week_agg[wk] = [0, 0]
                        week_agg[wk][0] += fr
                        week_agg[wk][1] += 1
                    except Exception:
                        pass
                week_avgs = {wk: round(v[0]/v[1], 1) for wk, v in week_agg.items()}
                attend_rate = round(freq_avg / meta * 100, 1) if meta and meta > 0 and freq_avg > 0 else 0
                evasao      = round(matr - freq_avg, 1) if matr else None
                evasao_pct  = round(evasao / matr * 100, 1) if matr and evasao is not None else None

                centr     = int(row.iloc[0]) if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], (int, float)) else 0
                sentr     = int(row.iloc[1]) if pd.notna(row.iloc[1]) and isinstance(row.iloc[1], (int, float)) else 0
                dem_total = centr + sentr

                all_courses.append({
                    "unit":         sheet_name,
                    "period":       period,
                    "course":       course_name,
                    "meta":         meta,
                    "matr":         matr,
                    "inseridos":    inseridos,
                    "freq_avg":     freq_avg,
                    "n_classes":    n_classes,
                    "attend_rate":  attend_rate,
                    "evasao":       evasao,
                    "daily":        [[sd, dd, int(fr)] for sd, dd, fr in daily],
                    "last_freq":    last_freq,
                    "week_avgs":    week_avgs,
                    "last_date":    last_date,
                    "last_date_display": last_date_display,
                    "evasao_pct":   evasao_pct,
                    "centr":        centr,
                    "sentr":        sentr,
                    "dem_total":    dem_total,
                })
                cursos_encontrados += 1

        print(f"{cursos_encontrados} turmas")

    return all_courses


def gerar_html_cursos(cursos, data_atualizacao):
    """Gera o HTML do dashboard de frequência por curso."""

    dados_js = json.dumps(cursos, ensure_ascii=False, indent=2)

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Frequência por Curso — CEDESP Dom Bosco Itaquera</title>
<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<style>
:root {{
  --bg:#f5f2ee;--surface:#fff;--surface2:#ede9e3;--border:#d8d2c8;
  --text:#1a1612;--text-muted:#8a8279;--accent:#c84b31;--accent2:#2b5f8a;
  --green:#2d7a4f;--orange:#c97c1a;--yellow:#e8c547;--ink:#21438e;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--text);font-family:'Poppins',sans-serif;min-height:100vh}}
header{{background:var(--ink);color:#f5f2ee;padding:0 48px;display:flex;align-items:stretch;justify-content:space-between;height:80px;position:sticky;top:0;z-index:100}}
.header-left{{display:flex;align-items:center;gap:20px}}
.header-accent{{width:4px;height:40px;background:var(--accent);border-radius:2px}}
.header-title{{font-family:'Poppins',sans-serif;font-weight:800;font-size:17px;letter-spacing:-0.3px;line-height:1.25}}
.header-sub{{font-family:'JetBrains Mono',monospace;font-size:10px;color:rgba(245,242,238,.45);margin-top:3px;letter-spacing:.5px}}
.header-right{{display:flex;align-items:center;gap:32px}}
.header-stat{{text-align:right}}
.hs-val{{font-family:'Poppins',sans-serif;font-size:22px;font-weight:700;letter-spacing:-.5px;color:var(--yellow)}}
.hs-label{{font-family:'JetBrains Mono',monospace;font-size:10px;color:rgba(245,242,238,.45);letter-spacing:.5px}}
.controls{{background:var(--surface);border-bottom:1px solid var(--border);padding:14px 48px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;position:sticky;top:80px;z-index:99}}
.filter-label{{font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--text-muted);letter-spacing:.5px;text-transform:uppercase}}
.filter-btn{{display:inline-flex;align-items:center;padding:5px 14px;border-radius:2px;border:1px solid var(--border);background:transparent;color:var(--text-muted);font-family:'Poppins',sans-serif;font-size:12px;cursor:pointer;transition:all .15s}}
.filter-btn:hover{{border-color:var(--ink);color:var(--ink)}}
.filter-btn.active{{background:var(--ink);color:var(--bg);border-color:var(--ink)}}
.filter-btn.manha.active{{background:var(--orange);border-color:var(--orange);color:#fff}}
.filter-btn.tarde.active{{background:var(--accent2);border-color:var(--accent2);color:#fff}}
.filter-btn.noite.active{{background:#4a3a7a;border-color:#4a3a7a;color:#fff}}
.sort-select{{padding:5px 28px 5px 12px;border:1px solid var(--border);background:transparent;font-family:'Poppins',sans-serif;font-size:12px;color:var(--text);border-radius:2px;cursor:pointer;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6' viewBox='0 0 10 6'%3E%3Cpath d='M1 1l4 4 4-4' stroke='%238a8279' stroke-width='1.5' fill='none'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 8px center}}
.controls-spacer{{flex:1}}
.search-wrap{{position:relative}}
.search-input{{padding:6px 12px 6px 32px;border:1px solid var(--border);background:var(--bg);font-family:'Poppins',sans-serif;font-size:12px;color:var(--text);border-radius:2px;width:220px;outline:none}}
.search-input:focus{{border-color:var(--ink)}}
.search-icon{{position:absolute;left:10px;top:50%;transform:translateY(-50%);color:var(--text-muted);font-size:12px}}
main{{padding:32px 48px 60px;max-width:1440px;margin:0 auto}}
.kpi-row{{display:grid;grid-template-columns:repeat(4,1fr);gap:1px;background:var(--border);border:1px solid var(--border);margin-bottom:32px}}
.kpi-cell{{background:var(--surface);padding:24px 28px}}
.kpi-label{{font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--text-muted);letter-spacing:.8px;text-transform:uppercase;margin-bottom:10px}}
.kpi-val{{font-family:'Poppins',sans-serif;font-size:40px;font-weight:800;letter-spacing:-2px;line-height:1;margin-bottom:4px}}
.kpi-desc{{font-size:12px;color:var(--text-muted)}}
.section-row{{display:flex;align-items:baseline;gap:12px;margin-bottom:20px;margin-top:40px}}
.section-row:first-child{{margin-top:0}}
.section-title{{font-family:'Poppins',sans-serif;font-size:13px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--text-muted)}}
.section-rule{{flex:1;height:1px;background:var(--border)}}
.section-count{{font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--text-muted)}}
.charts-grid{{display:grid;grid-template-columns:1.6fr 1fr;gap:20px;margin-bottom:20px}}
.chart-box{{background:var(--surface);border:1px solid var(--border);padding:24px}}
.chart-box-title{{font-family:'Poppins',sans-serif;font-size:13px;font-weight:700;letter-spacing:-.2px;margin-bottom:3px}}
.chart-box-sub{{font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--text-muted);margin-bottom:20px}}
.course-table-wrap{{background:var(--surface);border:1px solid var(--border);overflow:hidden}}
table{{width:100%;border-collapse:collapse}}
thead th{{padding:10px 16px;text-align:left;font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:.8px;text-transform:uppercase;color:var(--text-muted);background:var(--surface2);border-bottom:1px solid var(--border);white-space:nowrap;cursor:pointer;user-select:none;transition:color .15s}}
thead th:hover{{color:var(--text)}}
thead th.sort-active{{color:var(--accent2)}}
thead th.right{{text-align:right}}
thead th .sort-arrow{{margin-left:4px;opacity:.5;font-size:10px}}
thead th.sort-active .sort-arrow{{opacity:1;color:var(--accent2)}}
tbody tr{{border-bottom:1px solid var(--border);transition:background .1s}}
tbody tr:last-child{{border-bottom:none}}
tbody tr:hover{{background:var(--surface2)}}
td{{padding:11px 16px;font-size:12.5px}}
td.mono{{font-family:'JetBrains Mono',monospace;font-size:11.5px}}
td.right{{text-align:right}}
.course-name-cell{{max-width:240px;font-weight:600;font-size:12px;line-height:1.3}}
.unit-chip{{display:inline-block;padding:2px 8px;border-radius:2px;font-family:'JetBrains Mono',monospace;font-size:10px;font-weight:500;background:var(--surface2);color:var(--text-muted);white-space:nowrap}}
.period-dot{{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:6px}}
.period-dot.manhã{{background:var(--orange)}}.period-dot.tarde{{background:var(--accent2)}}.period-dot.noite{{background:#4a3a7a}}
.rate-cell{{min-width:140px}}
.rate-wrap{{display:flex;align-items:center;gap:8px}}
.rate-bar{{flex:1;height:5px;background:var(--surface2);overflow:hidden;min-width:60px}}
.rate-fill{{height:100%}}
.rate-pct{{font-family:'JetBrains Mono',monospace;font-size:11px;min-width:40px;text-align:right;font-weight:500}}
.status-flag{{display:inline-flex;align-items:center;gap:5px;font-size:11px;font-family:'JetBrains Mono',monospace;padding:2px 8px;border:1px solid;border-radius:2px}}
.status-flag.high{{border-color:#2d7a4f;color:#2d7a4f;background:rgba(45,122,79,.06)}}
.status-flag.mid{{border-color:#c97c1a;color:#c97c1a;background:rgba(201,124,26,.06)}}
.status-flag.low{{border-color:var(--accent);color:var(--accent);background:rgba(200,75,49,.06)}}
.empty-state{{text-align:center;padding:48px;color:var(--text-muted);font-size:13px}}
.update-badge{{font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--text-muted);background:var(--surface2);border:1px solid var(--border);padding:3px 10px;border-radius:2px}}
footer{{border-top:1px solid var(--border);padding:20px 48px;display:flex;justify-content:space-between;align-items:center;margin-top:48px}}
.footer-text{{font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--text-muted)}}
@media(max-width:900px){{header,.controls,main{{padding-left:20px;padding-right:20px}}.charts-grid{{grid-template-columns:1fr}}.kpi-row{{grid-template-columns:1fr 1fr}}}}
</style>
</head>
<body>

<header>
  <div class="header-left">
    <div class="header-accent"></div>
    <div>
      <div class="header-title">Frequência por Curso</div>
      <div class="header-sub">CEDESP DOM BOSCO ITAQUERA · 1º SEMESTRE 2026</div>
    </div>
  </div>
  <div class="header-right">
    <div class="header-stat"><div class="hs-val" id="hdr-cursos">—</div><div class="hs-label">CURSOS</div></div>
    <div class="header-stat"><div class="hs-val" id="hdr-avg">—</div><div class="hs-label">FREQ. MÉDIA</div></div>
    <div class="header-stat"><div class="hs-val" id="hdr-ins">—</div><div class="hs-label">INSERIDOS</div></div>
    <div class="header-stat"><div class="hs-val" id="hdr-matr">—</div><div class="hs-label">MATRÍCULAS</div></div>
  </div>
</header>

<div class="controls">
  <div style="display:flex;align-items:center;gap:6px">
    <span class="filter-label">Unidade:</span>
    <button class="filter-btn active" data-filter="unit" data-value="all">Todas</button>
    {"".join(f'<button class="filter-btn" data-filter="unit" data-value="CEDESP {n}">C{n}</button>' for n in range(1,9))}
  </div>
  <div style="display:flex;align-items:center;gap:6px">
    <span class="filter-label">Horário:</span>
    <button class="filter-btn active" data-filter="period" data-value="all">Todos</button>
    <button class="filter-btn manha" data-filter="period" data-value="Manhã">☀ Manhã</button>
    <button class="filter-btn tarde" data-filter="period" data-value="Tarde">🌤 Tarde</button>
    <button class="filter-btn noite" data-filter="period" data-value="Noite">🌙 Noite</button>
  </div>
  <div class="controls-spacer"></div>
  <span class="update-badge">📅 Atualizado em: {data_atualizacao}</span>
  <div style="display:flex;align-items:center;gap:6px">

  </div>
  <div class="search-wrap">
    <span class="search-icon">⌕</span>
    <input type="text" class="search-input" id="searchInput" placeholder="Buscar curso...">
  </div>
</div>

<main>

<div class="kpi-row">
  <div class="kpi-cell"><div class="kpi-label">Cursos ativos</div><div class="kpi-val" id="kpi-cursos">—</div><div class="kpi-desc">turmas com frequência registrada</div></div>
  <div class="kpi-cell"><div class="kpi-label">Taxa de frequência média</div><div class="kpi-val" id="kpi-rate" style="color:var(--accent2)">—</div><div class="kpi-desc">freq. média ÷ meta do curso</div></div>

  <div class="kpi-cell"><div class="kpi-label">Cursos com taxa &lt;80%</div><div class="kpi-val" id="kpi-low" style="color:var(--accent)">—</div><div class="kpi-desc">requerem atenção</div></div>
  <div class="kpi-cell"><div class="kpi-label">Evasão média</div><div class="kpi-val" id="kpi-evasao" style="color:#c84b31">—</div><div class="kpi-desc">matriculados que não frequentam</div></div>
</div>

<div id="snapshot-panel" style="display:none;background:linear-gradient(135deg,#1a2340 0%,#21438e 100%);border-radius:10px;padding:20px 28px;margin-bottom:24px;color:#fff;position:relative;overflow:hidden">
  <div style="position:absolute;top:-30px;right:-30px;width:140px;height:140px;background:rgba(255,255,255,.04);border-radius:50%"></div>
  <div style="position:absolute;bottom:-40px;right:60px;width:90px;height:90px;background:rgba(255,255,255,.03);border-radius:50%"></div>
  <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px">
    <span style="font-size:14px">📅</span>
    <span style="font-size:11px;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;color:rgba(255,255,255,.6)">Dados do Dia</span>
    <span id="snapshot-date" style="font-size:11px;color:rgba(255,255,255,.4);margin-left:4px"></span>
  </div>
  <div style="display:flex;gap:40px;flex-wrap:wrap;align-items:flex-end">
    <div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:48px;font-weight:800;line-height:1;letter-spacing:-2px" id="snap-last-freq">—</div>
      <div style="font-size:12px;color:rgba(255,255,255,.65);margin-top:6px">alunos presentes na última aula registrada</div>
    </div>
    <div style="width:1px;height:60px;background:rgba(255,255,255,.15);align-self:center"></div>
    <div>
      <div style="display:flex;align-items:baseline;gap:6px">
        <span style="font-family:'JetBrains Mono',monospace;font-size:48px;font-weight:800;line-height:1;letter-spacing:-2px" id="snap-max-aulas">—</span>
        <span style="font-size:14px;color:rgba(255,255,255,.5)">aulas</span>
      </div>
      <div style="font-size:12px;color:rgba(255,255,255,.65);margin-top:6px">máximo de chamadas registradas</div>
    </div>
    <div style="width:1px;height:60px;background:rgba(255,255,255,.15);align-self:center"></div>
    <div>
      <div style="display:flex;align-items:baseline;gap:6px">
        <span style="font-family:'JetBrains Mono',monospace;font-size:48px;font-weight:800;line-height:1;letter-spacing:-2px;color:#f59e0b" id="snap-atrasados">—</span>
        <span style="font-size:14px;color:rgba(255,255,255,.5)">cursos</span>
      </div>
      <div style="font-size:12px;color:rgba(255,255,255,.65);margin-top:6px">com menos aulas que o máximo <span style="color:#f59e0b;font-weight:600">(possível chamada em falta)</span></div>
    </div>
  </div>
</div>

<div class="section-row">
  <div class="section-title">Tabela Completa</div>
  <div class="section-rule"></div>
  <span class="section-count" id="resultCount">—</span>
</div>

<div class="course-table-wrap">
  <table id="courseTable">
    <thead>
      <tr>
        <th data-sort="course">Curso <span class="sort-arrow">↕</span></th>
        <th data-sort="unit">Unidade <span class="sort-arrow">↕</span></th>
        <th data-sort="period">Período <span class="sort-arrow">↕</span></th>
        <th class="right" data-sort="inseridos">Inseridos <span class="sort-arrow">↕</span></th>
        <th class="right" data-sort="matr">Matr. <span class="sort-arrow">↕</span></th>
        <th class="right" data-sort="freq_avg">Freq. Média/dia <span class="sort-arrow">↕</span></th>
        <th class="right" data-sort="n_classes">Aulas <span class="sort-arrow">↕</span></th>
        <th data-sort="attend_rate">Taxa s/ Meta <span class="sort-arrow">↕</span></th>
        <th class="right" data-sort="evasao_pct">Evasão <span class="sort-arrow">↕</span></th>
        <th class="right" data-sort="dem_total">Demanda <span class="sort-arrow">↕</span></th>
        <th data-sort="status">Status <span class="sort-arrow">↕</span></th>
      </tr>
    </thead>
    <tbody id="courseTableBody"></tbody>
  </table>
  <div class="empty-state" id="emptyState" style="display:none">Nenhum curso encontrado.</div>
</div>

<div class="section-row"><div class="section-title">Visão Gráfica</div><div class="section-rule"></div></div>

<div class="charts-grid">
  <div class="chart-box">
    <div class="chart-box-title">Taxa de Frequência por Curso</div>
    <div class="chart-box-sub">Frequência diária média ÷ meta do curso — top 20 cursos</div>
    <canvas id="courseBarChart" height="320"></canvas>
  </div>
  <div class="chart-box">
    <div class="chart-box-title">Distribuição de Taxas</div>
    <div class="chart-box-sub">Agrupamento por faixa de frequência</div>
    <canvas id="distChart" height="320"></canvas>
  </div>
</div>
<div class="charts-grid" style="grid-template-columns:1fr">
  <div class="chart-box">
    <div class="chart-box-title">Tendência de Frequência Semanal</div>
    <div class="chart-box-sub">Evolução da presença média semana a semana ao longo do semestre</div>
    <canvas id="trendChart" height="110"></canvas>
  </div>
</div>

<div class="charts-grid" style="grid-template-columns:1fr">
  <div class="chart-box">
    <div class="chart-box-title">Evasão por Curso — Top 20</div>
    <div class="chart-box-sub">Cursos com maior evasão · matriculados que não frequentam</div>
    <canvas id="evasaoChart" height="160"></canvas>
  </div>
</div>

</main>

<footer>
  <div class="footer-text">CEDESP Dom Bosco Itaquera · Dashboard gerado automaticamente</div>
  <div class="footer-text">Atualizado em {data_atualizacao}</div>
</footer>

<script>
const RAW = {dados_js};

let filterUnit='all', filterPeriod='all', sortMode='attend_rate', sortDir=-1, searchQuery='';

function getColor(r){{ return r>=100?'#2d7a4f':r>=80?'#c97c1a':'#c84b31'; }}
function getStatus(r){{ return r>=100?{{cls:'high',label:'✓ Meta atingida'}}:r>=80?{{cls:'mid',label:'⚠ Atenção média'}}:{{cls:'low',label:'✗ Atenção imediata'}}; }}
function periodColor(p){{ return p==='Manhã'?'#c97c1a':p==='Tarde'?'#2b5f8a':'#4a3a7a'; }}

function getFiltered(){{
  return RAW.filter(d=>{{
    if(filterUnit!=='all'&&d.unit!==filterUnit)return false;
    if(filterPeriod!=='all'&&d.period!==filterPeriod)return false;
    if(searchQuery&&!d.course.toLowerCase().includes(searchQuery)&&!d.unit.toLowerCase().includes(searchQuery))return false;
    return true;
  }});
}}

function getSorted(data){{
  return [...data].sort((a,b)=>{{
    const statusOrder={{high:0,mid:1,low:2}};
    let av=a[sortMode], bv=b[sortMode];
    if(sortMode==='status'){{av=statusOrder[getStatus(a.attend_rate).cls]??3;bv=statusOrder[getStatus(b.attend_rate).cls]??3;}}
    if(typeof av==='string')return av.localeCompare(bv)*sortDir;
    av=av??-Infinity; bv=bv??-Infinity;
    return(av-bv)*sortDir;
  }});
}}

Chart.defaults.color='#8a8279';
Chart.defaults.font.family="'Poppins',sans-serif";
Chart.defaults.borderColor='#d8d2c8';
let charts={{}};

function buildCharts(data){{
  Object.values(charts).forEach(c=>c&&c.destroy());
  const sorted=[...data].sort((a,b)=>a.attend_rate-b.attend_rate).slice(0,20);

  charts.bar=new Chart(document.getElementById('courseBarChart'),{{
    type:'bar',
    data:{{
      labels:sorted.map(d=>d.course.length>32?d.course.substring(0,30)+'…':d.course),
      datasets:[{{label:'Taxa %',data:sorted.map(d=>d.attend_rate),backgroundColor:sorted.map(d=>getColor(d.attend_rate)+'cc'),borderColor:sorted.map(d=>getColor(d.attend_rate)),borderWidth:1,borderRadius:0}}]
    }},
    options:{{indexAxis:'y',responsive:true,plugins:{{legend:{{display:false}},tooltip:{{backgroundColor:'#1a1612',borderColor:'#d8d2c8',borderWidth:1,callbacks:{{label:ctx=>{{const d=sorted[ctx.dataIndex];return[` ${{ctx.raw.toFixed(1)}}%  —  ${{d.unit}} · ${{d.period}}`,` Presença média: ${{d.freq_avg}} alunos/dia`];}}}}}}}},scales:{{x:{{min:0,max:120,grid:{{color:'rgba(216,210,200,.4)'}},ticks:{{font:{{size:10}},callback:v=>v+'%'}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:10}}}}}}}}}}
  }});

  const bands=[
    {{label:'≥ 100% (Meta atingida)',min:100,max:Infinity,color:'#2d7a4f'}},
    {{label:'80–99.9% (Atenção média)',min:80,max:100,color:'#c97c1a'}},
    {{label:'< 80% (Atenção imediata)',min:0,max:80,color:'#c84b31'}},
  ];
  charts.dist=new Chart(document.getElementById('distChart'),{{
    type:'doughnut',
    data:{{labels:bands.map(b=>b.label),datasets:[{{data:bands.map(b=>data.filter(d=>d.attend_rate>=b.min&&d.attend_rate<b.max).length),backgroundColor:bands.map(b=>b.color+'cc'),borderColor:bands.map(b=>b.color),borderWidth:2,hoverOffset:8}}]}},
    options:{{responsive:true,cutout:'58%',plugins:{{legend:{{position:'right',labels:{{boxWidth:12,padding:12,font:{{size:11}},usePointStyle:true}}}},tooltip:{{backgroundColor:'#1a1612',borderColor:'#d8d2c8',borderWidth:1}}}}}}
  }});

  // Tendência de Frequência Semanal
  const weekMap={{}};
  data.forEach(d=>{{
    if(!d.daily) return;
    d.daily.forEach(([sd, dd, fr])=>{{
      if(!sd) return;
      const parts=sd.split('/');
      if(parts.length<2) return;
      const month=parseInt(parts[0]), day=parseInt(parts[1]);
      const dt=new Date(2026, month-1, day);
      const dow=dt.getDay()||7;
      const weekStart=new Date(dt); weekStart.setDate(dt.getDate()-dow+1);
      // Label: "Sem 1 Fev", "Sem 2 Fev" etc.
      const MONTHS=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
      const m=weekStart.getMonth(); // 0-based
      // Which week of the month? (1st day of month = week 1)
      const firstDayOfMonth=new Date(weekStart.getFullYear(), m, 1);
      const weekOfMonth=Math.ceil((weekStart.getDate()+firstDayOfMonth.getDay())/7);
      const wk=`Sem ${{weekOfMonth}} ${{MONTHS[m]}}`;
      const sk=`${{String(weekStart.getMonth()+1).padStart(2,'0')}}/${{String(weekStart.getDate()).padStart(2,'0')}}`;
      if(!weekMap[wk]) weekMap[wk]={{sum:0,n:0,sortKey:sk}};
      weekMap[wk].sum+=fr; weekMap[wk].n++;
    }});
  }});
  // Sort by first date seen in each week bucket (stored as sortKey)
  const weeks=Object.keys(weekMap).sort((a,b)=>{{
    const ka=weekMap[a].sortKey||''; const kb=weekMap[b].sortKey||'';
    return ka<kb?-1:ka>kb?1:0;
  }});
  const weekAvgs=weeks.map(w=>+(weekMap[w].sum/weekMap[w].n).toFixed(1));
  const weekMeta=weeks.map(()=>{{
    const tm=data.reduce((s,d)=>s+(d.meta||0),0);
    return +((tm||0)/(data.length||1)).toFixed(1);
  }});
  const trendColor=weekAvgs.length<2||weekAvgs[weekAvgs.length-1]>=weekAvgs[0]?'#2d7a4f':'#c84b31';
  charts.unit=new Chart(document.getElementById('trendChart'),{{
    type:'line',
    data:{{labels:weeks,datasets:[
      {{label:'Presença média',data:weekAvgs,borderColor:trendColor,backgroundColor:trendColor+'22',borderWidth:2.5,pointRadius:4,pointHoverRadius:6,tension:0.35,fill:true,spanGaps:true}},
      {{label:'Meta média',data:weekMeta,borderColor:'#c97c1a',backgroundColor:'transparent',borderWidth:1.5,borderDash:[6,4],pointRadius:0,tension:0,fill:false}}
    ]}},
    options:{{responsive:true,plugins:{{
      legend:{{position:'top',labels:{{boxWidth:12,padding:14,font:{{size:10}}}}}},
      tooltip:{{backgroundColor:'#1a1612',borderColor:'#d8d2c8',borderWidth:1,callbacks:{{
        label:ctx=>ctx.datasetIndex===0?` Presença média: ${{ctx.raw}} alunos/dia`:` Meta média: ${{ctx.raw}} alunos/dia`
      }}}}
    }},scales:{{
      x:{{grid:{{display:false}},ticks:{{font:{{size:10}}}}}},
      y:{{
          grid:{{color:'rgba(216,210,200,.4)'}},
          ticks:{{font:{{size:10}}}},
          min:weekAvgs.length?Math.floor(Math.min(...weekAvgs,...weekMeta)*0.95):0,
          max:weekAvgs.length?Math.ceil(Math.max(...weekAvgs,...weekMeta)*1.05):100,
        }}
    }}}}
  }});
}}

// ── EVASION CHART ──
let evasaoChart;
function buildEvasaoChart(data){{
  if(evasaoChart)evasaoChart.destroy();
  const withEvasao=data.filter(d=>d.matr&&d.evasao!=null).sort((a,b)=>b.evasao_pct-a.evasao_pct).slice(0,20);
  const labels=withEvasao.map(d=>{{const n=d.course.length>30?d.course.substring(0,28)+'…':d.course;return n+' ('+d.unit.replace('CEDESP ','C')+')';}});
  evasaoChart=new Chart(document.getElementById('evasaoChart'),{{
    type:'bar',
    data:{{labels,datasets:[{{
      label:'Evasão (alunos)',
      data:withEvasao.map(d=>d.evasao),
      backgroundColor:withEvasao.map(d=>d.evasao_pct>30?'rgba(200,75,49,.85)':d.evasao_pct>15?'rgba(201,124,26,.85)':'rgba(125,133,144,.6)'),
      borderColor:withEvasao.map(d=>d.evasao_pct>30?'#c84b31':d.evasao_pct>15?'#c97c1a':'#7d8590'),
      borderWidth:1,borderRadius:0
    }}]}},
    options:{{
      indexAxis:'y',responsive:true,
      plugins:{{
        legend:{{display:false}},
        tooltip:{{backgroundColor:'#1a1612',borderColor:'#d8d2c8',borderWidth:1,
          callbacks:{{label:ctx=>{{const d=withEvasao[ctx.dataIndex];return` ${{d.evasao.toFixed(0)}} alunos evadidos — ${{d.evasao_pct.toFixed(1)}}% · ${{d.period}}`;}}}}
        }}
      }},
      scales:{{
        x:{{grid:{{color:'rgba(216,210,200,.3)'}},ticks:{{font:{{size:10}}}}}},
        y:{{grid:{{display:false}},ticks:{{font:{{size:10}}}}}}
      }}
    }}
  }});
}}

function buildTable(data){{
  const tbody=document.getElementById('courseTableBody');
  tbody.innerHTML='';
  const sorted=getSorted(data);
  document.getElementById('resultCount').textContent=sorted.length+' cursos';
  document.getElementById('emptyState').style.display=sorted.length?'none':'block';
  sorted.forEach(d=>{{
    const st=getStatus(d.attend_rate);
    const cap=Math.min(d.attend_rate,100);
    const fc=getColor(d.attend_rate);
    const pd=d.period.toLowerCase().replace('ã','a');
    const tr=document.createElement('tr');
    tr.innerHTML=`
      <td><div class="course-name-cell">${{d.course}}</div></td>
      <td><span class="unit-chip">${{d.unit.replace('CEDESP ','C')}}</span></td>
      <td class="mono"><span class="period-dot ${{pd}}"></span>${{d.period}}</td>
      <td class="right mono">${{d.inseridos||'—'}}</td>
      <td class="right mono">${{d.matr||'—'}}</td>
      <td class="right mono">${{d.freq_avg.toFixed(1)}}</td>
      <td class="right mono">${{d.n_classes}}</td>
      <td class="rate-cell"><div class="rate-wrap"><div class="rate-bar"><div class="rate-fill" style="width:${{cap}}%;background:${{fc}}"></div></div><div class="rate-pct" style="color:${{fc}}">${{d.attend_rate.toFixed(1)}}%</div></div></td>
      <td class="right mono">${{d.evasao_pct!=null?'<span style="color:'+(d.evasao_pct>30?'#c84b31':d.evasao_pct>15?'#c97c1a':'#7d8590')+'">'+d.evasao_pct.toFixed(1)+'%</span>':'—'}}</td>
      <td class="right">${{(()=>{{
        const c=d.centr,s=d.sentr,t=d.dem_total;
        if(!t)return'<span style="color:var(--text-muted)">—</span>';
        const tag=c>0?'<span style="display:inline-flex;gap:3px;align-items:center">'
          +'<span title="Com entrevista (aguardando vaga)" style="font-size:10px;padding:1px 5px;border-radius:2px;background:rgba(45,122,79,.15);color:#2d7a4f;font-family:JetBrains Mono,monospace">'+c+'✓</span>'
          +(s>0?'<span title="Sem entrevista (só ficha)" style="font-size:10px;padding:1px 5px;border-radius:2px;background:rgba(88,166,255,.12);color:#58a6ff;font-family:JetBrains Mono,monospace">'+s+'~</span>':'')
          +'</span>'
          :'<span title="Só ficha, sem entrevista" style="font-size:10px;padding:1px 5px;border-radius:2px;background:rgba(88,166,255,.12);color:#58a6ff;font-family:JetBrains Mono,monospace">'+s+'~</span>';
        return tag;
      }})()}}</td>
      <td><span class="status-flag ${{st.cls}}">${{st.label}}</span></td>
    `;
    tbody.appendChild(tr);
  }});
}}

function updateKPIs(data){{
  document.getElementById('kpi-cursos').textContent=data.length;
  const tm=data.reduce((s,d)=>s+(d.meta||0),0);
  const wr=tm>0?(data.reduce((s,d)=>s+d.attend_rate*(d.meta||0),0)/tm).toFixed(1)+'%':'—';
  document.getElementById('kpi-rate').textContent=wr;

  document.getElementById('kpi-low').textContent=data.filter(d=>d.attend_rate<80).length;
  const withEv=data.filter(d=>d.evasao_pct!=null);
  const avgEv=withEv.length?(withEv.reduce((s,d)=>s+d.evasao_pct,0)/withEv.length).toFixed(1):'—';
  document.getElementById('kpi-evasao').textContent=avgEv!=='—'?avgEv+'%':'—';
  document.getElementById('hdr-cursos').textContent=data.length;
  document.getElementById('hdr-avg').textContent=wr;
  document.getElementById('hdr-matr').textContent=data.reduce((s,d)=>s+(d.matr||0),0);
  document.getElementById('hdr-ins').textContent=data.reduce((s,d)=>s+(d.inseridos||0),0);
  // Snapshot do Dia
  const maxAulas=data.reduce((m,d)=>Math.max(m,d.n_classes||0),0);
  // Find the most recent date across all courses, sum only those on that date
  // Per-unit last day: find latest date per unit, sum last_freq on that date
  const unitLatest={{}};
  data.forEach(d=>{{if(d.last_date&&(!unitLatest[d.unit]||d.last_date>unitLatest[d.unit]))unitLatest[d.unit]=d.last_date;}});
  const totalLastFreq=data.reduce((s,d)=>d.last_date&&d.last_date===unitLatest[d.unit]?s+(d.last_freq||0):s,0);
  // For the date label, show the most recent date among all units
  const latestDate=Object.values(unitLatest).reduce((mx,d)=>d>mx?d:mx,'');
  const atrasados=data.filter(d=>(d.n_classes||0)<maxAulas).length;
  const panel=document.getElementById('snapshot-panel');
  if(maxAulas>0||totalLastFreq>0){{
    panel.style.display='block';
    const fmtDate=d=>{{if(!d)return'';const p=d.match(/(\\d+)\\/(\\d+)/);if(p)return d;return d;}};
  const latestDisplay=data.find(d=>d.last_date===latestDate)?.last_date_display||'';
  document.getElementById('snapshot-date').textContent=latestDisplay?'· '+latestDisplay:'';
  document.getElementById('snap-last-freq').textContent=totalLastFreq||'—';
    document.getElementById('snap-max-aulas').textContent=maxAulas||'—';
    document.getElementById('snap-atrasados').textContent=atrasados;
  }} else {{
    panel.style.display='none';
  }}
}}

function update(){{
  const f=getFiltered();
  updateKPIs(f);
  buildCharts(f);
  buildEvasaoChart(f);
  buildTable(f);
}}

document.querySelectorAll('[data-filter="unit"]').forEach(btn=>{{
  btn.addEventListener('click',()=>{{
    document.querySelectorAll('[data-filter="unit"]').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');filterUnit=btn.dataset.value;update();
  }});
}});
document.querySelectorAll('[data-filter="period"]').forEach(btn=>{{
  btn.addEventListener('click',()=>{{
    document.querySelectorAll('[data-filter="period"]').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');filterPeriod=btn.dataset.value;update();
  }});
}});
document.querySelectorAll('#courseTable thead th[data-sort]').forEach(th=>{{
  th.addEventListener('click',()=>{{
    const col=th.dataset.sort;
    if(sortMode===col){{sortDir*=-1;}}else{{sortMode=col;sortDir=col==='attend_rate'?-1:1;}}
    document.querySelectorAll('#courseTable thead th').forEach(t=>{{
      t.classList.remove('sort-active');
      const arr=t.querySelector('.sort-arrow');
      if(arr)arr.textContent='↕';
    }});
    th.classList.add('sort-active');
    const arr=th.querySelector('.sort-arrow');
    if(arr)arr.textContent=sortDir===-1?'↓':'↑';
    buildTable(getFiltered());
  }});
}});
(()=>{{
  const th=document.querySelector('#courseTable thead th[data-sort="attend_rate"]');
  if(th){{th.classList.add('sort-active');const arr=th.querySelector('.sort-arrow');if(arr)arr.textContent='↓';}}
}})();
document.getElementById('searchInput').addEventListener('input',e=>{{searchQuery=e.target.value.toLowerCase().trim();update();}});

update();
</script>
</body>
</html>"""

    return html


def main():
    caminho = sys.argv[1] if len(sys.argv) > 1 else PLANILHA_PADRAO
    data_atualizacao = datetime.now(timezone(timedelta(hours=-3))).strftime("%d/%m/%Y às %H:%M")

    print(f"\n{'='*55}")
    print(f"  Dashboard CEDESP — Gerador Automático")
    print(f"  {data_atualizacao}")
    print(f"{'='*55}\n")

    sheets  = carregar_planilha(caminho)
    totais  = extrair_totais(sheets)
    cursos  = extrair_cursos(sheets)

    # Filtra apenas cursos com dados válidos
    cursos_validos = [c for c in cursos if c["attend_rate"] > 0 and c["matr"] and c["matr"] > 0]

    print(f"\n✅  {len(cursos_validos)} turmas com frequência registrada")
    print(f"✅  {len(totais)} unidades no resumo\n")

    # Gera dashboard de cursos
    html_cursos = gerar_html_cursos(cursos_validos, data_atualizacao)
    with open(SAIDA_CURSOS, "w", encoding="utf-8") as f:
        f.write(html_cursos)
    print(f"📄  Gerado: {SAIDA_CURSOS}")

    print(f"\n{'='*55}")
    print(f"  ✅  Concluído! Arquivos prontos para publicar.")
    print(f"{'='*55}\n")


if __name__ == "__main__":
    main()

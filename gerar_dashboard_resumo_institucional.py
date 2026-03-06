#!/usr/bin/env python3
"""
gerar_dashboard_resumo.py
=========================
Gera o Dashboard de Resumo por Unidade — CEDESP Dom Bosco Itaquera.

Uso:
    python gerar_dashboard_resumo.py                          # usa planilha padrão
    python gerar_dashboard_resumo.py "caminho/planilha.xlsx"  # especifica arquivo

Saída:
    dashboard_cedesp_2026.html
"""

import sys
import os
import json
from datetime import datetime
import pandas as pd

# ── CONFIGURAÇÃO ──────────────────────────────────────────────────────────────
PLANILHA_PADRAO = "1º_Sem__-__2026.xlsx"
SAIDA           = "dashboard_cedesp_2026.html"
SKIP_DATES      = {"30/01", "30/1"}   # dia opcional — ignorado no cálculo
# ─────────────────────────────────────────────────────────────────────────────


def carregar_planilha(caminho):
    print(f"📂  Lendo planilha: {caminho}")
    if not os.path.exists(caminho):
        print(f"❌  Arquivo não encontrado: {caminho}")
        sys.exit(1)
    return pd.read_excel(caminho, sheet_name=None, header=None)


def extrair_unidades(sheets):
    """
    Extrai meta, matrículas, frequência e saldo por unidade.

    - meta / matr  → linha TOTAL GERAL de cada aba CEDESP
    - freq         → soma da última aula com presença registrada
                     em todos os períodos (Manhã + Tarde + Noite)
    - saldo        → freq − meta
    """
    results = []

    for i in range(1, 9):
        sheet_name = f"CEDESP {i}"
        df = sheets.get(sheet_name)
        if df is None:
            print(f"  ⚠️  Aba '{sheet_name}' não encontrada, pulando.")
            continue

        print(f"  📊  Processando {sheet_name}...", end=" ")

        # ── meta e matr da linha TOTAL GERAL ──────────────────────────────
        meta = matr = 0
        for idx in range(len(df)):
            row = df.iloc[idx]
            for c in range(len(row)):
                if str(row.iloc[c]).strip().upper() == "TOTAL GERAL:":
                    meta = float(row.iloc[5]) if pd.notna(row.iloc[5]) else 0
                    matr = float(row.iloc[8]) if pd.notna(row.iloc[8]) else 0
                    break

        # ── colunas FREQ (excluindo 30/01) ────────────────────────────────
        header_idx = None
        for idx in range(len(df)):
            for c in range(len(df.iloc[idx])):
                val = str(df.iloc[idx, c]).strip().upper() if pd.notna(df.iloc[idx, c]) else ""
                if "CURSO" in val and ("MANHÃ" in val or "MANHA" in val):
                    header_idx = idx
                    break
            if header_idx is not None:
                break

        if header_idx is None:
            print("cabeçalho não encontrado, pulando.")
            continue

        header   = df.iloc[header_idx]
        date_row = df.iloc[header_idx + 1] if header_idx + 1 < len(df) else None

        def parse_date_br(val):
            """Converte data para (mês, dia) respeitando formato DD/MM.
            Datas armazenadas pelo Excel como datetime precisam de swap (lê MM/DD mas planilha é DD/MM)."""
            import re as _re
            if val is None or (isinstance(val, float) and pd.isna(val)):
                return None
            if hasattr(val, "month"):
                # Excel interpretou DD/MM como MM/DD — inverte
                return (val.day, val.month)
            s = str(val).strip()
            m = _re.match(r"^(\d{1,2})[/\-](\d{1,2})$", s)
            if m:
                return (int(m.group(2)), int(m.group(1)))  # (mês, dia)
            return None

        freq_cols = []
        for c in range(len(header)):
            if pd.notna(header.iloc[c]) and str(header.iloc[c]).strip().upper() == "FREQ":
                raw = date_row.iloc[c] if date_row is not None and pd.notna(date_row.iloc[c]) else None
                date_val = str(raw).strip() if raw is not None else ""
                if date_val in SKIP_DATES:
                    continue
                parsed = parse_date_br(raw)
                if parsed is not None:
                    freq_cols.append((c, parsed))

        # ── linhas TOTAL: (uma por período) ───────────────────────────────
        total_rows_idx = []
        for idx in range(len(df)):
            row = df.iloc[idx]
            for c in range(len(row)):
                if str(row.iloc[c]).strip().upper() == "TOTAL:":
                    total_rows_idx.append(idx)
                    break

        # Última coluna FREQ (cronologicamente) com dados em qualquer período
        freq = 0
        candidates = []
        for fc, date in freq_cols:
            vals = [
                float(df.iloc[ridx, fc])
                for ridx in total_rows_idx
                if fc < len(df.iloc[ridx])
                and pd.notna(df.iloc[ridx, fc])
                and isinstance(df.iloc[ridx, fc], (int, float))
                and df.iloc[ridx, fc] > 0
            ]
            if vals:
                candidates.append((date, fc, sum(vals)))

        if candidates:
            candidates.sort(key=lambda x: x[0])
            freq = candidates[-1][2]

        saldo = round(freq - meta, 1)
        print(f"meta={int(meta)} matr={int(matr)} freq={freq} saldo={saldo}")

        results.append({
            "unit":  sheet_name,
            "meta":  int(meta),
            "matr":  int(matr),
            "freq":  freq,
            "saldo": saldo,
        })

    return results


def extrair_freq_por_periodo(sheets):
    """Soma a frequência do último dia registrado por período (Manhã/Tarde/Noite)."""
    import re as _re
    SKIP = {'30/01','30/1'}

    def date_to_tuple(val):
        if val is None or (isinstance(val, float) and pd.isna(val)): return None
        if hasattr(val, 'month'): return (val.day, val.month)
        s = str(val).strip()
        if s in SKIP: return None
        m = _re.match(r'^(\d{1,2})[/\-](\d{1,2})$', s)
        if m: return (int(m.group(2)), int(m.group(1)))
        return None

    freq_by_period = {'Manhã': 0, 'Tarde': 0, 'Noite': 0}
    cedesp_sheets = [s for s in sheets if 'CEDESP' in s]

    for sheet_name in cedesp_sheets:
        df = sheets.get(sheet_name)
        if df is None: continue
        period_rows = []
        for idx in range(len(df)):
            for col in range(df.shape[1]):
                val = str(df.iloc[idx,col]).strip().upper() if pd.notna(df.iloc[idx,col]) else ''
                if 'CURSO' in val:
                    if 'MANHÃ' in val or 'MANHA' in val: period_rows.append((idx,'Manhã'))
                    elif 'TARDE' in val: period_rows.append((idx,'Tarde'))
                    elif 'NOITE' in val: period_rows.append((idx,'Noite'))
                    break
        for pidx, (header_idx, turno) in enumerate(period_rows):
            header   = df.iloc[header_idx]
            date_row = df.iloc[header_idx+1] if header_idx+1 < len(df) else None
            if date_row is None: continue
            freq_cols = []
            for col in range(df.shape[1]):
                if pd.notna(header.iloc[col]) and str(header.iloc[col]).strip().upper()=='FREQ':
                    dt = date_to_tuple(date_row.iloc[col])
                    if dt: freq_cols.append((col, dt))
            next_idx = period_rows[pidx+1][0] if pidx+1<len(period_rows) else len(df)
            total_row_idx = None
            for ridx in range(header_idx+2, next_idx):
                row = df.iloc[ridx]
                for col in range(df.shape[1]):
                    if str(row.iloc[col]).strip().upper() == 'TOTAL:':
                        total_row_idx = ridx; break
                if total_row_idx: break
            if not total_row_idx: continue
            candidates = []
            for fc, dt in freq_cols:
                if fc < df.shape[1]:
                    v = df.iloc[total_row_idx, fc]
                    if pd.notna(v) and isinstance(v,(int,float)) and v > 0:
                        candidates.append((dt, v))
            if candidates:
                candidates.sort(key=lambda x: x[0])
                freq_by_period[turno] += candidates[-1][1]
    return freq_by_period


def extrair_horarios(sheets):
    """Extrai totais por horário da aba TOTAIS."""
    df = sheets.get("TOTAIS")
    if df is None:
        return None

    horarios = {}
    periodo_atual = None

    for idx in range(len(df)):
        row = df.iloc[idx]
        val0 = str(row.iloc[0]).strip().upper() if pd.notna(row.iloc[0]) else ""

        if val0 in ("MANHÃ", "TARDE", "NOITE"):
            periodo_atual = val0.capitalize()
            horarios[periodo_atual] = {"meta": 0, "matr": 0, "c1a7": 0, "c8": 0}
        elif periodo_atual and val0 == "TOTAL":
            horarios[periodo_atual]["meta"] = int(row.iloc[1]) if pd.notna(row.iloc[1]) else 0
            horarios[periodo_atual]["matr"] = int(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
        elif periodo_atual and "CEDESP 1" in val0 and "7" in val0:
            horarios[periodo_atual]["c1a7"] = int(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
        elif periodo_atual and "CEDESP 8" in val0:
            horarios[periodo_atual]["c8"] = int(row.iloc[2]) if pd.notna(row.iloc[2]) else 0

    return horarios


def gerar_html(units, horarios, freq_por_periodo, data_atualizacao):
    units_js = json.dumps(units, ensure_ascii=False)

    # Calcula KPIs globais
    total_meta  = sum(u["meta"]  for u in units)
    total_matr  = sum(u["matr"]  for u in units)
    total_freq  = sum(u["freq"]  for u in units)
    superaram   = sum(1 for u in units if u["freq"] >= u["meta"])

    # Monta blocos de horário
    def horario_block(periodo, emoji, badge_class, color_var, data, freq_val):
        if not data:
            return ""
        pct_num = round((data["matr"] / data["meta"] - 1) * 100) if data["meta"] > 0 else 0
        pct_str = f"+{pct_num}%" if pct_num >= 0 else f"{pct_num}%"
        freq_pct = round(freq_val / data["matr"] * 100, 1) if data["matr"] > 0 and freq_val > 0 else 0
        freq_color = "#1a7a3e" if freq_pct >= 85 else "#c97c1a" if freq_pct >= 70 else "#e63827"
        freq_html = f'<div style="margin-top:10px;padding-top:10px;border-top:1px solid var(--border);display:flex;align-items:center;justify-content:space-between"><div><div class="h-val" style="color:{freq_color};font-size:22px">{freq_val:.0f}</div><div class="h-label">Frequentes no dia</div></div><div style="text-align:right"><div style="font-family:JetBrains Mono,monospace;font-size:13px;color:{freq_color};font-weight:600">{freq_pct}%</div><div class="h-meta">das matrículas</div></div></div>' if freq_val > 0 else ""
        return f"""
    <div class="horario-card">
      <div class="horario-badge {badge_class}">{emoji} {periodo}</div>
      <div class="horario-nums">
        <div>
          <div class="h-val" style="color:{color_var}">{data["matr"]:,}</div>
          <div class="h-label">Matrículas</div>
        </div>
        <div style="text-align:right">
          <div class="h-meta">META: {data["meta"]}</div>
          <div class="h-meta">CEDESP 1–7: {data["c1a7"]}</div>
          <div class="h-meta">CEDESP 8: {data["c8"]}</div>
        </div>
      </div>
      <div class="prog-wrap">
        <div class="prog-bar"><div class="prog-fill" style="width:100%;background:{color_var}"></div></div>
        <div class="prog-pct" style="color:{color_var}">{pct_str}</div>
      </div>
      {freq_html}
    </div>"""

    h = horarios or {}
    manha_block = horario_block("Manhã", "☀", "badge-manha", "var(--accent)",  h.get("Manhã", {}), freq_por_periodo.get("Manhã", 0))
    tarde_block  = horario_block("Tarde", "🌤", "badge-tarde", "var(--accent2)", h.get("Tarde", {}), freq_por_periodo.get("Tarde", 0))
    noite_block  = horario_block("Noite", "🌙", "badge-noite", "var(--purple)",  h.get("Noite", {}), freq_por_periodo.get("Noite", 0))

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard CEDESP Dom Bosco Itaquera — 1º Semestre 2026</title>
<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<style>
  :root {{
    --bg: #f4f6fb;
    --surface: #ffffff;
    --surface2: #eef1f8;
    --border: #d0d8ea;
    --text: #1a2340;
    --text-muted: #5c6b8a;
    --accent: #21438e;
    --accent2: #e63827;
    --green: #1a7a3e;
    --red: #e63827;
    --orange: #c97c1a;
    --purple: #21438e;
  }}

  * {{ box-sizing: border-box; margin: 0; padding: 0; }}

  body {{
    background: var(--bg);
    color: var(--text);
    font-family: 'Poppins', sans-serif;
    min-height: 100vh;
    overflow-x: hidden;
  }}

  header {{
    border-bottom: 1px solid var(--border);
    padding: 0 40px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    height: 72px;
    position: sticky;
    top: 0;
    background: #21438e;
    box-shadow: 0 2px 8px rgba(33,67,142,0.18);
    backdrop-filter: blur(12px);
    z-index: 100;
  }}

  .logo {{ display: flex; align-items: center; gap: 12px; }}

  .logo-mark {{
    width: 36px; height: 36px;
    background: #e63827;
    border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    font-family: 'JetBrains Mono', monospace;
    font-weight: 500; color: #ffffff; font-size: 13px; letter-spacing: -0.5px;
  }}

  .logo-text {{
    font-family: 'Poppins', sans-serif;
    font-size: 18px; letter-spacing: -0.3px; color: #ffffff;
  }}
  .logo-text span {{ color: #e63827; }}

  .header-meta {{
    display: flex; align-items: center; gap: 24px;
  }}

  .hm-item {{ text-align: right; }}
  .hm-val {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 20px; font-weight: 600; color: #ffffff;
    letter-spacing: -0.5px;
  }}
  .hm-label {{
    font-size: 11px; color: rgba(255,255,255,0.70); margin-top: 1px;
  }}

  main {{ max-width: 1280px; margin: 0 auto; padding: 32px 40px 60px; }}

  .section-header {{
    display: flex; align-items: center; gap: 12px;
    margin-bottom: 20px; margin-top: 36px;
  }}
  .section-header:first-child {{ margin-top: 0; }}
  .section-title {{
    font-family: 'Poppins', sans-serif;
    font-size: 18px; letter-spacing: -0.3px; white-space: nowrap;
  }}
  .section-line {{ flex: 1; height: 1px; background: var(--border); }}
  .section-badge {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 10px; color: var(--text-muted);
    border: 1px solid var(--border); padding: 3px 10px; border-radius: 20px;
    white-space: nowrap;
  }}

  .kpi-grid {{
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 12px;
    margin-bottom: 8px;
  }}

  .kpi-card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 20px 24px;
    position: relative;
    overflow: hidden;
  }}

  .kpi-card::before {{
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
  }}
  .kpi-card.gold::before   {{ background: #21438e; }}
  .kpi-card.blue::before   {{ background: #e63827; }}
  .kpi-card.green::before  {{ background: #1a7a3e; }}
  .kpi-card.purple::before {{ background: #21438e; }}

  .kpi-label {{
    font-size: 11px; color: var(--text-muted);
    text-transform: uppercase; letter-spacing: 0.6px; margin-bottom: 8px;
  }}
  .kpi-value {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 36px; font-weight: 500;
    letter-spacing: -1px; line-height: 1;
    margin-bottom: 4px;
  }}
  .kpi-sub {{ font-size: 12px; color: var(--text-muted); }}

  .two-col {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 16px; }}

  .chart-card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 24px;
  }}
  .chart-title {{
    font-family: 'Poppins', sans-serif;
    font-size: 15px; margin-bottom: 4px;
  }}
  .chart-sub {{ font-size: 12px; color: var(--text-muted); margin-bottom: 20px; }}

  .table-card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    overflow: hidden;
  }}
  .table-header {{
    padding: 20px 24px 16px;
    border-bottom: 1px solid var(--border);
  }}
  .table-title {{
    font-family: 'Poppins', sans-serif;
    font-size: 16px;
  }}
  .table-sub {{ font-size: 12px; color: var(--text-muted); margin-top: 2px; }}

  table {{ width: 100%; border-collapse: collapse; }}
  thead th {{
    padding: 10px 24px;
    text-align: left;
    font-size: 11px; font-weight: 500;
    text-transform: uppercase; letter-spacing: 0.5px;
    color: #ffffff;
    background: #21438e;
    border-bottom: 1px solid #1a3570;
  }}
  th.right, td.right {{ text-align: right; }}
  tbody tr {{ border-bottom: 1px solid var(--border); transition: background 0.15s; }}
  tbody tr:last-child {{ border-bottom: none; }}
  tbody tr:hover {{ background: var(--surface2); }}
  td {{ padding: 14px 24px; font-size: 14px; }}
  td.mono {{ font-family: 'JetBrains Mono', monospace; font-size: 13px; }}

  .unit-name {{ display: flex; align-items: center; gap: 10px; font-weight: 500; }}
  .unit-dot {{ width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }}

  .prog-wrap {{ display: flex; align-items: center; gap: 10px; }}
  .prog-bar {{
    flex: 1; height: 6px;
    background: var(--surface2);
    border-radius: 3px; overflow: hidden;
  }}
  .prog-fill {{ height: 100%; border-radius: 3px; transition: width 0.6s ease; }}
  .prog-pct {{ font-family: 'JetBrains Mono', monospace; font-size: 12px; min-width: 46px; }}

  .saldo-pill {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 12px; padding: 3px 10px;
    border-radius: 20px; font-weight: 500;
  }}
  .saldo-pill.pos {{ background: rgba(26,122,62,0.12); color: #1a7a3e; }}
  .saldo-pill.neg {{ background: rgba(230,56,39,0.10); color: #e63827; }}

  .horario-grid {{
    display: grid; grid-template-columns: repeat(3, 1fr);
    gap: 12px; margin-bottom: 8px;
  }}
  .horario-card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px; padding: 20px 24px;
  }}
  .horario-badge {{
    display: inline-block;
    font-size: 11px; font-weight: 600; letter-spacing: 0.5px;
    padding: 4px 12px; border-radius: 20px;
    margin-bottom: 16px;
  }}
  .badge-manha {{ background: rgba(33,67,142,0.12);  color: #21438e; }}
  .badge-tarde  {{ background: rgba(230,56,39,0.12);  color: #e63827; }}
  .badge-noite  {{ background: rgba(33,67,142,0.08);  color: #21438e; }}

  .horario-nums {{
    display: flex; justify-content: space-between; align-items: flex-end;
    margin-bottom: 12px;
  }}
  .h-val  {{ font-family: 'JetBrains Mono', monospace; font-size: 28px; font-weight: 500; letter-spacing: -1px; }}
  .h-label {{ font-size: 12px; color: var(--text-muted); margin-top: 2px; }}
  .h-meta {{ font-size: 11px; color: var(--text-muted); text-align: right; }}

  footer {{
    border-top: 1px solid var(--border);
    padding: 20px 40px;
    display: flex; justify-content: space-between; align-items: center;
    margin-top: 40px;
  }}
  .footer-text {{ font-size: 12px; color: var(--text-muted); }}

  @media (max-width: 900px) {{
    header, main, footer {{ padding-left: 20px; padding-right: 20px; }}
    .two-col {{ grid-template-columns: 1fr; }}
    .kpi-grid {{ grid-template-columns: 1fr 1fr; }}
    .horario-grid {{ grid-template-columns: 1fr; }}
  }}
</style>
</head>
<body>

<header>
  <div class="logo">
    <div class="logo-mark">CD</div>
    <div class="logo-text">CEDESP <span>Dom Bosco</span> Itaquera</div>
  </div>
  <div class="header-meta">
    <div class="hm-item">
      <div class="hm-val">{total_meta:,}</div>
      <div class="hm-label">Meta Total</div>
    </div>
    <div class="hm-item">
      <div class="hm-val">{total_matr:,}</div>
      <div class="hm-label">Matrículas</div>
    </div>
    <div class="hm-item">
      <div class="hm-val" style="color:#4ade80">{superaram}/8</div>
      <div class="hm-label">Superaram a Meta</div>
    </div>
    <div style="width:1px;height:36px;background:var(--border);margin:0 4px"></div>
    <div class="hm-item">
      <div style="font-family:'DM Mono',monospace;font-size:12px;color:var(--text-muted);background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:6px 14px;text-align:center;line-height:1.5">
        📅 Última atualização<br>
        <span style="color:#21438e;font-size:13px;font-weight:600">{data_atualizacao}</span>
      </div>
    </div>
  </div>
</header>

<main>

  <div class="section-header">
    <div class="section-title">Indicadores Gerais</div>
    <div class="section-line"></div>
    <div class="section-badge">1º Semestre 2026</div>
  </div>

  <div class="kpi-grid">
    <div class="kpi-card gold">
      <div class="kpi-label">Meta Total Convênio</div>
      <div class="kpi-value" style="color:var(--accent)">{total_meta:,}</div>
      <div class="kpi-sub">alunos planejados · 8 unidades</div>
    </div>
    <div class="kpi-card blue">
      <div class="kpi-label">Total Matrículas</div>
      <div class="kpi-value" style="color:var(--accent2)">{total_matr:,}</div>
      <div class="kpi-sub">{round(total_matr/total_meta*100,1)}% da meta total</div>
    </div>
    <div class="kpi-card green">
      <div class="kpi-label">Unidades Acima da Meta de Freq.</div>
      <div class="kpi-value" style="color:var(--green)">{superaram}</div>
      <div class="kpi-sub">de 8 unidades ativas</div>
    </div>
    <div class="kpi-card purple">
      <div class="kpi-label">Alunos Frequentes no Dia</div>
      <div class="kpi-value" style="color:var(--purple)">{total_freq:.0f}</div>
      <div class="kpi-sub">presentes no último dia registrado</div>
    </div>
  </div>

  <div class="section-header" style="margin-top:32px">
    <div class="section-title">Performance Detalhada por Unidade</div>
    <div class="section-line"></div>
    <div class="section-badge">CEDESPs 1–8</div>
  </div>

  <div class="table-card" style="margin-bottom:20px">
    <div class="table-header">
      <div class="table-sub">Dados consolidados · 1º Semestre 2026</div>
    </div>
    <table>
      <thead>
        <tr>
          <th>Unidade</th>
          <th class="right">Meta</th>
          <th class="right">Matrículas</th>
          <th class="right">Frequência</th>
          <th>Atingimento da Meta</th>
          <th class="right">Saldo</th>
          <th class="right">Status</th>
        </tr>
      </thead>
      <tbody id="mainTable"></tbody>
    </table>
  </div>

  <div class="section-header" style="margin-top:32px">
    <div class="section-title">Análise por Unidade</div>
    <div class="section-line"></div>
    <div class="section-badge">CEDESPs 1–8</div>
  </div>

  <div class="two-col">
    <div class="chart-card">
      <div class="chart-title">Meta vs Matrículas vs Frequência</div>
      <div class="chart-sub">Comparativo por unidade — valores absolutos</div>
      <canvas id="barChart" height="240"></canvas>
    </div>
    <div class="chart-card">
      <div class="chart-title">Taxa de Frequência (% da Meta)</div>
      <div class="chart-sub">Frequência realizada ÷ Meta convênio × 100</div>
      <canvas id="radarChart" height="240"></canvas>
    </div>
  </div>

  <div class="section-header">
    <div class="section-title">Distribuição por Horário</div>
    <div class="section-line"></div>
    <div class="section-badge">Sede + Metrô</div>
  </div>

  <div class="horario-grid">
    {manha_block}
    {tarde_block}
    {noite_block}
  </div>

  <div class="section-header">
    <div class="section-title">Matrículas × Cursos Oferecidos</div>
    <div class="section-line"></div>
    <div class="section-badge">Seleção de Cursos</div>
  </div>

  <div class="two-col">
    <div class="chart-card">
      <div class="chart-title">Matrículas por Unidade — Donut</div>
      <div class="chart-sub">Participação de cada CEDESP no total geral</div>
      <canvas id="donutChart" height="260"></canvas>
    </div>
    <div class="chart-card">
      <div class="chart-title">Saldo de Frequência por Unidade</div>
      <div class="chart-sub">Diferença entre frequência realizada e meta convênio</div>
      <canvas id="saldoChart" height="260"></canvas>
    </div>
  </div>

</main>

<footer>
  <div class="footer-text">CEDESP Dom Bosco Itaquera · 1º Semestre 2026 · Dashboard de Acompanhamento</div>
  <div class="footer-text">Atualizado em {data_atualizacao} · Frequência excluindo 30/01 (dia opcional)</div>
</footer>

<script>
  const units = {units_js};

  const COLORS = [
    '#21438e','#3a6bc4','#5a89d4','#e63827',
    '#bc8cff','#ff7b72','#56d364','#79c0ff'
  ];

  // ── TABLE ──
  const tbody = document.getElementById('mainTable');
  units.forEach((u, i) => {{
    const pct = ((u.freq / u.meta) * 100).toFixed(1);
    const pctNum = parseFloat(pct);
    const fillColor = pctNum >= 100 ? '#1a7a3e' : pctNum >= 85 ? '#c97c1a' : '#e63827';
    const fillW = Math.min(pctNum, 100);
    const saldoClass = u.saldo >= 0 ? 'pos' : 'neg';
    const saldoSign = u.saldo >= 0 ? '+' : '';
    const status = u.freq >= u.meta ? '✓ Atingiu' : '✗ Abaixo';
    const statusColor = u.freq >= u.meta ? '#1a7a3e' : '#e63827';
    tbody.innerHTML += `
      <tr>
        <td><div class="unit-name"><div class="unit-dot" style="background:${{COLORS[i]}}"></div>${{u.unit}}</div></td>
        <td class="right mono">${{u.meta}}</td>
        <td class="right mono">${{u.matr}}</td>
        <td class="right mono">${{u.freq}}</td>
        <td>
          <div class="prog-wrap">
            <div class="prog-bar" style="max-width:160px">
              <div class="prog-fill" style="width:${{fillW}}%;background:${{fillColor}}"></div>
            </div>
            <div class="prog-pct" style="color:${{fillColor}}">${{pct}}%</div>
          </div>
        </td>
        <td class="right"><span class="saldo-pill ${{saldoClass}}">${{saldoSign}}${{u.saldo}}</span></td>
        <td class="right" style="font-size:12px;color:${{statusColor}};font-weight:600">${{status}}</td>
      </tr>`;
  }});

  // ── CHARTS ──
  Chart.defaults.color = '#5c6b8a';
  Chart.defaults.font.family = "'Poppins', sans-serif";
  Chart.defaults.borderColor = '#30363d';

  const gridOpts = {{ color: 'rgba(0,0,0,0.07)', drawBorder: false }};

  new Chart(document.getElementById('barChart'), {{
    type: 'bar',
    data: {{
      labels: units.map(u => u.unit),
      datasets: [
        {{ label: 'Meta', data: units.map(u => u.meta),
           backgroundColor: 'rgba(251,191,36,0.30)', borderColor: '#f59e0b',
           borderWidth: 1.5, borderRadius: 4 }},
        {{ label: 'Matrículas', data: units.map(u => u.matr),
           backgroundColor: 'rgba(99,179,237,0.75)', borderColor: '#3b82f6',
           borderWidth: 1, borderRadius: 4 }},
        {{ label: 'Frequência', data: units.map(u => u.freq),
           backgroundColor: 'rgba(52,211,153,0.75)', borderColor: '#10b981',
           borderWidth: 1, borderRadius: 4 }},
      ]
    }},
    options: {{
      responsive: true,
      plugins: {{
        legend: {{ position: 'top', labels: {{ boxWidth: 12, padding: 16 }} }},
        tooltip: {{ backgroundColor: '#1a2340', borderColor: '#d0d8ea', borderWidth: 1, titleColor:'#ffffff', bodyColor:'#d0d8ea' }}
      }},
      scales: {{
        x: {{ grid: gridOpts, ticks: {{ font: {{ size: 11 }} }} }},
        y: {{ grid: gridOpts, ticks: {{ font: {{ size: 11 }} }} }}
      }}
    }}
  }});

  new Chart(document.getElementById('radarChart'), {{
    type: 'radar',
    data: {{
      labels: units.map(u => u.unit),
      datasets: [
        {{ label: 'Freq/Meta %',
           data: units.map(u => +((u.freq / u.meta) * 100).toFixed(1)),
           backgroundColor: 'rgba(99,102,241,0.20)', borderColor: '#6366f1',
           pointBackgroundColor: '#6366f1', pointBorderColor: '#ffffff',
           pointRadius: 4, borderWidth: 2 }},
        {{ label: 'Referência 100%',
           data: Array(8).fill(100),
           backgroundColor: 'rgba(251,191,36,0.06)', borderColor: 'rgba(251,191,36,0.7)',
           borderDash: [5,3], pointRadius: 0, borderWidth: 1.5 }}
      ]
    }},
    options: {{
      responsive: true,
      plugins: {{
        legend: {{ position: 'top', labels: {{ boxWidth: 12, padding: 16 }} }},
        tooltip: {{ backgroundColor: '#1a2340', borderColor: '#d0d8ea', borderWidth: 1, titleColor:'#ffffff', bodyColor:'#d0d8ea' }}
      }},
      scales: {{
        r: {{
          min: 50, max: 130,
          ticks: {{ backdropColor: 'transparent', font: {{ size: 10 }}, stepSize: 20 }},
          grid: {{ color: 'rgba(0,0,0,0.07)' }},
          pointLabels: {{ font: {{ size: 11 }} }}
        }}
      }}
    }}
  }});

  new Chart(document.getElementById('donutChart'), {{
    type: 'doughnut',
    data: {{
      labels: units.map(u => u.unit),
      datasets: [{{
        data: units.map(u => u.matr),
        backgroundColor: COLORS.map(c => c + 'cc'),
        borderColor: COLORS,
        borderWidth: 2,
        hoverOffset: 6,
      }}]
    }},
    options: {{
      responsive: true,
      cutout: '62%',
      plugins: {{
        legend: {{ position: 'right', labels: {{ boxWidth: 12, padding: 12, font: {{ size: 11 }} }} }},
        tooltip: {{ backgroundColor: '#1a2340', borderColor: '#d0d8ea', borderWidth: 1, titleColor:'#ffffff', bodyColor:'#d0d8ea' }}
      }}
    }}
  }});

  new Chart(document.getElementById('saldoChart'), {{
    type: 'bar',
    data: {{
      labels: units.map(u => u.unit),
      datasets: [{{
        label: 'Saldo (Freq − Meta)',
        data: units.map(u => u.saldo),
        backgroundColor: units.map(u => u.saldo >= 0 ? 'rgba(26,122,62,0.65)' : 'rgba(230,56,39,0.65)'),
        borderColor: units.map(u => u.saldo >= 0 ? '#1a7a3e' : '#e63827'),
        borderWidth: 1.5, borderRadius: 4,
      }}]
    }},
    options: {{
      responsive: true,
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{ backgroundColor: '#1a2340', borderColor: '#d0d8ea', borderWidth: 1, titleColor:'#ffffff', bodyColor:'#d0d8ea' }}
      }},
      scales: {{
        x: {{ grid: gridOpts, ticks: {{ font: {{ size: 11 }} }} }},
        y: {{ grid: gridOpts, ticks: {{ font: {{ size: 11 }} }}, border: {{ dash: [4,4] }} }}
      }}
    }}
  }});
</script>
</body>
</html>"""


def main():
    caminho = sys.argv[1] if len(sys.argv) > 1 else PLANILHA_PADRAO
    data_atualizacao = datetime.now().strftime("%d/%m/%Y às %H:%M")

    print(f"\n{'='*55}")
    print(f"  Dashboard Resumo CEDESP — Gerador Automático")
    print(f"  {data_atualizacao}")
    print(f"{'='*55}\n")

    sheets   = carregar_planilha(caminho)
    units    = extrair_unidades(sheets)
    horarios = extrair_horarios(sheets)

    freq_por_periodo = extrair_freq_por_periodo(sheets)
    html = gerar_html(units, horarios, freq_por_periodo, data_atualizacao)
    with open(SAIDA, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n📄  Gerado: {SAIDA}")
    print(f"\n{'='*55}")
    print(f"  ✅  Concluído!")
    print(f"{'='*55}\n")


if __name__ == "__main__":
    main()

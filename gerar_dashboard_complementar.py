#!/usr/bin/env python3
"""
Dashboard Grade de Aulas Complementares — CEDESP Dom Bosco
Cruza grade com frequência da planilha principal.
"""
import sys, json, re, glob, os
from datetime import datetime
import pandas as pd

SKIP_DATES = {'30/01','30/1'}

def date_to_str(val):
    if val is None or (isinstance(val, float) and pd.isna(val)): return None
    if hasattr(val, 'month'): return f"2026-{val.day:02d}-{val.month:02d}"
    s = str(val).strip()
    if s in SKIP_DATES: return None
    m = re.match(r'^(\d{1,2})[/\-](\d{1,2})$', s)
    if m:
        d,mo = int(m.group(1)),int(m.group(2))
        return f"2026-{mo:02d}-{d:02d}"
    return None

MANUAL_MAP = {
    'BARBEIRO': 'BARBEIRO',
    'PROGRAMADOR DE WEB': 'PROGRAMADOR WEB',
    'AUXILIAR ADMINISTRATIVO - TURMA A': 'ASSISTENTE ADMINISTRATIVO - A',
    'AUXILIAR ADMINISTRATIVO - TURMA B': 'ASSISTENTE ADMINISTRATIVO - B',
    'ASSITENTE DE RECURSOS HUMANOS': 'ASSISTENTE ADMINISTRATIVO - A',
    'COSTUREIRO INDUSTRIAL': 'COSTUREIRO INDUSTRIAL DO VESTUÁRIO',
    'CONFEITEIRO': 'CONFEITEIRO (COM NOÇÕES DE SORVETERIA)',
    'OP. E PROGR. DE SISTEMAS AUTOMATIZADOS DE SOLDAGEM A': 'OPERADOR E PROGRAMADOR DE SISTEMAS AUTOMATIZADOS DE SOLDAGEM - A',
    'OP. E PROGR. DE SISTEMAS AUTOMATIZADOS DE SOLDAGEM B': 'OPERADOR E PROGRAMADOR DE SISTEMAS AUTOMATIZADOS DE SOLDAGEM - B',
    'OP. E PROGR. DE SISTEMAS AUTOMATIZADOS DE SOLDAGEM C': 'OPERADOR E PROGRAMADOR DE SISTEMAS AUTOMATIZADOS DE SOLDAGEM - C',
}

def normalize(s):
    return re.sub(r'\s+',' ', re.sub(r'[^A-ZÁÉÍÓÚÂÊÎÔÛÀÃÕÇ ]','', s.upper().strip()))

def extrair_freq(freq_path):
    xl = pd.ExcelFile(freq_path)
    cedesp_sheets = [s for s in xl.sheet_names if 'CEDESP' in s or s in ['LibrasHidráulica','CURSOS DIVERSOS']]
    all_freq = {}; course_meta = {}
    for sheet in cedesp_sheets:
        df = pd.read_excel(freq_path, sheet_name=sheet, header=None)
        period_rows = []
        for idx in range(len(df)):
            for c in range(df.shape[1]):
                val = str(df.iloc[idx,c]).strip().upper() if pd.notna(df.iloc[idx,c]) else ''
                if 'CURSO' in val and any(p in val for p in ['MANHÃ','TARDE','NOITE','MANHA']):
                    period_rows.append(idx); break
        for pidx, header_idx in enumerate(period_rows):
            header = df.iloc[header_idx]; date_row = df.iloc[header_idx+1]
            freq_date_map = {}
            for c in range(df.shape[1]):
                if pd.notna(header.iloc[c]) and str(header.iloc[c]).strip().upper()=='FREQ':
                    ds = date_to_str(date_row.iloc[c])
                    if ds: freq_date_map[c] = ds
            next_idx = period_rows[pidx+1] if pidx+1<len(period_rows) else len(df)
            name_col=3; matr_col=None
            for c in range(df.shape[1]):
                if pd.notna(header.iloc[c]):
                    h = str(header.iloc[c]).upper()
                    if 'CURSO' in h: name_col=c
                    if 'MATR' in h and matr_col is None: matr_col=c
            for ridx in range(header_idx+2, next_idx):
                row = df.iloc[ridx]
                cv = row.iloc[name_col] if name_col<len(row) else None
                if not (pd.notna(cv) and isinstance(cv,str)): continue
                cn = cv.strip()
                if any(s in cn.upper() for s in ['TOTAL','SALDO','PLANJ','NÃO FEZ','ANTES','ELETROTÉC','GUIA','PARADA','CURSOS AB']) or len(cn)<3: continue
                key = re.sub(r'\s+',' ', cn.upper().strip())
                if key not in all_freq: all_freq[key]={}
                if matr_col and matr_col<len(row):
                    mv=row.iloc[matr_col]
                    if pd.notna(mv) and isinstance(mv,(int,float)): course_meta[key]=int(mv)
                for col,ds in freq_date_map.items():
                    if col<len(row):
                        v=row.iloc[col]
                        if pd.notna(v) and isinstance(v,(int,float)): all_freq[key][ds]=int(v)
    course_freq_avg = {}
    for key, dates in all_freq.items():
        vals = [v for v in dates.values() if v > 0]
        if vals:
            course_freq_avg[key] = round(sum(vals) / len(vals), 1)
    return all_freq, course_meta, course_freq_avg

def find_freq_key(nome, all_freq, freq_keys):
    if not nome: return None
    nu = nome.upper().strip()
    for mk, mv in MANUAL_MAP.items():
        if mk in nu or nu in mk:
            fk = re.sub(r'\s+',' ', mv.upper().strip())
            if fk in all_freq: return fk
    n = normalize(nome)
    if n in freq_keys: return freq_keys[n]
    for fk, orig in freq_keys.items():
        if n[:15]==fk[:15]: return orig
        if len(n)>10 and n[:10] in fk: return orig
    return None

def extrair_grade(grade_files):
    courses={}; schedule=[]
    for turno, path in grade_files.items():
        ref = pd.read_excel(path, sheet_name='Ref', header=None)
        hdr = [str(ref.iloc[0,c]).strip() for c in range(ref.shape[1])]
        subj_cols = {c:hdr[c] for c in range(2,ref.shape[1]) if hdr[c] not in ['','Tot. Comp','nan']}
        code_map={}
        for idx in range(1,len(ref)):
            row=ref.iloc[idx]
            if not (pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]) and isinstance(row.iloc[1],str)): continue
            cod=str(row.iloc[0]).strip(); nome=str(row.iloc[1]).strip()
            if any(s in nome for s in ['Legenda:','Possível','Feriados','Fim de','Sem comp']) or len(nome)<3: continue
            planned={subj_cols[c]: int(row.iloc[c]) if pd.notna(row.iloc[c]) and isinstance(row.iloc[c],(int,float)) else 0 for c in subj_cols}
            code_map[cod]=nome
            courses[f"{turno}|{cod}"]={'turno':turno,'cod':cod,'nome':nome,'planned':planned}
        aux=pd.read_excel(path, sheet_name='Aux', header=None)
        subj_hdr={c:str(aux.iloc[0,c]).strip() for c in range(1,aux.shape[1]) if pd.notna(aux.iloc[0,c])}
        for idx in range(1,len(aux)):
            row=aux.iloc[idx]
            if pd.isna(row.iloc[0]) or not hasattr(row.iloc[0],'strftime'): continue
            date_str=row.iloc[0].strftime('%Y-%m-%d')
            for col,subj in subj_hdr.items():
                if col>=len(row): continue
                cell=row.iloc[col]
                if pd.isna(cell) or not isinstance(cell,str): continue
                val=str(cell).strip()
                if val in ['Sem Complementar','Aprendiz','','-']: continue
                parts=val.split('/')
                cod1=parts[0].strip() if parts[0].strip()!='-' else None
                cod2=parts[1].strip() if len(parts)>1 and parts[1].strip()!='-' else None
                schedule.append({'date':date_str,'turno':turno,'subject':subj,
                    'cod1':cod1,'cod2':cod2,
                    'nome1':code_map.get(cod1,cod1) if cod1 else None,
                    'nome2':code_map.get(cod2,cod2) if cod2 else None})
    return list(courses.values()), schedule

def enriquecer(schedule, all_freq, course_meta, course_freq_avg):
    freq_keys = {normalize(k):k for k in all_freq}
    enriched=[]
    for s in schedule:
        e=dict(s)
        for w in ['1','2']:
            nome=s[f'nome{w}']
            fk=find_freq_key(nome,all_freq,freq_keys) if nome else None
            freq_on_date = all_freq[fk].get(s['date']) if fk else None
            freq_avg     = course_freq_avg.get(fk)     if fk else None
            matr         = course_meta.get(fk)         if fk else None
            pct = round(freq_on_date / freq_avg * 100, 1) if (freq_on_date is not None and freq_avg) else None
            e[f'freq{w}']     = freq_on_date
            e[f'freq_avg{w}'] = freq_avg
            e[f'matr{w}']     = matr
            e[f'pct{w}']      = pct
        enriched.append(e)
    return enriched

def gerar_html(courses, schedule, data_at):
    courses_json  = json.dumps(courses,  ensure_ascii=False)
    schedule_json = json.dumps(schedule, ensure_ascii=False)
    n_cursos   = len(courses)
    n_aulas    = len(schedule)
    n_materias = len(set(s['subject'] for s in schedule))
    com_freq   = sum(1 for s in schedule if s.get('freq1') is not None or s.get('freq2') is not None)
    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Grade Complementar — CEDESP Dom Bosco 2026</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --bg:#f0f3fa;--surface:#fff;--surface2:#e8edf7;--border:#cdd5e8;
  --text:#1a2340;--muted:#5c6b8a;
  --blue:#21438e;--red:#e63827;--green:#1a7a3e;--amber:#c97c1a;
  --indigo:#4f46e5;--teal:#0d9488;--pink:#db2777;--orange:#ea580c;--purple:#7c3aed;
}}
body{{font-family:'Poppins',sans-serif;background:var(--bg);color:var(--text);min-height:100vh}}
/* HEADER */
header{{background:var(--blue);padding:0 40px;display:flex;align-items:center;justify-content:space-between;height:64px;box-shadow:0 2px 12px rgba(33,67,142,.25);position:sticky;top:0;z-index:100}}
.hd-left{{display:flex;align-items:center;gap:12px}}
.hd-dot{{width:10px;height:10px;border-radius:50%;background:var(--red)}}
.hd-title{{font-size:15px;font-weight:600;color:#fff}}
.hd-sub{{font-size:11px;color:rgba(255,255,255,.6);margin-top:1px}}
.hd-badge{{font-family:'JetBrains Mono',monospace;font-size:11px;background:rgba(255,255,255,.12);color:rgba(255,255,255,.8);padding:4px 12px;border-radius:20px;border:1px solid rgba(255,255,255,.2)}}
/* LAYOUT */
main{{max-width:1400px;margin:0 auto;padding:32px 40px 60px}}
.section-row{{display:flex;align-items:center;gap:12px;margin-bottom:18px}}
.section-bar{{width:4px;height:18px;background:var(--blue);border-radius:2px}}
.section-title{{font-size:12px;font-weight:600;color:var(--text);text-transform:uppercase;letter-spacing:.8px}}
.section-line{{flex:1;height:1px;background:var(--border)}}
/* KPI */
.kpi-row{{display:grid;grid-template-columns:repeat(5,1fr);gap:16px;margin-bottom:28px}}
.kpi-card{{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:18px 22px;position:relative;overflow:hidden}}
.kpi-card::before{{content:'';position:absolute;top:0;left:0;width:4px;height:100%}}
.kpi-card.blue::before{{background:var(--blue)}}.kpi-card.red::before{{background:var(--red)}}
.kpi-card.green::before{{background:var(--green)}}.kpi-card.amber::before{{background:var(--amber)}}
.kpi-card.indigo::before{{background:var(--indigo)}}
.kpi-lbl{{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;margin-bottom:5px}}
.kpi-val{{font-family:'JetBrains Mono',monospace;font-size:26px;font-weight:600}}
.kpi-sub{{font-size:10px;color:var(--muted);margin-top:3px}}
/* FILTERS */
.filters{{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:14px 18px;margin-bottom:20px;display:flex;flex-wrap:wrap;gap:10px;align-items:center}}
.flbl{{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;white-space:nowrap}}
.chip-group{{display:flex;gap:5px;flex-wrap:wrap}}
.chip{{padding:3px 11px;border-radius:20px;font-size:11px;cursor:pointer;border:1px solid var(--border);background:transparent;color:var(--muted);transition:all .15s;font-family:'Poppins',sans-serif}}
.chip:hover{{border-color:var(--blue);color:var(--blue)}}
.chip.active{{background:var(--blue);border-color:var(--blue);color:#fff}}
.search-input{{padding:5px 14px;border:1px solid var(--border);border-radius:20px;font-family:'Poppins',sans-serif;font-size:12px;color:var(--text);background:var(--bg);outline:none;width:210px;transition:border .15s;margin-left:auto}}
.search-input:focus{{border-color:var(--blue)}}
/* CARD */
.card{{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:22px;margin-bottom:22px}}
.card-title{{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.7px;color:var(--muted);margin-bottom:3px}}
.card-sub{{font-size:11px;color:var(--muted);margin-bottom:18px}}
.grid-2{{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:22px}}
/* TABLE */
.tbl-wrap{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
thead th{{padding:9px 14px;text-align:left;font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:.7px;text-transform:uppercase;color:#fff;background:var(--blue);white-space:nowrap;cursor:pointer;user-select:none}}
thead th:hover{{background:#1a3570}}
thead th .arr{{margin-left:3px;opacity:.7}}
tbody tr{{border-bottom:1px solid var(--border);transition:background .12s}}
tbody tr:hover{{background:var(--surface2)}}
td{{padding:8px 14px;vertical-align:middle}}
.mono{{font-family:'JetBrains Mono',monospace}}
.muted{{color:var(--muted)}}
/* PILLS */
.subj-pill{{display:inline-flex;align-items:center;padding:2px 8px;border-radius:3px;font-size:10px;font-family:'JetBrains Mono',monospace;font-weight:500;border:1px solid;white-space:nowrap}}
.turno-M{{background:rgba(234,88,12,.1);color:#ea580c}}
.turno-T{{background:rgba(2,132,199,.1);color:#0284c7}}
.turno-N{{background:rgba(124,58,237,.1);color:#7c3aed}}
.cod-badge{{font-family:'JetBrains Mono',monospace;font-size:9px;background:var(--surface2);border:1px solid var(--border);padding:1px 5px;border-radius:2px;color:var(--muted)}}
/* FREQ CELL */
.freq-cell{{display:flex;flex-direction:column;gap:4px}}
.freq-row{{display:flex;align-items:center;gap:6px;font-size:11px}}
.freq-bar-wrap{{width:60px;height:5px;background:var(--surface2);border-radius:3px;overflow:hidden}}
.freq-bar-fill{{height:100%;border-radius:3px}}
.freq-pct{{font-family:'JetBrains Mono',monospace;font-size:10px;min-width:36px}}
/* HOJE */
.hoje-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:12px;margin-bottom:8px}}
.hoje-item{{background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:14px 16px}}
.hoje-subj{{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}}
.hoje-pair{{font-size:12px;color:var(--muted);line-height:1.7}}
.hoje-pair strong{{color:var(--text);font-weight:500}}
.empty-state{{text-align:center;padding:40px;color:var(--muted);font-size:13px}}
.result-count{{font-size:11px;color:var(--muted);margin-bottom:10px}}
/* PAGINATION */
.pagination{{display:flex;align-items:center;gap:6px;padding:12px 16px;border-top:1px solid var(--border)}}
.page-btn{{padding:3px 9px;border:1px solid var(--border);border-radius:4px;background:transparent;font-family:'Poppins',sans-serif;font-size:11px;color:var(--muted);cursor:pointer;transition:all .15s}}
.page-btn:hover{{border-color:var(--blue);color:var(--blue)}}
.page-btn.active{{background:var(--blue);border-color:var(--blue);color:#fff}}
.page-info{{font-size:11px;color:var(--muted)}}
/* ALERTA BAIXA FREQ */
.alerta{{background:rgba(230,56,39,.07);border:1px solid rgba(230,56,39,.25);border-radius:6px;padding:2px 8px;font-size:10px;color:var(--red);white-space:nowrap}}
footer{{text-align:center;padding:32px;font-size:11px;color:var(--muted)}}
</style>
</head>
<body>
<header>
  <div class="hd-left">
    <div class="hd-dot"></div>
    <div><div class="hd-title">Grade de Aulas Complementares</div>
    <div class="hd-sub">CEDESP Dom Bosco · 1º Semestre 2026 · Cruzamento com Frequência</div></div>
  </div>
  <div class="hd-badge">📅 {data_at}</div>
</header>
<main>

<!-- KPIs -->
<div style="margin-bottom:28px">
  <div class="section-row"><div class="section-bar"></div><div class="section-title">Visão Geral</div><div class="section-line"></div></div>
  <div class="kpi-row">
    <div class="kpi-card blue"><div class="kpi-lbl">Cursos</div><div class="kpi-val" style="color:var(--blue)">{n_cursos}</div><div class="kpi-sub">nos 3 turnos</div></div>
    <div class="kpi-card red"><div class="kpi-lbl">Aulas Agendadas</div><div class="kpi-val" style="color:var(--red)">{n_aulas}</div><div class="kpi-sub">no semestre</div></div>
    <div class="kpi-card green"><div class="kpi-lbl">Matérias</div><div class="kpi-val" style="color:var(--green)">{n_materias}</div><div class="kpi-sub">disciplinas complementares</div></div>
    <div class="kpi-card amber"><div class="kpi-lbl">Hoje</div><div class="kpi-val" style="color:var(--amber)" id="kpi-hoje">—</div><div class="kpi-sub">aulas complementares</div></div>
    <div class="kpi-card indigo"><div class="kpi-lbl">Com Frequência</div><div class="kpi-val" style="color:var(--indigo)">{com_freq}</div><div class="kpi-sub">aulas com dados cruzados</div></div>
  </div>
</div>

<!-- HOJE -->
<div style="margin-bottom:28px">
  <div class="section-row">
    <div class="section-bar"></div><div class="section-title">Hoje</div><div class="section-line"></div>
    <span id="hoje-label" style="font-size:12px;color:var(--muted)"></span>
  </div>
  <div id="hoje-container"><div class="empty-state">Nenhuma aula complementar agendada para hoje.</div></div>
</div>

<!-- GRÁFICOS -->
<div class="grid-2">
  <div class="card">
    <div class="card-title">Frequência Média por Matéria</div>
    <div class="card-sub">% média da matrícula nos dias com aula complementar</div>
    <canvas id="freqSubjChart" height="260"></canvas>
  </div>
  <div class="card">
    <div class="card-title">Evolução da Presença nas Complementares</div>
    <div class="card-sub">Média mensal: presença no dia da complementar ÷ média geral do curso (100% = comportamento normal)</div>
    <canvas id="evolucaoChart" height="260"></canvas>
  </div>
</div>

<!-- FILTROS + TABELA -->
<div style="margin-bottom:28px">
  <div class="section-row"><div class="section-bar"></div><div class="section-title">Grade Completa</div><div class="section-line"></div></div>
  <div class="filters">
    <div class="flbl">Turno:</div>
    <div class="chip-group" id="chips-turno">
      <button class="chip active" data-v="all">Todos</button>
      <button class="chip" data-v="Manhã">Manhã</button>
      <button class="chip" data-v="Tarde">Tarde</button>
      <button class="chip" data-v="Noite">Noite</button>
    </div>
    <div class="flbl" style="margin-left:8px">Matéria:</div>
    <div class="chip-group" id="chips-subj"></div>
    <div class="flbl" style="margin-left:8px">Mês:</div>
    <div class="chip-group" id="chips-month">
      <button class="chip active" data-v="all">Todos</button>
    </div>
    <div class="flbl" style="margin-left:8px">Data:</div>
    <input type="date" id="filter-date" min="2026-02-01" max="2026-06-30"
      style="padding:4px 10px;border:1px solid var(--border);border-radius:6px;font-family:'Poppins',sans-serif;font-size:11px;color:var(--text);background:var(--bg);outline:none;cursor:pointer"
      title="Filtrar por data específica">
    <button id="clear-date" style="display:none;padding:3px 8px;border:1px solid var(--border);border-radius:4px;background:transparent;font-size:10px;color:var(--muted);cursor:pointer" onclick="clearDate()">✕</button>
    <div class="flbl" style="margin-left:8px">Freq:</div>
    <div class="chip-group" id="chips-freq">
      <button class="chip active" data-v="all">Todos</button>
      <button class="chip" data-v="low">Abaixo 80%</button>
      <button class="chip" data-v="has">Com dados</button>
    </div>
    <input class="search-input" id="search" placeholder="🔍 Buscar curso..." type="text">
  </div>
  <div class="result-count" id="result-count"></div>
  <div class="card" style="padding:0;overflow:hidden">
    <div class="tbl-wrap">
      <table id="mainTable">
        <thead><tr>
          <th data-sort="date">Data <span class="arr">↕</span></th>
          <th data-sort="weekday">Dia</th>
          <th data-sort="turno">Turno <span class="arr">↕</span></th>
          <th data-sort="subject">Matéria <span class="arr">↕</span></th>
          <th>Curso 1</th>
          <th style="text-align:right">Freq 1</th>
          <th>Curso 2</th>
          <th style="text-align:right">Freq 2</th>
        </tr></thead>
        <tbody id="mainBody"></tbody>
      </table>
      <div class="empty-state" id="emptyState" style="display:none">Nenhuma aula encontrada.</div>
    </div>
    <div class="pagination" id="pagination"></div>
  </div>
</div>

<!-- PROGRESSO POR CURSO -->
<div>
  <div class="section-row"><div class="section-bar"></div><div class="section-title">Progresso por Curso</div><div class="section-line"></div></div>
  <div id="progContainer"></div>
</div>

</main>
<footer>Grade de Aulas Complementares · CEDESP Dom Bosco · 1º Semestre 2026</footer>

<script>
const COURSES  = {courses_json};
const SCHEDULE = {schedule_json};
const WD = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];
const MONTHS = {{'02':'Fev','03':'Mar','04':'Abr','05':'Mai','06':'Jun'}};
const SUBJ_COLORS = {{
  'Artes':'#e63827','Português':'#21438e','Port(He)':'#21438e','Port(Li)':'#3a6bc4',
  'Port(-)':'#3a6bc4','Esporte':'#1a7a3e','Esp(Br)':'#1a7a3e','Esp(Em)':'#16a34a',
  'OT':'#c97c1a','Matemática':'#7c3aed','Mat(Al)':'#7c3aed','Mat(Ti)':'#9333ea',
  'Mat(Fe)':'#a855f7','Cidadania':'#0891b2','ID':'#db2777','ID/IA':'#db2777',
}};
function sc(s){{for(const k of Object.keys(SUBJ_COLORS)){{if(s.includes(k))return SUBJ_COLORS[k];}}return'#5c6b8a';}}
function wd(ds){{return WD[new Date(ds+'T12:00:00').getDay()];}}
function fmt(ds){{const[y,m,d]=ds.split('-');return`${{d}}/${{m}}`;}}
function pctColor(p){{return p>=90?'#1a7a3e':p>=70?'#c97c1a':'#e63827';}}

const ALL_SUBJS  = [...new Set(SCHEDULE.map(s=>s.subject))].sort();
const ALL_MONTHS = [...new Set(SCHEDULE.map(s=>s.date.substring(0,7)))].sort();

// Build filter chips
const sc2 = document.getElementById('chips-subj');
['all',...ALL_SUBJS].forEach(s=>{{
  const b=document.createElement('button');
  b.className='chip'+(s==='all'?' active':'');
  b.dataset.v=s; b.textContent=s==='all'?'Todas':s;
  sc2.appendChild(b);
}});
const mc = document.getElementById('chips-month');
ALL_MONTHS.forEach(m=>{{
  const b=document.createElement('button');
  b.className='chip'; b.dataset.v=m;
  b.textContent=MONTHS[m.split('-')[1]]||m;
  mc.appendChild(b);
}});

// STATE
let fTurno='all',fSubj='all',fMonth='all',fFreq='all',fDate='',search='';
let sortCol='date',sortDir=1,page=1;
const PAGE=50;


function getFiltered(){{
  return SCHEDULE.filter(s=>{{
    if(fTurno!=='all'&&s.turno!==fTurno) return false;
    if(fSubj!=='all'&&s.subject!==fSubj) return false;
    if(fMonth!=='all'&&!s.date.startsWith(fMonth)) return false;
    if(fDate&&s.date!==fDate) return false;
    if(fFreq==='has'&&s.freq1==null&&s.freq2==null) return false;
    if(fFreq==='low'){{
      const p1=s.pct1,p2=s.pct2;
      if((p1==null||p1>=80)&&(p2==null||p2>=80)) return false;
    }}
    if(search){{
      const q=search.toLowerCase();
      if(!((s.nome1||'').toLowerCase().includes(q)||(s.nome2||'').toLowerCase().includes(q)||
           (s.cod1||'').toLowerCase().includes(q)||(s.cod2||'').toLowerCase().includes(q)||
           s.subject.toLowerCase().includes(q))) return false;
    }}
    return true;
  }}).sort((a,b)=>{{
    let av=a[sortCol]??'',bv=b[sortCol]??'';
    return av<bv?-sortDir:av>bv?sortDir:0;
  }});
}}

function freqCell(nome,cod,freq,freqAvg,matr,pct){{
  if(!nome&&!cod) return '<td colspan="2"><span style="display:inline-flex;align-items:center;gap:5px;background:rgba(201,124,26,.12);border:1px solid rgba(201,124,26,.35);color:#c97c1a;font-size:10px;font-weight:600;padding:2px 9px;border-radius:3px;letter-spacing:.3px">— SEM TURMA</span></td>';
  const nc = `<span class="cod-badge">${{cod||''}}</span> ${{(nome||'').substring(0,32)}}`;
  if(freq==null){{
    return `<td>${{nc}}</td><td class="muted mono" style="font-size:10px;text-align:right">sem dados</td>`;
  }}
  const col=pctColor(pct);
  const alerta=pct!=null&&pct<90?'<span class="alerta">⚠ faltou</span>':'';
  const tip=freqAvg!=null?`title="Dia da complementar: ${{freq}} | Média geral: ${{freqAvg}}"`:'';
  return `<td>${{nc}}</td>
  <td style="text-align:right;white-space:nowrap" ${{tip}}>
    <div style="display:flex;align-items:center;gap:6px;justify-content:flex-end">
      ${{alerta}}
      <span class="mono" style="font-size:11px;color:${{col}};font-weight:500">${{freq}}</span>
      <span style="color:var(--muted);font-size:10px">/ ~${{freqAvg??'—'}}</span>
    </div>
  </td>`;
}}

function renderTable(){{
  const filtered=getFiltered();
  const maxP=Math.ceil(filtered.length/PAGE)||1;
  if(page>maxP) page=maxP;
  const slice=filtered.slice((page-1)*PAGE, page*PAGE);
  const tbody=document.getElementById('mainBody');
  const empty=document.getElementById('emptyState');
  document.getElementById('result-count').textContent=`${{filtered.length}} aulas encontradas`;
  if(!slice.length){{tbody.innerHTML='';empty.style.display='block';renderPag(0,1);return;}}
  empty.style.display='none';
  const col=s=>sc(s);
  tbody.innerHTML=slice.map(s=>{{
    const tc=s.turno==='Manhã'?'M':s.turno==='Tarde'?'T':'N';
    const cc=col(s.subject);
    return`<tr>
      <td class="mono muted" style="font-size:11px">${{fmt(s.date)}}</td>
      <td class="muted" style="font-size:11px">${{wd(s.date)}}</td>
      <td><span class="turno-${{tc}}" style="font-size:10px;padding:2px 7px;border-radius:3px;font-weight:500">${{s.turno}}</span></td>
      <td><span class="subj-pill" style="color:${{cc}};border-color:${{cc}}22;background:${{cc}}11">${{s.subject}}</span></td>
      ${{freqCell(s.nome1,s.cod1,s.freq1,s.freq_avg1,s.matr1,s.pct1)}}
      ${{freqCell(s.nome2,s.cod2,s.freq2,s.freq_avg2,s.matr2,s.pct2)}}
    </tr>`;
  }}).join('');
  renderPag(filtered.length,maxP);
}}

function renderPag(total,maxP){{
  const el=document.getElementById('pagination');
  if(maxP<=1){{el.innerHTML='';return;}}
  let h=`<span class="page-info">${{(page-1)*PAGE+1}}–${{Math.min(page*PAGE,total)}} de ${{total}}</span>`;
  if(page>1) h+=`<button class="page-btn" onclick="goPage(${{page-1}})">‹</button>`;
  const s=Math.max(1,page-2),e=Math.min(maxP,page+2);
  for(let i=s;i<=e;i++) h+=`<button class="page-btn${{i===page?' active':''}}" onclick="goPage(${{i}})">${{i}}</button>`;
  if(page<maxP) h+=`<button class="page-btn" onclick="goPage(${{page+1}})">›</button>`;
  el.innerHTML=h;
}}
function goPage(p){{page=p;renderTable();window.scrollTo({{top:0,behavior:'smooth'}});}}

// HOJE
function renderHoje(){{
  const today=new Date().toISOString().split('T')[0];
  const items=SCHEDULE.filter(s=>s.date===today);
  document.getElementById('kpi-hoje').textContent=items.length;
  document.getElementById('hoje-label').textContent=
    new Date(today+'T12:00:00').toLocaleDateString('pt-BR',{{weekday:'long',day:'numeric',month:'long'}});
  const cnt=document.getElementById('hoje-container');
  if(!items.length){{cnt.innerHTML='<div class="empty-state">Nenhuma aula complementar hoje.</div>';return;}}

  const TURNO_CFG={{
    'Manhã':  {{icon:'☀',color:'#ea580c',bg:'rgba(234,88,12,.08)',border:'rgba(234,88,12,.25)'}},
    'Tarde':  {{icon:'🌤',color:'#0284c7',bg:'rgba(2,132,199,.08)',border:'rgba(2,132,199,.25)'}},
    'Noite':  {{icon:'🌙',color:'#7c3aed',bg:'rgba(124,58,237,.08)',border:'rgba(124,58,237,.25)'}},
  }};

  const SEM_TURMA = '<span style="display:inline-flex;align-items:center;gap:5px;background:rgba(201,124,26,.12);border:1px solid rgba(201,124,26,.35);color:#c97c1a;font-size:10px;font-weight:600;padding:2px 9px;border-radius:3px;letter-spacing:.3px">— SEM TURMA</span>';

  const fmtH=(freq,avg,pct)=>freq==null?'':`<span class="mono" style="font-size:10px;color:${{pctColor(pct)}}">${{freq}} / ~${{avg??'—'}}</span>${{pct!=null&&pct<90?' <span class="alerta">⚠ faltou</span>':''}}`;

  let html='';
  ['Manhã','Tarde','Noite'].forEach(turno=>{{
    const tItems=items.filter(s=>s.turno===turno);
    if(!tItems.length) return;
    const cfg=TURNO_CFG[turno];

    // Group by subject within this turno
    const bySubj={{}};
    tItems.forEach(s=>{{
      if(!bySubj[s.subject]) bySubj[s.subject]=[];
      bySubj[s.subject].push(s);
    }});

    const cards=Object.entries(bySubj).map(([subj,entries])=>{{
      const col=sc(subj);
      const pairs=entries.map(i=>{{
        const f1=fmtH(i.freq1,i.freq_avg1,i.pct1);
        const f2=fmtH(i.freq2,i.freq_avg2,i.pct2);
        const l1=i.nome1?`<strong style="font-weight:600">${{i.cod1}}</strong> ${{i.nome1.substring(0,35)}} ${{f1}}`:SEM_TURMA;
        const l2=i.nome2?`<strong style="font-weight:600">${{i.cod2}}</strong> ${{i.nome2.substring(0,35)}} ${{f2}}`:SEM_TURMA;
        return `<div style="padding:7px 0;border-bottom:1px solid var(--border);line-height:1.8;font-size:12px">
          <div>${{l1}}</div>
          <div>${{l2}}</div>
        </div>`;
      }}).join('');
      return `<div style="background:var(--surface);border:1px solid var(--border);border-top:3px solid ${{col}};border-radius:6px;padding:14px 16px;min-width:240px;flex:1">
        <div style="font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;color:${{col}};margin-bottom:2px">${{subj}}</div>
        <div>${{pairs}}</div>
      </div>`;
    }}).join('');

    html+=`<div style="margin-bottom:24px">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;padding:10px 16px;background:${{cfg.bg}};border:1px solid ${{cfg.border}};border-radius:6px">
        <span style="font-size:16px">${{cfg.icon}}</span>
        <span style="font-size:13px;font-weight:600;color:${{cfg.color}}">${{turno}}</span>
        <span style="font-size:11px;color:var(--muted);margin-left:4px">${{tItems.length}} aula${{tItems.length>1?'s':''}}</span>
      </div>
      <div style="display:flex;flex-wrap:wrap;gap:12px">${{cards}}</div>
    </div>`;
  }});

  cnt.innerHTML=html;
}}

// PROGRESS TABLE
function renderProg(){{
  // Build sched index: turno|cod -> subject -> {{count, freqSum, freqN}}
  // count = aulas JÁ REALIZADAS (data <= hoje); freq = apenas dias com dado cruzado
  const today=new Date().toISOString().split('T')[0];
  const sched={{}};
  SCHEDULE.forEach(s=>{{
    if(s.date>today) return;  // ignora aulas futuras
    [{{cod:s.cod1,pct:s.pct1}},{{cod:s.cod2,pct:s.pct2}}].forEach(x=>{{
      if(!x.cod) return;
      const key=`${{s.turno}}|${{x.cod}}`;
      if(!sched[key]) sched[key]={{}};
      if(!sched[key][s.subject]) sched[key][s.subject]={{count:0,freqSum:0,freqN:0}};
      sched[key][s.subject].count++;
      if(x.pct!=null){{sched[key][s.subject].freqSum+=x.pct;sched[key][s.subject].freqN++;}}
    }});
  }});

  // Normalise schedule subject names → REF planned keys
  const SUBJ_NORM={{
    'Português':'Port.','Port(He)':'Port.','Port(Li)':'Port.','Port(-)':'Port.',
    'Esporte':'Esp.','Esp(Br)':'Esp.','Esp(Em)':'Esp.',
    'Matemática':'Mat.','Mat(Al)':'Mat.','Mat(Ti)':'Mat.','Mat(Fe)':'Mat.',
    'Cidadania':'Cid.',
    'ID':'ID/IA','ID/IA':'ID/IA',
    'Artes':'Artes','OT':'OT',
  }};
  const normSubj=s=>SUBJ_NORM[s]||s;

  const TURNO_CFG={{
    'Manhã': {{icon:'☀',color:'#ea580c',border:'rgba(234,88,12,.3)',bg:'rgba(234,88,12,.06)'}},
    'Tarde': {{icon:'🌤',color:'#0284c7',border:'rgba(2,132,199,.3)',bg:'rgba(2,132,199,.06)'}},
    'Noite': {{icon:'🌙',color:'#7c3aed',border:'rgba(124,58,237,.3)',bg:'rgba(124,58,237,.06)'}},
  }};

  let html='';
  ['Manhã','Tarde','Noite'].forEach(turno=>{{
    const courses=COURSES.filter(c=>c.turno===turno);
    if(!courses.length) return;
    const cfg=TURNO_CFG[turno];

    // Only subjects that exist for this turno
    const turnoSubjs=[...new Set(
      SCHEDULE.filter(s=>s.turno===turno).map(s=>s.subject)
    )].sort();

    const headerCells='<th style="min-width:60px">Cod.</th><th style="min-width:190px">Curso</th>'+
      turnoSubjs.map(s=>`<th style="text-align:center;min-width:100px">${{s}}</th>`).join('')+
      '<th style="text-align:center;min-width:70px">Total</th>';

    const rows=courses.map(c=>{{
      const key=`${{turno}}|${{c.cod}}`;
      const cs=sched[key]||{{}};
      let totalDone=0,totalPlan=0;
      const cells=turnoSubjs.map(s=>{{
        const done    = cs[s]?.count  || 0;
        const planned = c.planned[normSubj(s)] || 0;
        totalDone+=done; totalPlan+=planned;
        if(!planned) return`<td style="text-align:center;color:var(--muted);font-size:11px">—</td>`;
        const pct = Math.min(100,Math.round(done/planned*100));
        const col = pct>=100?'var(--green)':pct>=50?'var(--amber)':'var(--red)';
        const avgF = cs[s]?.freqN>0?(cs[s].freqSum/cs[s].freqN).toFixed(0)+'%':null;
        return`<td style="text-align:center">
          <div style="display:inline-flex;flex-direction:column;align-items:center;gap:2px">
            <div style="display:flex;align-items:center;gap:5px">
              <div style="width:40px;height:5px;background:var(--surface2);border-radius:3px;overflow:hidden">
                <div style="width:${{pct}}%;height:100%;background:${{col}};border-radius:3px"></div>
              </div>
              <span class="mono" style="font-size:9px;color:${{col}}">${{done}}/${{planned}}</span>
            </div>
            ${{avgF?`<span style="font-size:9px;color:var(--muted)">${{avgF}}</span>`:''}}
          </div>
        </td>`;
      }}).join('');
      const totPct  = totalPlan?Math.round(totalDone/totalPlan*100):0;
      const totCol  = totPct>=100?'var(--green)':totPct>=50?'var(--amber)':'var(--red)';
      return`<tr>
        <td><span class="cod-badge">${{c.cod}}</span></td>
        <td style="font-size:11px">${{c.nome.substring(0,48)}}</td>
        ${{cells}}
        <td style="text-align:center"><span class="mono" style="font-size:11px;color:${{totCol}};font-weight:700">${{totPct}}%</span></td>
      </tr>`;
    }}).join('');

    html+=`<div style="margin-bottom:24px">
      <div style="display:flex;align-items:center;gap:10px;padding:10px 16px;background:${{cfg.bg}};border:1px solid ${{cfg.border}};border-radius:6px 6px 0 0;border-bottom:none">
        <span style="font-size:15px">${{cfg.icon}}</span>
        <span style="font-size:13px;font-weight:600;color:${{cfg.color}}">${{turno}}</span>
        <span style="font-size:11px;color:var(--muted)">${{courses.length}} cursos · ${{turnoSubjs.length}} matérias</span>
      </div>
      <div style="overflow-x:auto;border:1px solid ${{cfg.border}};border-radius:0 0 6px 6px;background:var(--surface)">
        <table style="width:100%;border-collapse:collapse;font-size:12px">
          <thead><tr style="background:${{cfg.color}};color:#fff">${{headerCells}}</tr></thead>
          <tbody style="">${{rows}}</tbody>
        </table>
      </div>
    </div>`;
  }});

  document.getElementById('progContainer').innerHTML=html;
}}

// CHARTS
function renderCharts(){{
  // Freq por matéria (das aulas com dados)
  const freqBySubj={{}};
  SCHEDULE.forEach(s=>{{
    [{{p:s.pct1}},{{p:s.pct2}}].forEach(x=>{{
      if(x.p==null) return;
      if(!freqBySubj[s.subject]) freqBySubj[s.subject]={{sum:0,n:0}};
      freqBySubj[s.subject].sum+=x.p; freqBySubj[s.subject].n++;
    }});
  }});
  const fLabels=Object.keys(freqBySubj).sort((a,b)=>freqBySubj[b].sum/freqBySubj[b].n - freqBySubj[a].sum/freqBySubj[a].n);
  const fData=fLabels.map(k=>(freqBySubj[k].sum/freqBySubj[k].n).toFixed(1));
  const fColors=fLabels.map(k=>sc(k));
  new Chart(document.getElementById('freqSubjChart'),{{
    type:'bar',
    data:{{labels:fLabels,datasets:[{{data:fData,backgroundColor:fColors.map(c=>c+'aa'),borderColor:fColors,borderWidth:1,borderRadius:4}}]}},
    options:{{indexAxis:'y',responsive:true,plugins:{{legend:{{display:false}},tooltip:{{backgroundColor:'#1a2340',borderColor:'#cdd5e8',borderWidth:1,callbacks:{{label:ctx=>`Freq média: ${{ctx.raw}}%`}}}}}},scales:{{x:{{min:0,max:100,grid:{{color:'rgba(0,0,0,.06)'}},ticks:{{callback:v=>v+'%'}}}},y:{{grid:{{display:false}}}}}}}}
  }});

  // Evolução mensal da presença nas complementares
  const MONTH_LABELS={{'02':'Fev','03':'Mar','04':'Abr','05':'Mai','06':'Jun'}};
  const monthsAll=[...new Set(SCHEDULE.map(s=>s.date.substring(0,7)))].sort();
  const turnosEv=['Manhã','Tarde','Noite'];
  const turnoColors={{'Manhã':'#ea580c','Tarde':'#0284c7','Noite':'#7c3aed'}};

  const evolDatasets=turnosEv.map(turno=>{{
    const data=monthsAll.map(m=>{{
      const entries=SCHEDULE.filter(s=>s.turno===turno&&s.date.startsWith(m));
      const vals=[];
      entries.forEach(s=>{{
        if(s.pct1!=null) vals.push(s.pct1);
        if(s.pct2!=null) vals.push(s.pct2);
      }});
      return vals.length?+(vals.reduce((a,b)=>a+b,0)/vals.length).toFixed(1):null;
    }});
    return{{
      label:turno, data,
      borderColor:turnoColors[turno],
      backgroundColor:turnoColors[turno]+'22',
      borderWidth:2, pointRadius:5, pointHoverRadius:7,
      tension:0.35, fill:false,
      spanGaps:true,
    }};
  }});

  // Add overall average line
  const overallData=monthsAll.map(m=>{{
    const vals=[];
    SCHEDULE.filter(s=>s.date.startsWith(m)).forEach(s=>{{
      if(s.pct1!=null) vals.push(s.pct1);
      if(s.pct2!=null) vals.push(s.pct2);
    }});
    return vals.length?+(vals.reduce((a,b)=>a+b,0)/vals.length).toFixed(1):null;
  }});
  evolDatasets.push({{
    label:'Geral', data:overallData,
    borderColor:'#1a2340', backgroundColor:'transparent',
    borderWidth:2, borderDash:[5,4],
    pointRadius:4, pointHoverRadius:6,
    tension:0.35, fill:false, spanGaps:true,
  }});

  new Chart(document.getElementById('evolucaoChart'),{{
    type:'line',
    data:{{labels:monthsAll.map(m=>MONTH_LABELS[m.split('-')[1]]||m), datasets:evolDatasets}},
    options:{{
      responsive:true,
      plugins:{{
        legend:{{position:'top',labels:{{boxWidth:12,padding:16,font:{{size:11}}}}}},
        tooltip:{{
          backgroundColor:'#1a2340',borderColor:'#cdd5e8',borderWidth:1,
          callbacks:{{label:ctx=>`${{ctx.dataset.label}}: ${{ctx.raw!=null?ctx.raw+'%':'sem dados'}}`}}
        }},
        annotation:{{}}
      }},
      scales:{{
        x:{{grid:{{display:false}}}},
        y:{{
          min:50,max:130,
          grid:{{color:'rgba(0,0,0,.06)'}},
          ticks:{{callback:v=>v+'%'}},
        }}
      }}
    }}
  }});
}}

// EVENTS
function clearDate(){{
  fDate=''; document.getElementById('filter-date').value='';
  document.getElementById('clear-date').style.display='none';
  // re-activate month chip if date cleared
  page=1; renderTable();
}}
document.getElementById('filter-date').addEventListener('change',e=>{{
  fDate=e.target.value;
  document.getElementById('clear-date').style.display=fDate?'inline':'none';
  // when picking a specific date, deactivate month filter to avoid conflict
  if(fDate){{
    fMonth='all';
    document.querySelectorAll('#chips-month .chip').forEach(b=>b.classList.remove('active'));
    document.querySelector('#chips-month .chip[data-v="all"]').classList.add('active');
  }}
  page=1; renderTable();
}});
function bindChips(groupId, setter){{
  document.querySelectorAll(`#${{groupId}} .chip`).forEach(b=>{{
    b.addEventListener('click',()=>{{
      document.querySelectorAll(`#${{groupId}} .chip`).forEach(x=>x.classList.remove('active'));
      b.classList.add('active'); setter(b.dataset.v); page=1; renderTable();
    }});
  }});
}}
bindChips('chips-turno', v=>fTurno=v);
bindChips('chips-subj',  v=>fSubj=v);
bindChips('chips-month', v=>fMonth=v);
bindChips('chips-freq',  v=>fFreq=v);
document.getElementById('search').addEventListener('input',e=>{{search=e.target.value.toLowerCase().trim();page=1;renderTable();}});
document.querySelectorAll('#mainTable thead th[data-sort]').forEach(th=>{{
  th.addEventListener('click',()=>{{
    if(sortCol===th.dataset.sort)sortDir*=-1;else{{sortCol=th.dataset.sort;sortDir=1;}}
    document.querySelectorAll('#mainTable thead th .arr').forEach(a=>a.textContent='↕');
    const arr=th.querySelector('.arr'); if(arr)arr.textContent=sortDir===1?'↑':'↓';
    page=1;renderTable();
  }});
}});
// turno-prog chips removed (sections now always show all 3 turnos separately)

// INIT
renderHoje(); renderTable(); renderProg(); renderCharts();
</script>
</body></html>"""

if __name__ == '__main__':
    print("\n" + "="*55)
    print("  Dashboard Grade Complementar — Gerador v1.0")
    print("="*55)

    base = sys.argv[1] if len(sys.argv) > 1 else '.'
    freq_path = None
    for p in [os.path.join(base, f) for f in os.listdir(base) if '1º_Sem' in f or '1o_Sem' in f or 'frequencia' in f.lower() or 'Frequência' in f]:
        if p.endswith('.xlsx'):
            freq_path = p; break
    if not freq_path:
        # Try uploads dir
        for p in glob.glob(os.path.join(base, '*.xlsx')):
            if '2026' in p and 'Grade' not in p:
                freq_path = p; break
    if not freq_path:
        freq_path = os.path.join(base, '1º_Sem__-__2026.xlsx')

    grade_files = {}
    for turno, pat in [('Manhã','*Manhã*'), ('Tarde','*Tarde*'), ('Noite','*Noite*')]:
        matches = glob.glob(os.path.join(base, pat))
        if matches:
            grade_files[turno] = matches[0]
            print(f"  ✅ Grade {turno}: {os.path.basename(matches[0])}")

    if os.path.exists(freq_path):
        print(f"  ✅ Frequência: {os.path.basename(freq_path)}")
    else:
        print(f"  ⚠  Frequência não encontrada: {freq_path}")

    print("\n📊 Extraindo frequência...")
    all_freq, course_meta, course_freq_avg = extrair_freq(freq_path) if os.path.exists(freq_path) else ({},{},{})

    print("📅 Extraindo grade...")
    courses, schedule = extrair_grade(grade_files)

    print("🔗 Cruzando dados...")
    enriched = enriquecer(schedule, all_freq, course_meta, course_freq_avg)

    com_freq = sum(1 for s in enriched if s.get('freq1') is not None or s.get('freq2') is not None)
    print(f"   {len(schedule)} aulas | {com_freq} com frequência cruzada")

    data_at = datetime.now().strftime('%d/%m/%Y às %H:%M')
    html = gerar_html(courses, enriched, data_at)

    out = 'dashboard_complementar.html'
    with open(out, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"\n📄  Gerado: {out}")
    print("="*55)
    print("  ✅  Concluído!")
    print("="*55 + "\n")

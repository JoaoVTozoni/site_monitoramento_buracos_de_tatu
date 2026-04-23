import pandas as pd
import json
from collections import defaultdict
from pathlib import Path

# ======================================================
# CONFIGURACAO
# ======================================================
BASE_DIR     = Path(__file__).parent
ARQUIVO_XLSX = BASE_DIR / "dados" / "dados_brutos_207.xlsx"
SAIDA_HTML   = BASE_DIR / "index.html"
DATA_INICIO  = pd.Timestamp("2026-03-02")

IDX_CODIGO          = 4
IDX_AREA_TRABALHADA = 9
IDX_AREA_TOTAL      = 10
IDX_DATA            = 15
IDX_BURACOS         = 28
IDX_SERVICO         = 7
# ======================================================


def fmt_num(v):
    return str(int(v)) if v == int(v) else "{:.1f}".format(v)


def sanitize(text):
    if not isinstance(text, str):
        return text
    return text.encode("utf-8", errors="replace").decode("utf-8")


def gerar():
    df_raw = pd.read_excel(ARQUIVO_XLSX, sheet_name="Table", header=0)

    df = pd.DataFrame()
    df["codigo"]          = df_raw.iloc[:, IDX_CODIGO].astype(str).str.strip()
    df["area_trabalhada"] = pd.to_numeric(df_raw.iloc[:, IDX_AREA_TRABALHADA], errors="coerce")
    df["area_total"]      = pd.to_numeric(df_raw.iloc[:, IDX_AREA_TOTAL],      errors="coerce")
    df["Data"]            = pd.to_datetime(df_raw.iloc[:, IDX_DATA],            errors="coerce").dt.normalize()
    df["buracos"]         = pd.to_numeric(df_raw.iloc[:, IDX_BURACOS],          errors="coerce")
    df["servico"]         = df_raw.iloc[:, IDX_SERVICO].astype(str).str.strip()

    for col in ["codigo", "servico"]:
        df[col] = df[col].apply(lambda x: x.encode("utf-8", errors="replace").decode("utf-8") if isinstance(x, str) else x)

    df = df[df["Data"] >= DATA_INICIO].copy()

    df_valido = df[df["buracos"].notna() & (df["buracos"] > 0) & (~df["codigo"].isin(["", "nan"]))].copy()
    df_valido["buracos"] = df_valido["buracos"].astype(int)
    df_valido = df_valido.drop_duplicates(subset=["codigo", "Data", "buracos"])

    data_mais_recente = df_valido["Data"].max().strftime("%d/%m/%Y")

    rows = []
    for (codigo, data), grp in df_valido.groupby(["codigo", "Data"]):
        areas         = grp["area_trabalhada"].dropna().tolist()
        buracos_lista = grp["buracos"].tolist()
        area_str      = " + ".join(fmt_num(a) for a in areas) if areas else "0"
        area_sum      = sum(areas) if areas else 0.0
        buracos_str   = " + ".join(str(b) for b in buracos_lista) if len(buracos_lista) > 1 else ""
        buracos_total = sum(buracos_lista)
        area_total    = float(grp["area_total"].iloc[0]) if pd.notna(grp["area_total"].iloc[0]) else 0.0
        servicos      = grp["servico"].dropna().unique().tolist()
        servico_str   = " / ".join(s for s in servicos if s and s != "nan")
        try:    data_str = data.strftime("%d/%m/%Y")
        except: data_str = "-"
        rows.append({
            "codigo":              sanitize(str(codigo).strip()),
            "area_total":          area_total,
            "area_trabalhada_num": float(area_sum),
            "area_trabalhada_str": area_str,
            "buracos":             buracos_total,
            "buracos_str":         buracos_str,
            "data":                data_str,
            "data_sort":           int(data.strftime("%Y%m%d")) if pd.notna(data) else 0,
            "servico":             sanitize(servico_str),
        })

    campo_area_trab      = defaultdict(float)
    campo_area_total_map = {}
    campo_buracos_total  = defaultdict(int)
    for r in rows:
        campo_area_trab[r["codigo"]]     += r["area_trabalhada_num"]
        campo_area_total_map[r["codigo"]] = r["area_total"]
        campo_buracos_total[r["codigo"]] += r["buracos"]

    campo_critico = set()
    for codigo in campo_area_trab:
        soma_trab  = campo_area_trab[codigo]
        area_ref   = campo_area_total_map[codigo]
        referencia = soma_trab if soma_trab < area_ref else area_ref
        if referencia <= 0:
            referencia = 1
        if campo_buracos_total[codigo] / referencia >= 1:
            campo_critico.add(codigo)

    for r in rows:
        r["critico"] = r["codigo"] in campo_critico

    data_fim_label    = max(
        (r["data"] for r in rows),
        key=lambda d: d.split("/")[2] + d.split("/")[1] + d.split("/")[0],
        default="-"
    )
    data_inicio_label = DATA_INICIO.strftime("%d/%m/%Y")
    rows_json         = json.dumps(rows, ensure_ascii=True)

    html = gerar_html(rows_json, data_inicio_label, data_fim_label, data_mais_recente)
    Path(SAIDA_HTML).write_bytes(html.encode("utf-8", errors="replace"))

    print("Periodo: " + data_inicio_label + " a " + data_fim_label)
    print("Planilha atualizada em: " + data_mais_recente)
    print("Linhas: " + str(len(rows)) + " | Criticos: " + str(len(campo_critico)))
    print("Arquivo gerado: " + str(SAIDA_HTML))
    print("Concluido com sucesso!")


def gerar_html(rows_json, data_inicio_label, data_fim_label, data_mais_recente):
    return (
        "<!DOCTYPE html>\n"
        "<html lang=\"pt-BR\" data-theme=\"dark\">\n"
        "<head>\n"
        "  <meta charset=\"UTF-8\"/>\n"
        "  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\"/>\n"
        "  <title>Monitoramento de Buracos de Tatu</title>\n"
        "  <style>\n"
        "    :root {--tr:.25s ease}\n"
        "    [data-theme=\"dark\"]{--bg:#0d1117;--bg2:#161b22;--bg3:#1c2330;--border:#21262d;--border-hi:#30363d;--txt0:#e6edf3;--txt1:#8b949e;--txt2:#484f58;--accent:#fbbf24;--green:#4ade80;--green-bg:rgba(34,197,94,.10);--green-bd:rgba(34,197,94,.25);--red:#f87171;--red-bg:rgba(239,68,68,.12);--red-bd:rgba(239,68,68,.30);--blue:#60a5fa;--blue-bg:rgba(37,99,235,.18);--blue-bd:#2563eb;--hdr-bg:linear-gradient(135deg,#0f2a4a 0%,#1a3d6b 100%);--hdr-bd:#2563eb;--row-even:rgba(255,255,255,.013);--row-hover:rgba(255,255,255,.050);--shadow:0 4px 24px rgba(0,0,0,.45);--card-bg:#1c2330;--card-bd:#30363d}\n"
        "    [data-theme=\"light\"]{--bg:#f4f6f9;--bg2:#ffffff;--bg3:#edf0f4;--border:#d1d9e0;--border-hi:#b0bac6;--txt0:#1a202c;--txt1:#4a5568;--txt2:#a0aec0;--accent:#d97706;--green:#16a34a;--green-bg:rgba(22,163,74,.10);--green-bd:rgba(22,163,74,.35);--red:#dc2626;--red-bg:rgba(220,38,38,.09);--red-bd:rgba(220,38,38,.30);--blue:#2563eb;--blue-bg:rgba(37,99,235,.10);--blue-bd:#2563eb;--hdr-bg:linear-gradient(135deg,#1e3a5f 0%,#2d5a8c 100%);--hdr-bd:#3b82f6;--row-even:rgba(0,0,0,.018);--row-hover:rgba(37,99,235,.06);--shadow:0 4px 16px rgba(0,0,0,.10);--card-bg:#ffffff;--card-bd:#d1d9e0}\n"
        "    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}\n"
        "    body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:var(--bg);color:var(--txt0);min-height:100vh;transition:background var(--tr),color var(--tr)}\n"
        "    header{background:var(--hdr-bg);border-bottom:2px solid var(--hdr-bd);padding:20px 32px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px}\n"
        "    .header-left h1{font-size:20px;font-weight:700;color:#fbbf24;letter-spacing:1.2px;text-transform:uppercase}\n"
        "    .header-left p{font-size:12px;color:#94a3b8;margin-top:4px}\n"
        "    .header-right{display:flex;align-items:center;gap:14px}\n"
        "    .header-meta{font-size:12px;color:#64748b;text-align:right;line-height:1.7}\n"
        "    #update-time{color:#4ade80;font-weight:600}\n"
        "    #planilha-time{color:#60a5fa;font-weight:600}\n"
        "    .theme-btn{background:rgba(255,255,255,.10);border:1.5px solid rgba(255,255,255,.25);border-radius:8px;color:#e2e8f0;font-size:18px;width:40px;height:40px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:background .2s,border-color .2s,transform .15s}\n"
        "    .theme-btn:hover{background:rgba(255,255,255,.20);border-color:#fbbf24;transform:scale(1.08)}\n"
        "    .controls{padding:18px 32px 10px;display:flex;gap:10px;flex-wrap:wrap;align-items:center}\n"
        "    .search-combo{position:relative;flex:1;min-width:260px}\n"
        "    .search-combo input{width:100%;background:var(--bg2);border:1.5px solid var(--border-hi);border-radius:8px;color:var(--txt0);font-size:14px;padding:9px 14px 9px 38px;outline:none;transition:border-color var(--tr),background var(--tr),color var(--tr)}\n"
        "    .search-combo input:focus{border-color:var(--blue)}\n"
        "    .search-combo::before{content:\"\\1F50D\";position:absolute;left:11px;top:50%;transform:translateY(-50%);font-size:14px;pointer-events:none;z-index:1}\n"
        "    .dropdown-list{position:absolute;top:calc(100% + 4px);left:0;right:0;background:var(--bg2);border:1.5px solid var(--blue-bd);border-radius:8px;max-height:220px;overflow-y:auto;z-index:1000;display:none;box-shadow:var(--shadow)}\n"
        "    .dropdown-list.open{display:block}\n"
        "    .dropdown-item{padding:9px 14px;cursor:pointer;font-size:13.5px;color:var(--txt0);transition:background .15s;display:flex;align-items:center;gap:8px}\n"
        "    .dropdown-item:hover,.dropdown-item.highlighted{background:var(--blue-bg);color:var(--blue)}\n"
        "    .dd-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}\n"
        "    .dd-dot.crit{background:var(--red)} .dd-dot.ok{background:var(--green)}\n"
        "    .dropdown-list::-webkit-scrollbar{width:6px}.dropdown-list::-webkit-scrollbar-track{background:var(--bg3)}.dropdown-list::-webkit-scrollbar-thumb{background:var(--border-hi);border-radius:3px}\n"
        "    .filter-btn{background:var(--bg2);border:1.5px solid var(--border-hi);border-radius:8px;color:var(--txt1);font-size:13px;padding:9px 15px;cursor:pointer;transition:border-color var(--tr),background var(--tr),color var(--tr);white-space:nowrap}\n"
        "    .filter-btn:hover{border-color:var(--blue-bd);color:var(--txt0)}\n"
        "    .filter-btn.active-all{border-color:var(--blue-bd);background:var(--blue-bg);color:var(--blue);font-weight:700}\n"
        "    .filter-btn.active-crit{border-color:var(--red-bd);background:var(--red-bg);color:var(--red);font-weight:700}\n"
        "    .filter-btn.active-ok{border-color:var(--green-bd);background:var(--green-bg);color:var(--green);font-weight:700}\n"
        "    .cards-area{padding:8px 32px 4px;display:none;gap:14px;flex-wrap:wrap}\n"
        "    .cards-area.visible{display:flex}\n"
        "    .card{background:var(--card-bg);border:1.5px solid var(--card-bd);border-radius:12px;padding:16px 22px;min-width:160px;flex:1;max-width:240px;box-shadow:var(--shadow);transition:background var(--tr),border-color var(--tr)}\n"
        "    .card-label{font-size:11px;color:var(--txt1);text-transform:uppercase;letter-spacing:.8px;margin-bottom:6px}\n"
        "    .card-value{font-size:26px;font-weight:700;color:var(--txt0);transition:color var(--tr)}\n"
        "    .card-value.red{color:var(--red)} .card-value.green{color:var(--green)}\n"
        "    .card-sub{font-size:11px;color:var(--txt2);margin-top:3px}\n"
        "    .table-wrap{padding:6px 32px 32px;overflow-x:auto}\n"
        "    table{width:100%;border-collapse:collapse;font-size:13.5px;background:var(--bg2);border-radius:12px;overflow:hidden;box-shadow:var(--shadow);transition:background var(--tr)}\n"
        "    thead tr{background:var(--bg3);border-bottom:2px solid var(--border-hi)}\n"
        "    th{padding:11px 16px;text-align:left;color:var(--txt1);font-size:11px;text-transform:uppercase;letter-spacing:1px;white-space:nowrap;cursor:pointer;user-select:none;transition:color var(--tr)}\n"
        "    th:hover{color:var(--txt0)} th.sorted{color:var(--accent)}\n"
        "    th .sa{margin-left:3px;opacity:.35;font-size:10px} th.sorted .sa{opacity:1;color:var(--accent)}\n"
        "    th.no-sort{cursor:default}\n"
        "    tbody tr{border-bottom:1px solid var(--border);cursor:pointer;transition:background var(--tr)}\n"
        "    tbody tr:nth-child(even){background:var(--row-even)}\n"
        "    tbody tr:hover{background:var(--row-hover)!important}\n"
        "    tbody tr.row-selected{background:rgba(96,165,250,.08)!important}\n"
        "    [data-theme=\"light\"] tbody tr.row-selected{background:rgba(37,99,235,.07)!important}\n"
        "    td{padding:11px 16px;vertical-align:middle}\n"
        "    .td-code{font-family:'Courier New',monospace;font-weight:700;color:var(--blue);font-size:14px}\n"
        "    .td-area{color:var(--txt1);white-space:nowrap}\n"
        "    .td-num .parts{color:var(--txt2);font-size:12px;font-weight:400;display:block;line-height:1.4}\n"
        "    .td-num .total{font-weight:700;font-size:15px}\n"
        "    .td-date{color:var(--txt2);white-space:nowrap;font-size:13px}\n"
        "    .td-servico{color:var(--txt1);font-size:12.5px;max-width:220px}\n"
        "    .badge{display:inline-flex;align-items:center;gap:4px;padding:4px 10px;border-radius:20px;font-size:11px;font-weight:700;letter-spacing:.3px;text-transform:uppercase;white-space:nowrap}\n"
        "    .badge.critico{background:var(--red-bg);color:var(--red);border:1px solid var(--red-bd)}\n"
        "    .badge.normal{background:var(--green-bg);color:var(--green);border:1px solid var(--green-bd)}\n"
        "    .initial-state{text-align:center;padding:80px 20px;color:var(--txt2)}\n"
        "    .initial-icon{font-size:48px;margin-bottom:14px;opacity:.6}\n"
        "    .initial-state p{font-size:15px;line-height:1.7} .initial-state strong{color:var(--blue)}\n"
        "    .empty-state{text-align:center;padding:60px 20px;color:var(--txt2);font-size:15px}\n"
        "    .empty-icon{font-size:40px;margin-bottom:10px}\n"
        "    footer{text-align:center;padding:16px;font-size:11px;color:var(--txt2);border-top:1px solid var(--border);margin-top:8px}\n"
        "    @media(max-width:640px){header,.controls,.table-wrap,.cards-area{padding-left:16px;padding-right:16px}}\n"
        "  </style>\n"
        "</head>\n"
        "<body>\n"
        "<header>\n"
        "  <div class=\"header-left\">\n"
        "    <h1>&#127807; Monitoramento de Buracos de Tatu</h1>\n"
        "    <p>Periodo: " + data_inicio_label + " ate " + data_fim_label + "</p>\n"
        "  </div>\n"
        "  <div class=\"header-right\">\n"
        "    <div class=\"header-meta\">\n"
        "      <div>Gerado em: <span id=\"update-time\">-</span></div>\n"
        "      <div>Planilha atualizada em: <span id=\"planilha-time\">" + data_mais_recente + "</span></div>\n"
        "    </div>\n"
        "    <button class=\"theme-btn\" onclick=\"toggleTheme()\" title=\"Alternar tema\"><span id=\"theme-icon\">&#9728;&#65039;</span></button>\n"
        "  </div>\n"
        "</header>\n"
        "<div class=\"controls\">\n"
        "  <div class=\"search-combo\">\n"
        "    <input type=\"text\" id=\"search-input\" placeholder=\"Buscar ou selecionar campo...\" autocomplete=\"off\"/>\n"
        "    <div class=\"dropdown-list\" id=\"dropdown-list\"></div>\n"
        "  </div>\n"
        "  <button class=\"filter-btn active-all\" id=\"btn-all\"  onclick=\"setFilter('all')\">Todos</button>\n"
        "  <button class=\"filter-btn\"           id=\"btn-crit\" onclick=\"setFilter('critico')\">&#9888;&#65039; Criticos</button>\n"
        "  <button class=\"filter-btn\"           id=\"btn-ok\"   onclick=\"setFilter('normal')\">&#9989; Normais</button>\n"
        "</div>\n"
        "<div class=\"cards-area\" id=\"cards-area\">\n"
        "  <div class=\"card\"><div class=\"card-label\">Quantidade total de buracos</div><div class=\"card-value\" id=\"card-total-buracos\">-</div><div class=\"card-sub\" id=\"card-campo-label\">no campo selecionado</div></div>\n"
        "  <div class=\"card\"><div class=\"card-label\">Buracos por ha</div><div class=\"card-value\" id=\"card-bph\">-</div><div class=\"card-sub\" id=\"card-ref-label\">referencia: area trabalhada</div></div>\n"
        "  <div class=\"card\"><div class=\"card-label\">Area total do campo</div><div class=\"card-value\" id=\"card-area-total\">-</div><div class=\"card-sub\">ha contratados</div></div>\n"
        "  <div class=\"card\"><div class=\"card-label\">Area trabalhada acumulada</div><div class=\"card-value\" id=\"card-area-trab\">-</div><div class=\"card-sub\">ha percorridos</div></div>\n"
        "</div>\n"
        "<div class=\"table-wrap\">\n"
        "  <div class=\"initial-state\" id=\"initial-state\">\n"
        "    <div class=\"initial-icon\">&#128269;</div>\n"
        "    <p>Selecione um <strong>codigo do campo</strong> na barra de busca<br>ou clique em <strong>Todos</strong> para listar todos os campos.</p>\n"
        "  </div>\n"
        "  <table id=\"main-table\" style=\"display:none\">\n"
        "    <thead><tr>\n"
        "      <th onclick=\"sortBy('critico')\"             data-col=\"critico\">             Status               <span class=\"sa\">&#8645;</span></th>\n"
        "      <th onclick=\"sortBy('codigo')\"              data-col=\"codigo\">              Codigo do Campo      <span class=\"sa\">&#8645;</span></th>\n"
        "      <th onclick=\"sortBy('data_sort')\"           data-col=\"data_sort\">           Data do Registro     <span class=\"sa\">&#8645;</span></th>\n"
        "      <th onclick=\"sortBy('buracos')\"             data-col=\"buracos\">             Buracos de Tatu      <span class=\"sa\">&#8645;</span></th>\n"
        "      <th onclick=\"sortBy('area_trabalhada_num')\" data-col=\"area_trabalhada_num\"> Area Trabalhada (ha) <span class=\"sa\">&#8645;</span></th>\n"
        "      <th onclick=\"sortBy('area_total')\"          data-col=\"area_total\">          Area Total (ha)      <span class=\"sa\">&#8645;</span></th>\n"
        "      <th class=\"no-sort\"                         data-col=\"servico\">             Servico Campo        </th>\n"
        "    </tr></thead>\n"
        "    <tbody id=\"table-body\"></tbody>\n"
        "  </table>\n"
        "  <div class=\"empty-state\" id=\"empty-state\" style=\"display:none\"><div class=\"empty-icon\">&#128269;</div><div>Nenhum campo encontrado.</div></div>\n"
        "</div>\n"
        "<footer>Monitoramento de Buracos de Tatu - Syngenta - Desenvolvido por Joao Tozoni</footer>\n"
        "<script>\n"
        "const ALL_ROWS     = " + rows_json + ";\n"
        "const ALL_CODIGOS  = [...new Set(ALL_ROWS.map(r => r.codigo))].sort();\n"
        "const CRITICOS_SET = new Set(ALL_ROWS.filter(r => r.critico).map(r => r.codigo));\n"
        "let filterMode = 'all', searchTerm = '', selectedCodigo = null, showAll = false;\n"
        "let sortCol = 'data_sort', sortAsc = false, ddHighlight = -1;\n"
        "function applyTheme(t){document.documentElement.setAttribute('data-theme',t);document.getElementById('theme-icon').textContent=t==='dark'?'\\u2600\\uFE0F':'\\uD83C\\uDF19';try{localStorage.setItem('tatu-theme',t)}catch(e){}}\n"
        "function toggleTheme(){applyTheme(document.documentElement.getAttribute('data-theme')==='dark'?'light':'dark')}\n"
        "function getFilteredCodigos(term){return term?ALL_CODIGOS.filter(c=>c.toLowerCase().includes(term.toLowerCase())):ALL_CODIGOS}\n"
        "function renderDropdown(term){const list=document.getElementById('dropdown-list'),items=getFilteredCodigos(term);ddHighlight=-1;if(!items.length){list.classList.remove('open');return}list.innerHTML=items.map(c=>`<div class=\"dropdown-item\" data-codigo=\"${c}\" onclick=\"selectCodigo('${c}')\"><span class=\"dd-dot ${CRITICOS_SET.has(c)?'crit':'ok'}\"></span>${c}</div>`).join('');list.classList.add('open')}\n"
        "function selectCodigo(codigo){selectedCodigo=codigo;showAll=false;document.getElementById('search-input').value=codigo;document.getElementById('dropdown-list').classList.remove('open');searchTerm=codigo.toLowerCase();setFilterSilent('all');render()}\n"
        "function closeDropdown(){setTimeout(()=>document.getElementById('dropdown-list').classList.remove('open'),150)}\n"
        "function ddKeyNav(e){const list=document.getElementById('dropdown-list'),items=list.querySelectorAll('.dropdown-item');if(!list.classList.contains('open')||!items.length)return;if(e.key==='ArrowDown'){e.preventDefault();ddHighlight=Math.min(ddHighlight+1,items.length-1);items.forEach((el,i)=>el.classList.toggle('highlighted',i===ddHighlight));items[ddHighlight].scrollIntoView({block:'nearest'})}else if(e.key==='ArrowUp'){e.preventDefault();ddHighlight=Math.max(ddHighlight-1,0);items.forEach((el,i)=>el.classList.toggle('highlighted',i===ddHighlight));items[ddHighlight].scrollIntoView({block:'nearest'})}else if(e.key==='Enter'){if(ddHighlight>=0)selectCodigo(items[ddHighlight].dataset.codigo);list.classList.remove('open')}else if(e.key==='Escape')list.classList.remove('open')}\n"
        "function setFilterSilent(mode){filterMode=mode;['all','crit','ok'].forEach(k=>document.getElementById('btn-'+k).className='filter-btn');const map={all:['btn-all','active-all'],critico:['btn-crit','active-crit'],normal:['btn-ok','active-ok']};document.getElementById(map[mode][0]).classList.add(map[mode][1])}\n"
        "function setFilter(mode){if(mode==='all'){selectedCodigo=null;searchTerm='';showAll=true;document.getElementById('search-input').value=''}else{showAll=true}setFilterSilent(mode);render()}\n"
        "function sortBy(col){sortAsc=(sortCol===col)?!sortAsc:(col==='codigo'||col==='data_sort');sortCol=col;document.querySelectorAll('th[data-col]').forEach(th=>th.classList.remove('sorted'));const el=document.querySelector('th[data-col=\"'+col+'\"]');if(el)el.classList.add('sorted');render()}\n"
        "function updateCards(data){const area=document.getElementById('cards-area');if(!selectedCodigo||!data.length){area.classList.remove('visible');return}area.classList.add('visible');const totalBuracos=data.reduce((s,r)=>s+r.buracos,0);const somaAreaTrab=data.reduce((s,r)=>s+r.area_trabalhada_num,0);const areaTotal=data[0].area_total;const referencia=somaAreaTrab<areaTotal?somaAreaTrab:areaTotal;const bph=referencia>0?(totalBuracos/referencia):0;const cc=data.some(r=>r.critico)?'red':'green';document.getElementById('card-total-buracos').textContent=totalBuracos.toLocaleString('pt-BR');document.getElementById('card-total-buracos').className='card-value '+cc;document.getElementById('card-campo-label').textContent='campo: '+selectedCodigo;document.getElementById('card-bph').textContent=bph.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});document.getElementById('card-bph').className='card-value '+cc;document.getElementById('card-ref-label').textContent=somaAreaTrab<areaTotal?'ref: area trabalhada ('+somaAreaTrab.toLocaleString('pt-BR',{maximumFractionDigits:1})+' ha)':'ref: area total ('+areaTotal.toLocaleString('pt-BR',{maximumFractionDigits:1})+' ha)';document.getElementById('card-area-total').textContent=areaTotal.toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1});document.getElementById('card-area-trab').textContent=somaAreaTrab.toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})}\n"
        "function onRowClick(codigo){selectCodigo(codigo)}\n"
        "function render(){const elInit=document.getElementById('initial-state'),elTable=document.getElementById('main-table'),elEmpty=document.getElementById('empty-state');const hasContent=selectedCodigo||searchTerm||showAll||filterMode!=='all';if(!hasContent){elInit.style.display='block';elTable.style.display='none';elEmpty.style.display='none';document.getElementById('cards-area').classList.remove('visible');return}let data=ALL_ROWS.slice();if(searchTerm)data=data.filter(r=>r.codigo.toLowerCase().includes(searchTerm));if(filterMode==='critico')data=data.filter(r=>r.critico);if(filterMode==='normal')data=data.filter(r=>!r.critico);data.sort((a,b)=>{let va=a[sortCol],vb=b[sortCol];if(typeof va==='string'){va=va.toLowerCase();vb=vb.toLowerCase()}return va<vb?(sortAsc?-1:1):va>vb?(sortAsc?1:-1):0});elInit.style.display='none';updateCards(selectedCodigo?ALL_ROWS.filter(r=>r.codigo===selectedCodigo):[]);if(!data.length){elTable.style.display='none';elEmpty.style.display='block';return}elEmpty.style.display='none';elTable.style.display='table';const fmt1={minimumFractionDigits:1,maximumFractionDigits:1};document.getElementById('table-body').innerHTML=data.map(r=>{const colNum=r.critico?'var(--red)':'var(--green)';const badge=r.critico?'<span class=\"badge critico\">&#9888;&#65039; Critico</span>':'<span class=\"badge normal\">&#9989; Normal</span>';const buracosCell=r.buracos_str?`<td class=\"td-num\"><span class=\"parts\">${r.buracos_str}</span><span class=\"total\" style=\"color:${colNum}\">= ${r.buracos.toLocaleString('pt-BR')}</span></td>`:`<td class=\"td-num\"><span class=\"total\" style=\"color:${colNum}\">${r.buracos.toLocaleString('pt-BR')}</span></td>`;const areaTxt=r.area_trabalhada_str.includes('+')?`<span style=\"color:var(--txt2);font-size:12px\">${r.area_trabalhada_str}</span> <span style=\"font-weight:700\">= ${r.area_trabalhada_num.toLocaleString('pt-BR',fmt1)} ha</span>`:r.area_trabalhada_str+' ha';return`<tr class=\"${selectedCodigo===r.codigo?' row-selected':''}\" onclick=\"onRowClick('${r.codigo}')\">`+`<td>${badge}</td><td class=\"td-code\">${r.codigo}</td><td class=\"td-date\">&#128197; ${r.data}</td>`+buracosCell+`<td class=\"td-area\">${areaTxt}</td><td class=\"td-area\">${r.area_total.toLocaleString('pt-BR',fmt1)} ha</td><td class=\"td-servico\">${r.servico||''}</td></tr>`}).join('');}\n"
        "document.addEventListener('DOMContentLoaded',()=>{let s='dark';try{s=localStorage.getItem('tatu-theme')||'dark'}catch(e){}applyTheme(s);document.getElementById('update-time').textContent=new Date().toLocaleString('pt-BR');const input=document.getElementById('search-input');input.addEventListener('input',e=>{searchTerm=e.target.value.trim().toLowerCase();selectedCodigo=null;showAll=false;renderDropdown(e.target.value.trim());render()});input.addEventListener('focus',e=>{renderDropdown(e.target.value.trim())});input.addEventListener('blur',closeDropdown);input.addEventListener('keydown',ddKeyNav)});\n"
        "</script>\n"
        "</body>\n"
        "</html>"
    )


if __name__ == "__main__":
    gerar()
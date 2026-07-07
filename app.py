import streamlit as st
from datetime import datetime, timedelta, timezone

st.set_page_config(page_title="Clipping Necton", layout="wide")

import main as _m

# ── CSS ───────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
.section-header {
    font-size: 1.6rem;
    font-weight: 700;
    margin-bottom: 0.2rem;
    padding-bottom: 0.4rem;
    border-bottom: 2px solid #e0e0e0;
}
.bubble {
    background: #f7f8fa;
    border: 1px solid #e4e6ea;
    border-radius: 10px;
    padding: 10px 14px 8px 14px;
    margin-bottom: 6px;
    transition: background 0.15s, border-color 0.15s;
}
.bubble.selected {
    background: #e6f4ea;
    border-color: #34a853;
}
.bubble.extra {
    background: #fffde7;
    border-color: #f9a825;
}
.bubble.extra.selected {
    background: #fff9c4;
    border-color: #f57f17;
}
.bubble-title {
    font-size: 1.05rem;
    font-weight: 500;
    line-height: 1.4;
    color: #1a1a1a;
}
.bubble-meta {
    font-size: 0.78rem;
    color: #888;
    margin-top: 4px;
}
.bubble-meta a {
    color: #1a73e8;
    text-decoration: none;
    font-weight: 500;
}
.fonte-header {
    font-size: 1.5rem !important;
    font-weight: 800 !important;
    color: #ffffff !important;
    background: #222;
    padding: 8px 16px;
    border-radius: 8px;
    margin: 1.4rem 0 0.6rem 0;
    display: inline-block;
}
</style>
""", unsafe_allow_html=True)

# ── helpers ───────────────────────────────────────────────────────────────────

def _run_pipeline(modo, horas=None, inicio=None, fim=None):
    fmt = "%d/%m/%Y %H:%M"
    _m._retool_run_pipeline(
        modo=modo,
        horas=horas,
        janela_inicio=inicio.strftime(fmt) if inicio else None,
        janela_fim=fim.strftime(fmt) if fim else None,
    )
    return _m._retool_state["all_selected"]

# ── session state ─────────────────────────────────────────────────────────────

if "manchetes" not in st.session_state:
    st.session_state.manchetes = {}
if "selecoes" not in st.session_state:
    st.session_state.selecoes = {}
if "expanded" not in st.session_state:
    st.session_state.expanded = {}
if "clipping" not in st.session_state:
    st.session_state.clipping = ""

# ── sidebar: parâmetros ───────────────────────────────────────────────────────

with st.sidebar:
    st.title("⚙️ Parâmetros")

    modo = st.selectbox("Modo de busca", ["Últimas X horas", "Janela específica"])

    horas = None
    inicio = None
    fim = None

    if modo == "Últimas X horas":
        horas = st.number_input("Horas", min_value=1, max_value=72, value=12, step=1)
    else:
        now = datetime.now()
        col1, col2 = st.columns(2)
        with col1:
            data_ini = st.date_input("Data início", value=now.date(), key="data_ini")
            hora_ini = st.time_input("Hora início", key="hora_ini",
                                     value=(now - timedelta(hours=12)).time()
                                     if "hora_ini" not in st.session_state else st.session_state["hora_ini"])
        with col2:
            data_fim = st.date_input("Data fim", value=now.date(), key="data_fim")
            hora_fim = st.time_input("Hora fim", key="hora_fim",
                                     value=now.time()
                                     if "hora_fim" not in st.session_state else st.session_state["hora_fim"])
        inicio = datetime.combine(data_ini, hora_ini)
        fim = datetime.combine(data_fim, hora_fim)

    buscar = st.button("🔍 Buscar manchetes", use_container_width=True, type="primary")

    if st.session_state.manchetes:
        _ABREV = {
            "Valor Econômico":    "Valor",
            "Estadão":            "Estadão",
            "O Globo":            "O Globo",
            "Folha de São Paulo": "Folha",
        }
        st.divider()
        # lê contagem direto dos checkboxes (session_state já atualizado)
        _counts = {}
        for _f, _items in st.session_state.manchetes.items():
            _counts[_f] = sum(
                1 for _i in _items
                if st.session_state.get(f"cb_{_f}_{_i['url']}", False)
            )
        _total = sum(_counts.values())
        st.markdown(f"### {_total} selecionados")
        for _f in st.session_state.manchetes:
            st.markdown(f"- **{_ABREV.get(_f, _f)}:** {_counts[_f]}")

# ── busca ─────────────────────────────────────────────────────────────────────

if buscar:
    import threading, io, sys, time, re as _re

    _URL_SOURCE = {
        "valor.globo.com":    "Valor Econômico",
        "estadao.com.br":     "Estadão",
        "oglobo.globo.com":   "O Globo",
        "folha.uol.com.br":   "Folha de São Paulo",
        "feeds.folha":        "Folha de São Paulo",
    }

    def _parse_logs(raw: str):
        """Transforma stdout bruto em linhas formatadas para exibição."""
        lines = []
        source_counts = {}
        current_source = None
        for line in raw.splitlines():
            line = line.strip()
            if not line:
                continue
            # detecta fonte pela URL
            for key, name in _URL_SOURCE.items():
                if key in line and "Buscando:" in line:
                    current_source = name
                    break
            # acumula contagem de itens por fonte
            m = _re.search(r"→ (\d+) itens", line)
            if m and current_source:
                source_counts[current_source] = source_counts.get(current_source, 0) + int(m.group(1))
            # clusters e janela
            if "Calculando clusters" in line:
                lines.append("📊  Calculando clusters e tendências...")
            elif "artigos na janela" in line:
                parts = _re.findall(r"\d+", line)
                if len(parts) >= 2:
                    lines.append(f"📰  {parts[0]} artigos na janela · {parts[1]} históricos")
            elif "clusters ativos" in line:
                parts = _re.findall(r"\d+", line)
                if len(parts) >= 2:
                    lines.append(f"🔗  {parts[0]} clusters ativos · {parts[1]} agrupados")
            elif "Buscando títulos" in line:
                parts = _re.findall(r"\d+", line)
                if parts:
                    lines.append(f"🔍  Buscando títulos reais em {parts[0]} artigos...")
            elif "Buscando datas" in line:
                parts = _re.findall(r"\d+", line)
                if parts:
                    lines.append(f"🕐  Verificando datas em {parts[0]} artigos...")

        # linha por fonte com contagem
        source_lines = [f"📡  {name}: {count} manchetes encontradas"
                        for name, count in source_counts.items()]
        return source_lines + [""] + lines if source_lines else lines

    resultado_pipeline = {}
    erro_pipeline = {}
    log_buffer = io.StringIO()

    class _Tee:
        def __init__(self, *streams):
            self.streams = streams
        def write(self, data):
            for s in self.streams:
                s.write(data)
        def flush(self):
            for s in self.streams:
                s.flush()

    original_stdout = sys.stdout
    sys.stdout = _Tee(original_stdout, log_buffer)

    def _run():
        try:
            if modo == "Últimas X horas":
                resultado_pipeline["data"] = _run_pipeline("horas", horas=int(horas))
            else:
                resultado_pipeline["data"] = _run_pipeline("janela", inicio=inicio, fim=fim)
        except Exception as e:
            erro_pipeline["msg"] = str(e)
        finally:
            sys.stdout = original_stdout

    t = threading.Thread(target=_run)
    t.start()

    clock_placeholder = st.empty()
    log_placeholder = st.empty()
    start = time.time()

    while t.is_alive():
        elapsed = int(time.time() - start)
        parsed = _parse_logs(log_buffer.getvalue())
        log_text = "\n".join(parsed) if parsed else "Iniciando..."
        clock_placeholder.markdown(
            f"<div style='text-align:center;font-size:3.5rem;font-weight:700;'>⏱ {elapsed}s</div>",
            unsafe_allow_html=True,
        )
        log_placeholder.code(log_text, language=None)
        time.sleep(1)

    t.join()
    clock_placeholder.empty()
    log_placeholder.empty()

    if "msg" in erro_pipeline:
        st.error(f"Erro ao buscar: {erro_pipeline['msg']}")
    else:
        manchetes = resultado_pipeline["data"]
        st.session_state.manchetes = manchetes
        st.session_state.selecoes = {name: set() for name in manchetes}
        st.session_state.expanded = {name: False for name in manchetes}
        st.session_state.clipping = ""
        st.success("Manchetes carregadas!")

# ── manchetes por fonte ───────────────────────────────────────────────────────

_SOURCE_BASE = {
    "Valor Econômico":    "https://valor.globo.com",
    "Estadão":            "https://www.estadao.com.br",
    "O Globo":            "https://oglobo.globo.com",
    "Folha de São Paulo": "https://www.folha.uol.com.br",
}

_EDITORIA_ORDER = {
    "Valor Econômico":    ["financas", "politica", "brasil", "mundo", "empresas", "impresso", "opiniao", "patrocinado"],
    "Estadão":            ["economia", "einvestidor", "politica", "internacional"],
    "O Globo":            ["economia", "politica", "brasil", "blog", "mundo", "rio"],
    "Folha de São Paulo": ["mercado", "poder", "cotidiano", "colunas", "tec", "ambiente"],
}

def _sort_editorias(fonte, editorias):
    order = _EDITORIA_ORDER.get(fonte, [])
    order_lower = [e.lower() for e in order]
    def key(e):
        k = e.lower()
        return (order_lower.index(k) if k in order_lower else len(order_lower), e)
    return sorted(editorias, key=key)

def _editoria_url(fonte, editoria):
    base = _SOURCE_BASE.get(fonte, "")
    if not base or not editoria:
        return None
    return f"{base}/{editoria}/"

manchetes = st.session_state.manchetes

def _render_bubble(fonte, item, checked, extra=False):
    lastmod = item.get("_date_from_lastmod", False)
    dt = item.get("dt")
    horario = dt.astimezone(_m.LOCAL_TZ).strftime("%d/%m/%Y %H:%M") if dt else "sem data"
    horario_label = f"última edição · {horario}" if lastmod and horario else horario
    title = item["title"]
    url = item["url"]
    cb_key = f"cb_{fonte}_{url}"

    col_check, col_bubble = st.columns([0.04, 0.96])
    with col_check:
        sel = st.checkbox("sel", value=checked, key=cb_key, label_visibility="collapsed")
        if sel:
            st.session_state.selecoes.setdefault(fonte, set()).add(url)
        else:
            st.session_state.selecoes.get(fonte, set()).discard(url)
    # lê o valor atual do checkbox (já atualizado pelo Streamlit antes do rerun)
    is_checked = st.session_state.get(cb_key, checked)
    with col_bubble:
        if extra:
            bubble_class = "bubble extra selected" if is_checked else "bubble extra"
        else:
            bubble_class = "bubble selected" if is_checked else "bubble"
        st.markdown(
            f"""<div class="{bubble_class}">
                <div class="bubble-title">{title}</div>
                <div class="bubble-meta">{horario_label} &nbsp;·&nbsp; <a href="{url}" target="_blank">link</a></div>
            </div>""",
            unsafe_allow_html=True,
        )


if manchetes:
    st.markdown('<div class="section-header">📰 Manchetes</div>', unsafe_allow_html=True)

    col_filtro, col_ordem = st.columns([0.5, 0.5])
    with col_filtro:
        filtro = st.radio("Ordenar por", ["Editoria", "Ranking", "Data/hora"], horizontal=True)
    with col_ordem:
        if filtro == "Data/hora":
            ordem = st.radio("Ordem", ["Mais recentes primeiro", "Mais antigas primeiro"], horizontal=True)
        else:
            ordem = None

    def _render_group(fonte, group_items, extras_set=None):
        if extras_set is None:
            extras_set = set()
        if filtro == "Ranking":
            for item in group_items:
                checked = item["url"] in st.session_state.selecoes.get(fonte, set())
                _render_bubble(fonte, item, checked, extra=(item["url"] in extras_set))
        elif filtro == "Data/hora":
            ordenados = sorted(group_items, key=lambda x: x.get("dt") or datetime.min.replace(tzinfo=timezone.utc), reverse=(ordem == "Mais recentes primeiro"))
            for item in ordenados:
                checked = item["url"] in st.session_state.selecoes.get(fonte, set())
                _render_bubble(fonte, item, checked, extra=(item["url"] in extras_set))
        elif filtro == "Editoria":
            grupos: dict = {}
            for item in group_items:
                editoria = _m._article_topic(item["url"]) or "outras"
                grupos.setdefault(editoria, []).append(item)
            for editoria, grupo in [(e, grupos[e]) for e in _sort_editorias(fonte, grupos.keys())]:
                ed_url = _editoria_url(fonte, editoria)
                ed_label = editoria.replace('-', ' ').title()
                if ed_url:
                    st.markdown(f'**<a href="{ed_url}" target="_blank" style="color:inherit;text-decoration:none;">{ed_label} ↗</a>**', unsafe_allow_html=True)
                else:
                    st.markdown(f"**{ed_label}**")
                for item in grupo:
                    checked = item["url"] in st.session_state.selecoes.get(fonte, set())
                    _render_bubble(fonte, item, checked, extra=(item["url"] in extras_set))

    for fonte, items in manchetes.items():
        fonte_url = _SOURCE_BASE.get(fonte, "")
        if fonte_url:
            st.markdown(
                f'<div class="fonte-header"><a href="{fonte_url}" target="_blank" style="color:inherit;text-decoration:none;">{fonte} ↗</a></div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(f'<div class="fonte-header">{fonte}</div>', unsafe_allow_html=True)

        expandido = st.session_state.expanded.get(fonte, False)
        principais = items[:20]
        extras = items[20:]
        extras_urls = {item["url"] for item in extras}

        visiveis = items if expandido else principais
        _render_group(fonte, visiveis, extras_set=extras_urls if expandido else set())

        if extras:
            if not expandido:
                if st.button(f"+ {len(extras)} artigos adicionais ↓", key=f"expand_{fonte}"):
                    st.session_state.expanded[fonte] = True
                    st.rerun()
            else:
                if st.button("Ver menos ↑", key=f"collapse_{fonte}"):
                    st.session_state.expanded[fonte] = False
                    st.rerun()

        st.divider()

    # ── ações ─────────────────────────────────────────────────────────────────

    total = sum(len(v) for v in st.session_state.selecoes.values())
    st.write(f"**{total} artigo(s) selecionado(s)**")

    col_clip, col_plan = st.columns(2)

    with col_clip:
        if st.button("📋 Gerar clipping", use_container_width=True, type="primary"):
            selecoes_lista = {k: list(v) for k, v in st.session_state.selecoes.items()}
            resultado = _m.retool_gerar_clipping(selecoes_lista)
            st.session_state.clipping = resultado["clipping"]

    with col_plan:
        if st.button("📊 Gerar planilha", use_container_width=True):
            selecoes_lista = {k: list(v) for k, v in st.session_state.selecoes.items()}
            resultado = _m.retool_gerar_planilha(selecoes_lista)
            import base64
            xlsx_bytes = base64.b64decode(resultado["xlsx_base64"])
            st.download_button(
                label="⬇️ Baixar planilha",
                data=xlsx_bytes,
                file_name=resultado["filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    if st.session_state.clipping:
        st.subheader("Clipping")
        st.code(st.session_state.clipping, language=None)

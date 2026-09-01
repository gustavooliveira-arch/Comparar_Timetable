import streamlit as st

from compare import (
    SHEET_NAME,
    compare_by_keys,
    compare_positional,
    load_timetable,
)
from export import build_updated_excel

st.set_page_config(page_title="Comparar TIMETABLE", layout="wide")
st.title("Comparar TIMETABLE")
st.caption("Envie dois arquivos Excel e veja o que mudou na aba TIMETABLE.")

col_old, col_new = st.columns(2)
with col_old:
    file_old = st.file_uploader("Arquivo antigo (antes)", type=["xlsx", "xls"], key="old")
with col_new:
    file_new = st.file_uploader("Arquivo novo (depois)", type=["xlsx", "xls"], key="new")

sheet = st.text_input("Nome da aba", value=SHEET_NAME)

if not file_old or not file_new:
    st.session_state.pop("compare", None)
    st.info("Faça o upload dos dois arquivos para começar.")
    st.stop()

file_old.seek(0)
file_new.seek(0)
try:
    df_old = load_timetable(file_old, sheet_name=sheet)
    df_new = load_timetable(file_new, sheet_name=sheet)
except Exception as exc:
    st.error(str(exc))
    st.stop()

st.success(
    f"Arquivo antigo: {len(df_old)} linhas × {len(df_old.columns)} colunas  ·  "
    f"Arquivo novo: {len(df_new)} linhas × {len(df_new.columns)} colunas"
)

shared_cols = [c for c in df_old.columns if c in df_new.columns]
mode = st.radio(
    "Como comparar",
    options=["posição", "chave"],
    format_func=lambda m: (
        "Por posição (mesmo layout, linha a linha)"
        if m == "posição"
        else "Por chave (linhas podem ter sido reordenadas)"
    ),
    horizontal=True,
)

keys: list[str] = []
if mode == "chave":
    keys = st.multiselect(
        "Colunas que identificam uma linha (ex.: data + horário, código da aula)",
        options=shared_cols,
        default=shared_cols[:1] if shared_cols else [],
    )
    if not keys:
        st.warning("Selecione pelo menos uma coluna-chave.")
        st.stop()

if st.button("Comparar", type="primary"):
    result = (
        compare_by_keys(df_old, df_new, keys)
        if mode == "chave"
        else compare_positional(df_old, df_new)
    )
    file_old.seek(0)
    st.session_state.compare = {
        "result": result,
        "old_bytes": file_old.getvalue(),
        "keys": keys,
        "sheet": sheet,
        "mode": mode,
    }

compare = st.session_state.get("compare")
if not compare:
    with st.expander("Prévia das planilhas"):
        left, right = st.columns(2)
        with left:
            st.subheader("Antigo")
            st.dataframe(df_old.head(20), use_container_width=True, hide_index=True)
        with right:
            st.subheader("Novo")
            st.dataframe(df_new.head(20), use_container_width=True, hide_index=True)
    st.stop()

result = compare["result"]

if result.duplicate_keys_old or result.duplicate_keys_new:
    st.warning(
        "Há chaves duplicadas "
        f"(antigo: {result.duplicate_keys_old}, novo: {result.duplicate_keys_new}). "
        "Só a primeira ocorrência de cada chave foi usada na comparação."
    )

if result.extra_columns_old:
    st.write("Colunas só no arquivo antigo:", ", ".join(result.extra_columns_old))
if result.extra_columns_new:
    st.write("Colunas só no arquivo novo:", ", ".join(result.extra_columns_new))

if not result.has_changes:
    st.success("Nenhuma diferença encontrada na aba TIMETABLE.")
    st.stop()

c1, c2, c3 = st.columns(3)
c1.metric("Células alteradas", len(result.modified))
c2.metric("Linhas adicionadas", len(result.added_rows))
c3.metric("Linhas removidas", len(result.removed_rows))

xlsx_bytes = build_updated_excel(
    compare["old_bytes"],
    result,
    compare["sheet"],
    key_cols=compare["keys"] or None,
)
st.download_button(
    "Baixar Excel atualizado",
    data=xlsx_bytes,
    file_name="timetable_atualizado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    help="Cópia do arquivo antigo, com a TIMETABLE já alterada. Amarelo = célula mudou, verde = linha nova.",
)

tab_mod, tab_add, tab_rem, tab_prev = st.tabs(
    ["Alterações", "Adicionadas", "Removidas", "Prévia das planilhas"]
)
with tab_mod:
    modified = result.modified_frame()
    if modified.empty:
        st.write("Nenhuma célula alterada nas colunas em comum.")
    else:
        st.dataframe(modified, use_container_width=True, hide_index=True)
with tab_add:
    if result.added_rows.empty:
        st.write("Nenhuma linha adicionada.")
    else:
        st.dataframe(result.added_rows, use_container_width=True, hide_index=True)
with tab_rem:
    if result.removed_rows.empty:
        st.write("Nenhuma linha removida.")
    else:
        st.dataframe(result.removed_rows, use_container_width=True, hide_index=True)
with tab_prev:
    left, right = st.columns(2)
    with left:
        st.subheader("Antigo")
        st.dataframe(df_old, use_container_width=True, hide_index=True)
    with right:
        st.subheader("Novo")
        st.dataframe(df_new, use_container_width=True, hide_index=True)

import io

import streamlit as st
import pandas as pd

from compare import (
    SHEET_NAME,
    compare_positional,
    load_timetable,
)

from export import (
    build_diff_only_excel,
    build_updated_excel,
)
from cenario import build_cenario_excel

# ============================================================
# CONFIGURAÇÃO
# ============================================================

st.set_page_config(
    page_title="Comparar TIMETABLE",
    layout="wide"
)

st.title("Comparar TIMETABLE")

st.caption(
    "Envie dois arquivos Excel e veja o que mudou na aba TIMETABLE."
)

# ============================================================
# UPLOAD DOS ARQUIVOS
# ============================================================

col_old, col_new = st.columns(2)

with col_old:

    file_old = st.file_uploader(
        "Arquivo antigo",
        type=["xlsx", "xls"],
        key="old"
    )

with col_new:

    file_new = st.file_uploader(
        "Arquivo novo",
        type=["xlsx", "xls"],
        key="new"
    )

# ============================================================
# ABA
# ============================================================

sheet = st.text_input(
    "Nome da aba",
    value=SHEET_NAME
)

# ============================================================
# VERIFICAÇÃO DOS UPLOADS
# ============================================================

if not file_old or not file_new:

    st.session_state.pop(
        "compare",
        None
    )

    st.info(
        "Faça o upload dos dois arquivos para começar."
    )

    st.stop()

# ============================================================
# CARREGAMENTO (com cache)
#
# Sem cache, `load_timetable` reprocessava os dois arquivos Excel
# inteiros a CADA interação da página (zoom, seleção de célula etc.),
# porque o Streamlit reexecuta o script do início a cada rerun.
#
# `st.cache_data` guarda o resultado por conteúdo (hash dos bytes +
# nome da aba), então o parsing do Excel só roda de novo se o
# arquivo ou a aba realmente mudarem.
# ============================================================


@st.cache_data(show_spinner="Lendo planilha...")
def _carregar_timetable_cache(conteudo: bytes, sheet_name: str):
    return load_timetable(
        io.BytesIO(conteudo),
        sheet_name=sheet_name,
    )


@st.cache_data(show_spinner="Lendo aba FROTA...")
def _carregar_frota_cache(conteudo: bytes):
    """Aba FROTA (sem cabeçalho) do arquivo novo.

    Cachead pelos bytes do arquivo: sem isso, o Excel inteiro era
    reparsed a CADA interação da página (zoom, clique numa célula
    da prévia), deixando a navegação lenta para arquivos grandes.
    """
    return pd.read_excel(
        io.BytesIO(conteudo),
        sheet_name="FROTA",
        header=None,
    )


file_old.seek(0)
file_new.seek(0)

try:

    df_old = _carregar_timetable_cache(
        file_old.getvalue(),
        sheet
    )

    df_new = _carregar_timetable_cache(
        file_new.getvalue(),
        sheet
    )

    # A aba FROTA também vem do arquivo novo.
    # Não alteramos df_new: ele continua sendo a TIMETABLE usada
    # na comparação e na prévia.
    df_FROTA = _carregar_frota_cache(
        file_new.getvalue(),
    )

except Exception as exc:

    st.error(str(exc))

    st.stop()

# ============================================================
# INFORMAÇÕES DOS ARQUIVOS
# ============================================================

st.success(
    f"Arquivo antigo: "
    f"{len(df_old)} linhas × "
    f"{len(df_old.columns)} colunas  ·  "

    f"Arquivo novo: "
    f"{len(df_new)} linhas × "
    f"{len(df_new.columns)} colunas"
)

# ============================================================
# UTILITÁRIOS
# ============================================================

def normalize_column_name(name):

    return (
        str(name)
        .strip()
        .casefold()
    )


# ============================================================

def destacar_diferencas_novo(df, result):
    """
    Destaca diferenças somente no arquivo NOVO.

    Cores:
        Amarelo = célula alterada
        Verde   = célula adicionada
        Vermelho = célula removida
    """

    df_view = df.copy()

    estilos = pd.DataFrame(
        "",
        index=df_view.index,
        columns=df_view.columns
    )

    # =========================================
    # CÉLULAS ALTERADAS
    # =========================================
    for change in result.modified:
        linha = change.row
        coluna = change.column

        if linha in df_view.index and coluna in df_view.columns:
            estilos.loc[linha, coluna] = (
                "background-color: #ffb84f;"
                "color: #000000;"
            )

    # =========================================
    # CÉLULAS ADICIONADAS
    # =========================================
    for change in result.added_cells:
        linha = change.row
        coluna = change.column

        if linha in df_view.index and coluna in df_view.columns:
            estilos.loc[linha, coluna] = (
                "background-color: #4ac58a;"
                "color: #000000;"
            )

    # =========================================
    # CÉLULAS REMOVIDAS
    # =========================================
    for change in result.removed_cells:
        linha = change.row
        coluna = change.column

        if linha in df_view.index and coluna in df_view.columns:
            estilos.loc[linha, coluna] = (
                "background-color: #d83c61;"
                "color: #000000;"
            )

    # =========================================
    # LINHAS NOVAS
    #
    # Como uma linha adicionada pode não possuir
    # registros em result.added_cells, destacamos
    # a linha inteira de verde.
    # =========================================
    for linha in result.added_rows.index:
        if linha in df_view.index:
            for coluna in df_view.columns:
                estilos.loc[linha, coluna] = (
                    "background-color: #4ac58a;"
                    "color: #000000;"
                )

    return df_view.style.apply(
        lambda _: estilos,
        axis=None
    )


def _coluna_correspondente(coluna, origem, destino):
    if coluna in destino.columns:
        return coluna

    alvo = normalize_column_name(coluna)

    for nome in destino.columns:
        if normalize_column_name(nome) == alvo:
            return nome

    origem_cols = list(origem.columns)
    destino_cols = list(destino.columns)

    if coluna in origem_cols:
        indice = origem_cols.index(coluna)
        if indice < len(destino_cols):
            return destino_cols[indice]

    return None


def _montar_mapa_colunas(origem, destino):
    """Pré-computa a coluna correspondente em `destino` para cada
    coluna de `origem` (mesmo nome, nome normalizado ou mesma
    posição).

    Mesmo comportamento de `_coluna_correspondente`, mas calculado
    UMA VEZ por comparação. Sem isso, cada clique numa célula da
    prévia refazia esse mapeamento (análise de nomes + busca linear)
    para sincronizar a seleção na outra planilha.
    """
    destino_set = set(destino.columns)
    dest_norm = {
        normalize_column_name(n): n
        for n in destino.columns
    }
    destino_list = list(destino.columns)
    origem_list = list(origem.columns)

    mapa = {}

    for i, col in enumerate(origem_list):

        if col in destino_set:
            mapa[col] = col
            continue

        norm = normalize_column_name(col)

        if norm in dest_norm:
            mapa[col] = dest_norm[norm]
        elif i < len(destino_list):
            mapa[col] = destino_list[i]
        else:
            mapa[col] = None

    return mapa


def _linha_correspondente(linha, origem, destino):
    if linha is None or linha < 0 or linha >= len(origem):
        return None

    if linha < len(destino):
        return linha

    return None


def _celula_correspondente(celula, origem, destino, col_map=None):
    if not celula or len(celula) != 2:
        return None

    linha, coluna = celula[0], celula[1]

    if col_map is not None:
        coluna_dest = col_map.get(coluna)
    else:
        coluna_dest = _coluna_correspondente(coluna, origem, destino)

    linha_dest = _linha_correspondente(linha, origem, destino)

    if coluna_dest is None or linha_dest is None:
        return None

    return [linha_dest, coluna_dest]


def _celulas_do_widget(chave_widget):
    estado = st.session_state.get(chave_widget)

    if not estado:
        return []

    try:
        return list(estado["selection"]["cells"])
    except (KeyError, TypeError):
        return []


def _sincronizar_selecao_previa(
    origem_key,
    destino_key,
    origem,
    destino,
    col_map=None,
):
    celulas = _celulas_do_widget(origem_key)

    if not celulas:
        st.session_state[destino_key] = {
            "selection": {"cells": []}
        }
        return

    correspondente = _celula_correspondente(
        celulas[0],
        origem,
        destino,
        col_map,
    )

    st.session_state[destino_key] = {
        "selection": {
            "cells": [correspondente] if correspondente else []
        }
    }

# ============================================================
# COMPARAR
# ============================================================

if st.button(
    "Comparar",
    type="primary"
):

    try:

        result = compare_positional(
            df_old,
            df_new,
        )

    except Exception as exc:

        st.error(
            f"Erro ao comparar os arquivos: {exc}"
        )

        st.stop()


    # Volta para o início do arquivo antigo
    file_old.seek(0)

    # Versão da comparação: usada como chave de cache para não
    # regerar os Excels/estilos em reruns que não mudam o resultado
    # (zoom, seleção de célula, gerar cenário etc.).
    st.session_state.compare_version = (
        st.session_state.get("compare_version", 0) + 1
    )

    st.session_state.compare = {

        "result": result,

        "old_bytes": file_old.getvalue(),

        "sheet": sheet,

        # Mapas de coluna pré-calculados (O(1) por clique na prévia).
        "col_old_to_new": _montar_mapa_colunas(df_old, df_new),

        "col_new_to_old": _montar_mapa_colunas(df_new, df_old),

        "version": st.session_state.compare_version,
    }

    st.session_state.pop("preview_old", None)
    st.session_state.pop("preview_new", None)

    # Invalida os caches de Excel/estilo da comparação anterior
    st.session_state.pop("_excel_cache_key", None)
    st.session_state.pop("_style_cache_key", None)


# ============================================================
# RESULTADO DA COMPARAÇÃO
# ============================================================

compare = st.session_state.get(
    "compare"
)

if not compare:

    with st.expander(
        "Prévia das planilhas"
    ):

        left, right = st.columns(2)

        with left:

            st.subheader("Antigo")

            st.dataframe(
                df_old.head(65),
                use_container_width=True,
                hide_index=True
            )

        with right:

            st.subheader("Novo")

            st.dataframe(
                df_new.head(65),
                use_container_width=True,
                hide_index=True
            )

    st.stop()

# ============================================================
# RESULTADO
# ============================================================

result = compare["result"]


# ============================================================
# COLUNAS QUE EXISTEM APENAS EM UM ARQUIVO
# ============================================================

if (
    result.extra_columns_old
    or result.extra_columns_new
):

    st.subheader(
        "Colunas diferentes entre os arquivos"
    )

    column_rows = []


    # --------------------------------------------------------
    # COLUNAS APENAS NO ANTIGO
    # --------------------------------------------------------

    for column in result.extra_columns_old:

        column_rows.append(
            {
                "Coluna": column.name,
                "Arquivo": "Antigo",
                "Visibilidade": column.visibility,
            }
        )


    # --------------------------------------------------------
    # COLUNAS APENAS NO NOVO
    # --------------------------------------------------------

    for column in result.extra_columns_new:

        column_rows.append(
            {
                "Coluna": column.name,
                "Arquivo": "Novo",
                "Visibilidade": column.visibility,
            }
        )


    if column_rows:

        df_columns = pd.DataFrame(
            column_rows
        )

        st.dataframe(
            df_columns,
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# NENHUMA ALTERAÇÃO
# ============================================================

if not result.has_changes:

    st.success(
        "Nenhuma diferença encontrada "
        "na aba TIMETABLE."
    )

    st.stop()


# ============================================================
# MÉTRICAS
# ============================================================

c1, c2, c3 = st.columns(3)

c1.metric(
    "Células alteradas",
    len(result.modified)
)

c2.metric(
    "Linhas adicionadas",
    len(result.added_rows)
)

c3.metric(
    "Linhas removidas",
    len(result.removed_rows)
)

# ============================================================
# GERAÇÃO DOS EXCELS (com cache por versão da comparação)
#
# Antes, esses dois arquivos eram reconstruídos do zero em TODA
# interação da página (zoom, clique em célula, gerar cenário...).
# Agora só são regenerados quando uma nova comparação é feita
# (`compare["version"]` muda), e ficam guardados em session_state
# no restante do tempo.
# ============================================================

excel_cache_key = compare["version"]

if st.session_state.get("_excel_cache_key") != excel_cache_key:

    with st.spinner("Gerando arquivos Excel para download..."):

        st.session_state.xlsx_bytes = build_updated_excel(
            compare["old_bytes"],
            result,
            compare["sheet"],
        )

        st.session_state.diff_only_bytes = build_diff_only_excel(
            compare["old_bytes"],
            result,
            compare["sheet"],
        )

    st.session_state._excel_cache_key = excel_cache_key

xlsx_bytes = st.session_state.xlsx_bytes
diff_only_bytes = st.session_state.diff_only_bytes

# ============================================================
# DOWNLOAD
# ============================================================

col_dl1, col_dl2, _spacer = st.columns(
    [1, 2.5, 3],
    gap="small"
)

with col_dl1:

    st.download_button(

        "Baixar Excel atualizado",

        data=xlsx_bytes,

        file_name="timetable_atualizado.xlsx",

        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),

        type="primary",

        help=(
            "Cópia do arquivo antigo, "
            "com a TIMETABLE já alterada. "
            "Amarelo = célula mudou, "
            "verde = linha nova."
        ),
    )


with col_dl2:

    st.download_button(

        "Baixar apenas alterações",

        data=diff_only_bytes,

        file_name=(
            "timetable_apenas_alteracoes.xlsx"
        ),

        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),

        help=(
            "Mesma configuração da tabela antiga, "
            "mas só com as células que mudaram "
            "ou linhas novas. "
            "Sem linhas removidas."
        ),
    )


# ============================================================
# ABAS DO RESULTADO
# ============================================================

tab_mod, tab_add, tab_rem, tab_prev = st.tabs(
    [
        "Alterações",
        "Adicionadas",
        "Removidas",
        "Prévia das planilhas",
    ]
)


# ============================================================
# ALTERAÇÕES
# ============================================================

with tab_mod:

    modified = result.modified_frame()

    if modified.empty:

        st.write(
            "Nenhuma célula alterada "
            "nas colunas em comum."
        )

    else:

        st.dataframe(
            modified,
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# LINHAS ADICIONADAS
# ============================================================

with tab_add:

    if result.added_rows.empty:

        st.write(
            "Nenhuma linha adicionada."
        )

    else:

        st.dataframe(
            result.added_rows,
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# LINHAS REMOVIDAS
# ============================================================

with tab_rem:

    if result.removed_rows.empty:

        st.write(
            "Nenhuma linha removida."
        )

    else:

        st.dataframe(
            result.removed_rows,
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# PRÉVIA
# ============================================================

with tab_prev:

    st.markdown(
        """
        <style>
        .block-container {
            max-width: 96% !important;
            padding-left: 1.4rem;
            padding-right: 1.4rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    if "preview_zoom" not in st.session_state:
        st.session_state.preview_zoom = 110

    z1, z2, z3, z4 = st.columns([1.1, 1.1, 1.1, 1.1])

    with z1:
        if st.button("Diminuir zoom", key="preview_zoom_out", use_container_width=True):
            st.session_state.preview_zoom = max(
                80, st.session_state.preview_zoom - 10
            )

    with z2:
        if st.button("Aumentar zoom", key="preview_zoom_in", use_container_width=True):
            st.session_state.preview_zoom = min(
                180, st.session_state.preview_zoom + 10
            )

    zoom_pct = st.session_state.preview_zoom
    zoom = zoom_pct / 100
    row_height = max(24, int(36 * zoom))
    col_width = max(90, int(160 * zoom))
    table_height = max(520, int(560 * zoom))

    with z3:
        gerar = st.button(
            "Gerar Novo Cenário",
            key="gerar_cenario",
            type="primary",
            use_container_width=True
        )

    if gerar:
        try:
            st.session_state.cenario_bytes = build_cenario_excel(
                df_new,
                frota=df_FROTA,
                timetable_old=df_old,
            )
            st.session_state.cenario_error = None
        except Exception as exc:
            st.session_state.cenario_bytes = None
            st.session_state.cenario_error = str(exc)

    with z4:
        if st.session_state.get("cenario_bytes"):
            st.download_button(
                "Baixar novo cenário",
                data=st.session_state.cenario_bytes,
                file_name="novo_cenario.xlsx",
                mime=(
                    "application/"
                    "vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                key="dl_cenario",
                use_container_width=True
            )

    if st.session_state.get("cenario_error"):
        st.error(st.session_state.cenario_error)

    # st.caption(
    #   "Clique numa célula para selecionar a correspondente na outra planilha. "
        #    "O cenário usa a planilha Nova e o modelo Modelo Cenario.xlsx. "
        #   "Etapa = 1. T.P. = 01:30 se Local (REC.) começa com GR; senão 00:45."
    # )

    old_cfg = {
        coluna: st.column_config.TextColumn(width=col_width)
        for coluna in df_old.columns
    }
    new_cfg = {
        coluna: st.column_config.TextColumn(width=col_width)
        for coluna in df_new.columns
    }

    left, right = st.columns(2, gap="small")

    # =========================================
    # ESTILO DO ARQUIVO NOVO (com cache por versão da comparação)
    #
    # Montar o Styler percorre célula a célula; sem cache isso
    # rodava de novo a cada clique de zoom/seleção, mesmo sem a
    # comparação ter mudado.
    # =========================================
    style_cache_key = compare["version"]

    if st.session_state.get("_style_cache_key") != style_cache_key:
        st.session_state.df_new_styled = destacar_diferencas_novo(
            df_new,
            result,
        )
        st.session_state._style_cache_key = style_cache_key

    df_new_styled = st.session_state.df_new_styled

    # =========================================
    # ARQUIVO ANTIGO — SEM DESTAQUE
    # =========================================
    with left:
        st.subheader("Antigo")

        st.dataframe(
            df_old,
            use_container_width=True,
            hide_index=True,
            key="preview_old",
            on_select=lambda: _sincronizar_selecao_previa(
                "preview_old",
                "preview_new",
                df_old,
                df_new,
                compare.get("col_old_to_new"),
            ),
            selection_mode="single-cell",
            height=table_height,
            row_height=row_height,
            column_config=old_cfg,
        )

    # =========================================
    # ARQUIVO NOVO — COM DESTAQUE
    # =========================================
    with right:
        st.subheader("Novo")

        st.dataframe(
            df_new_styled,
            use_container_width=True,
            hide_index=True,
            key="preview_new",
            on_select=lambda: _sincronizar_selecao_previa(
                "preview_new",
                "preview_old",
                df_new,
                df_old,
                compare.get("col_new_to_old"),
            ),
            selection_mode="single-cell",
            height=table_height,
            row_height=row_height,
            column_config=new_cfg,
        )

    # =========================================
    # LEGENDA
    # =========================================
    st.markdown(
        """
        **Legenda:**

        🟨 Alterado &nbsp;&nbsp;&nbsp;
        🟩 Adicionado &nbsp;&nbsp;&nbsp;
        🟥 Removido
        """
    )

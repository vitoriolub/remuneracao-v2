import io
import unicodedata

import pandas as pd
import streamlit as st

# ==================================================
# CONFIGURAÇÃO
# ==================================================

st.set_page_config(
    page_title="Remuneração Comercial",
    layout="wide"
)

VERDE = "#D4EDDA"
AMARELO = "#FFF3CD"
VERMELHO = "#F8D7DA"

# ==================================================
# FUNÇÕES AUXILIARES
# ==================================================

def sem_acento(texto):
    """'Índice Litragem' -> 'INDICE LITRAGEM'"""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.strip().upper()


@st.cache_data(show_spinner=False)
def listar_abas(conteudo):
    return pd.ExcelFile(io.BytesIO(conteudo)).sheet_names


@st.cache_data(show_spinner=False)
def ler_excel(conteudo, aba):
    return pd.read_excel(io.BytesIO(conteudo), sheet_name=aba)


def normalizar_colunas(df):
    """Padroniza Supervisor / Vendedor / TOTAL FINAL.
    Só a primeira coluna que bater é renomeada, para não gerar nomes duplicados
    (ex.: 'Supervisor' e 'Cód. Supervisor')."""

    mapa = {}
    usados = set()

    for col in df.columns:

        nome = sem_acento(col)

        if nome.startswith("COD") or nome.startswith("ID "):
            continue

        if "SUPERV" in nome:
            destino = "Supervisor"
        elif "VENDEDOR" in nome or "CONSULTOR" in nome:
            destino = "Vendedor"
        elif "TOTAL FINAL" in nome:
            destino = "TOTAL FINAL"
        else:
            continue

        if destino not in usados and destino not in df.columns:
            mapa[col] = destino
            usados.add(destino)

    return df.rename(columns=mapa)


def numero_br(valor, casas=2):
    return (
        f"{valor:,.{casas}f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def eh_percentual_pelo_nome(nome):
    nome = sem_acento(nome)
    if "INDICE" in nome or "%" in nome:
        return True
    # Inadimplência Meta / Realizado já vêm em % (1,75 = 1,75%)
    if "INADIMPL" in nome and "VARIAVEL" not in nome:
        return True
    return False


def eh_moeda_pelo_nome(nome):
    nome = sem_acento(nome)
    return "VARIAVEL" in nome or "TOTAL" in nome or "R$" in nome


def menor_eh_melhor(nome):
    return "INADIMPL" in sem_acento(nome)


def escala_percentual(nome, serie):
    """'fracao' quando 0,85 = 85% ; 'pontos' quando 85 = 85%."""
    if menor_eh_melhor(nome):
        return "pontos"
    maximo = pd.to_numeric(serie, errors="coerce").abs().max()
    if pd.notna(maximo) and maximo <= 5:
        return "fracao"
    return "pontos"


def achar_coluna_meta(nome, colunas):
    """'Inadimplência Realizado' -> 'Inadimplência Meta'"""
    alvo = sem_acento(nome).replace("REALIZADO", "META")
    if alvo == sem_acento(nome):
        return None
    for col in colunas:
        if sem_acento(col) == alvo:
            return col
    return None


def cor_texto_para(fundo):
    """Preto ou branco, o que tiver mais contraste com a cor de fundo."""
    fundo = fundo.lstrip("#")
    r, g, b = (int(fundo[i:i + 2], 16) for i in (0, 2, 4))
    luminancia = 0.299 * r + 0.587 * g + 0.114 * b
    return "#000000" if luminancia > 150 else "#FFFFFF"

# ==================================================
# SIDEBAR
# ==================================================

st.sidebar.title("Configurações")

arquivo = st.sidebar.file_uploader(
    "Selecione o Excel",
    type=["xlsx", "xls"]
)

cor_principal = st.sidebar.color_picker(
    "Cor do cabeçalho",
    "#1E88E5"
)

tamanho_fonte = st.sidebar.slider(
    "Tamanho da fonte da tabela",
    min_value=10,
    max_value=20,
    value=13
)

with st.sidebar.expander("Faixas do semáforo"):

    limite_vermelho = st.number_input(
        "Vermelho abaixo de (%)",
        value=80,
        step=5
    )

    limite_verde = st.number_input(
        "Verde a partir de (%)",
        value=100,
        step=5
    )

    st.caption(
        "Colunas de inadimplência usam a lógica inversa: "
        "verde quando o realizado fica igual ou abaixo da meta."
    )

st.title("📊 Remuneração Comercial")

if not arquivo:
    st.info("Selecione um arquivo Excel para iniciar.")
    st.stop()

# ==================================================
# LEITURA
# ==================================================

conteudo = arquivo.getvalue()

try:
    abas = listar_abas(conteudo)
except Exception as erro:
    st.error(f"Não consegui abrir o arquivo: {erro}")
    st.stop()

aba = abas[0]
if len(abas) > 1:
    aba = st.sidebar.selectbox("Aba da planilha", abas)

try:
    df = ler_excel(conteudo, aba)
except Exception as erro:
    st.error(f"Erro ao ler a aba '{aba}': {erro}")
    st.stop()

df = normalizar_colunas(df)

with st.expander("Tabela original"):
    st.dataframe(df, hide_index=True)

# ==================================================
# FILTROS
# ==================================================

st.subheader("Filtros")

df_filtrado = df.copy()
supervisor = "Todos"

col_sup, col_vend = st.columns([1, 3])

with col_sup:

    if "Supervisor" in df.columns:

        supervisor = st.selectbox(
            "Supervisor",
            ["Todos"] + sorted(df["Supervisor"].dropna().astype(str).unique())
        )

        if supervisor != "Todos":
            df_filtrado = df_filtrado[
                df_filtrado["Supervisor"].astype(str) == supervisor
            ]

with col_vend:

    if "Vendedor" in df.columns:

        lista_vendedores = sorted(
            df_filtrado["Vendedor"].dropna().astype(str).unique()
        )

        # Uma seleção por supervisor: ao trocar de supervisor,
        # todos os consultores dele já vêm marcados.
        chave = f"vendedores_{supervisor}"

        if chave not in st.session_state:
            st.session_state[chave] = lista_vendedores

        def marcar(valor, chave=chave):
            st.session_state[chave] = valor

        b1, b2, _ = st.columns([1, 1, 4])
        b1.button("Marcar todos", on_click=marcar, args=(lista_vendedores,))
        b2.button("Desmarcar todos", on_click=marcar, args=([],))

        vendedores = st.pills(
            f"Consultores ({len(lista_vendedores)}) — clique para incluir/remover",
            lista_vendedores,
            selection_mode="multi",
            key=chave
        )

        df_filtrado = df_filtrado[
            df_filtrado["Vendedor"].astype(str).isin(vendedores)
        ]

# ==================================================
# COLUNAS
# ==================================================

with st.expander("Colunas da tabela", expanded=False):

    colunas = st.multiselect(
        "Escolha as colunas (a ordem aqui é a ordem na tabela)",
        df.columns.tolist(),
        default=df.columns.tolist(),
        key="colunas"
    )

    numericas = [
        c for c in colunas
        if pd.api.types.is_numeric_dtype(df[c])
    ]

    colunas_percentuais = st.multiselect(
        "Colunas em % (recebem formatação % e semáforo automaticamente)",
        numericas,
        default=[c for c in numericas if eh_percentual_pelo_nome(c)],
        key="colunas_percentuais"
    )

    st.markdown("**Renomear colunas**")

    novos_nomes = {}
    grade = st.columns(3)

    for i, coluna in enumerate(colunas):
        novos_nomes[coluna] = grade[i % 3].text_input(
            coluna,
            value=str(coluna),
            key=f"nome_{coluna}"
        ).strip() or str(coluna)

if not colunas:
    st.warning("Escolha pelo menos uma coluna.")
    st.stop()

# Evita nomes repetidos depois de renomear (quebraria a tabela)
vistos = {}
for coluna, nome in novos_nomes.items():
    if nome in vistos:
        vistos[nome] += 1
        novos_nomes[coluna] = f"{nome} ({vistos[nome]})"
    else:
        vistos[nome] = 1

tabela = df_filtrado[colunas]
tabela_final = tabela.rename(columns=novos_nomes)

# ==================================================
# FORMATAÇÃO
# ==================================================

formatacao = {}
escalas = {}

for col in colunas:

    if not pd.api.types.is_numeric_dtype(tabela[col]):
        continue

    destino = novos_nomes[col]

    if col in colunas_percentuais:

        escala = escala_percentual(col, df[col])
        escalas[col] = escala

        if escala == "fracao":
            formatacao[destino] = (
                lambda v: "" if pd.isna(v) else numero_br(v * 100, 1) + "%"
            )
        else:
            formatacao[destino] = (
                lambda v: "" if pd.isna(v) else numero_br(v, 2) + "%"
            )

    elif eh_moeda_pelo_nome(col):

        formatacao[destino] = (
            lambda v: "" if pd.isna(v) else "R$ " + numero_br(v, 2)
        )

    else:

        valores = tabela[col].dropna()
        inteiro = len(valores) > 0 and (valores % 1 == 0).all()
        casas = 0 if inteiro else 2

        formatacao[destino] = (
            lambda v, casas=casas: "" if pd.isna(v) else numero_br(v, casas)
        )


def cor_semaforo(pct):
    if pd.isna(pct):
        return ""
    if pct < limite_vermelho:
        fundo = VERMELHO
    elif pct < limite_verde:
        fundo = AMARELO
    else:
        fundo = VERDE
    return f"background-color:{fundo};color:#000;font-weight:bold;"


def semaforo_maior_melhor(serie, escala):
    pct = pd.to_numeric(serie, errors="coerce")
    if escala == "fracao":
        pct = pct * 100
    return [cor_semaforo(v) for v in pct]


def semaforo_menor_melhor(serie, meta):
    realizado = pd.to_numeric(serie, errors="coerce")
    meta = pd.to_numeric(meta, errors="coerce")
    estilos = []
    for r, m in zip(realizado, meta):
        if pd.isna(r) or pd.isna(m):
            estilos.append("")
        elif r <= m:
            estilos.append(f"background-color:{VERDE};color:#000;font-weight:bold;")
        else:
            estilos.append(f"background-color:{VERMELHO};color:#000;font-weight:bold;")
    return estilos


styler = (
    tabela_final.style
    .format(formatacao)
    .hide(axis="index")
)

for col in colunas_percentuais:

    nome = sem_acento(col)
    destino = novos_nomes[col]

    # A própria meta não leva cor
    if "META" in nome.split():
        continue

    if menor_eh_melhor(col):

        coluna_meta = achar_coluna_meta(col, df.columns)

        if coluna_meta is not None:
            styler = styler.apply(
                lambda s, meta=df_filtrado[coluna_meta]: semaforo_menor_melhor(s, meta),
                subset=[destino]
            )

    else:

        styler = styler.apply(
            lambda s, escala=escalas[col]: semaforo_maior_melhor(s, escala),
            subset=[destino]
        )

# Texto à esquerda, números centralizados
colunas_texto = [
    novos_nomes[c] for c in colunas
    if not pd.api.types.is_numeric_dtype(tabela[c])
]

if colunas_texto:
    styler = styler.set_properties(
        subset=colunas_texto,
        **{"text-align": "left"}
    )

# ==================================================
# TABELA PARA PRINT
# ==================================================

st.markdown("---")

c_titulo, c_ocultar = st.columns([3, 1])

titulo = c_titulo.text_input(
    "Título acima da tabela",
    value="" if supervisor == "Todos" else supervisor
)

# Com um supervisor só, o nome dele já vai no título
coluna_supervisor = novos_nomes.get("Supervisor")
ocultar_supervisor = False

if supervisor != "Todos" and coluna_supervisor:
    c_ocultar.write("")
    ocultar_supervisor = c_ocultar.checkbox(
        "Ocultar coluna Supervisor",
        value=True
    )

if ocultar_supervisor:
    styler = styler.hide(subset=[coluna_supervisor], axis="columns")

cor_cabecalho_texto = cor_texto_para(cor_principal)

# width:1px faz cada coluna ficar com a largura mínima do conteúdo:
# os cabeçalhos quebram linha e os valores ficam inteiros,
# então a tabela já sai compacta para o print.
css = f"""
<style>
.rolagem {{
    max-width: 100%;
    overflow-x: auto;
}}
.area-print {{
    background: #FFFFFF;
    padding: 16px;
    display: inline-block;
    border-radius: 6px;
}}
.area-print h3 {{
    color: #222;
    margin: 0 0 10px 0;
    font-family: Calibri, Arial, sans-serif;
}}
.area-print table {{
    border-collapse: collapse;
    width: 1px;
    font-family: Calibri, Arial, sans-serif;
    font-size: {tamanho_fonte}px;
    color: #222;
}}
.area-print th {{
    background: {cor_principal};
    color: {cor_cabecalho_texto};
    padding: 6px 8px;
    text-align: center;
    vertical-align: middle;
    font-weight: bold;
    border: 1px solid #FFFFFF55;
    min-width: 60px;
}}
.area-print td {{
    padding: 4px 10px;
    text-align: center;
    white-space: nowrap;
    border-bottom: 1px solid #E3E3E3;
}}
.area-print tbody tr:nth-child(even) td {{
    background-color: #F7F9FB;
}}
</style>
"""

cabecalho = f"<h3>{titulo}</h3>" if titulo else ""

st.html(
    css
    + '<div class="rolagem"><div class="area-print">'
    + cabecalho
    + styler.to_html()
    + "</div></div>"
)

st.caption(
    f"{len(tabela_final)} linha(s). "
    "Dica: Win + Shift + S para recortar só a área branca da tabela. "
    "Se a tabela não couber na tela, diminua a fonte na barra lateral "
    "ou tire colunas em 'Colunas da tabela'."
)

# ==================================================
# DOWNLOAD
# ==================================================

csv = tabela_final.to_csv(
    index=False,
    sep=";",
    decimal=","
).encode("utf-8-sig")

st.download_button(
    "📥 Baixar Resultado (CSV)",
    csv,
    "resultado.csv",
    "text/csv"
)

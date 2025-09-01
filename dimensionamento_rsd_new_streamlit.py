import streamlit as st
from PIL import Image
import os
import math

# Caminho para a pasta assets
assets_path = os.path.join(os.path.dirname(__file__), "assets")

# Caminhos para logos e datasheet
logo_dsolar_path = os.path.join(assets_path, "logo_dsolar.png")
logo_advansol_path = os.path.join(assets_path, "logo_advansol.png")
pdf_path = os.path.join(assets_path, "datasheet20250812.pdf")

# Abrir imagens
logo_dsolar = Image.open(logo_dsolar_path)
logo_advansol = Image.open(logo_advansol_path)

# Streamlit layout
st.set_page_config(page_title="Dimensionamento RSD", layout="wide")

# Topo com logo DSolar
st.image(logo_dsolar, width=200)

st.title("Dimensionamento de RSD e Controladora de Dados")

# Inputs do usuário
st.sidebar.header("Inputs do Sistema")
pot_modulo = st.sidebar.number_input("Potência do módulo (W)", value=550)
pot_inversor = st.sidebar.number_input("Potência do inversor (kW)", value=30)
qtd_strings = st.sidebar.number_input("Quantidade de strings conectadas", value=6)
qtd_modulos_string = st.sidebar.number_input("Quantidade de módulos por string", value=10)

rsd_modelo = st.sidebar.selectbox(
    "Modelo de RSD",
    ("APT-MC-R-T2", "APT-MC-MR-T2", "APT-MC-MRO")
)

# Cálculo da quantidade de RSD (1 RSD para cada 2 módulos, arredondando para cima)
qtd_rsd = math.ceil(qtd_modulos_string / 2) * qtd_strings

# Seleção da Controladora de Dados baseada na quantidade de strings
if rsd_modelo == "APT-MC-R-T2":
    # Série R
    dcon_modelo = "APT-CB-DS"
    qtd_dcon = math.ceil(qtd_strings / 12)
else:
    # Série MR ou MRO
    if qtd_strings <= 4:
        dcon_modelo = "APT-CB-D-WIFI-S"
        qtd_dcon = 1
    elif qtd_strings <= 12:
        dcon_modelo = "APT-CB-D-WIFI-M"
        qtd_dcon = 1
    else:
        dcon_modelo = "APT-CB-D-WIFI-L"
        qtd_dcon = 1

# Layout de saída
st.header("Resultados do Dimensionamento")

col1, col2 = st.columns([2, 1])

with col1:
    st.subheader("Quantidade de Dispositivos")
    st.write(f"**Modelo RSD:** {rsd_modelo}")
    st.write(f"**Quantidade de RSD:** {qtd_rsd}")
    st.write(f"**Controladora de Dados:** {dcon_modelo}")
    st.write(f"**Quantidade de DCON:** {qtd_dcon}")

with col2:
    st.image(logo_advansol, width=150)
    st.download_button(
        label="Baixar datasheet",
        data=open(pdf_path, "rb").read(),
        file_name="datasheet20250812.pdf",
        mime="application/pdf"
    )

st.markdown("---")
st.info("Notas: 1 RSD para cada 2 módulos da string. Para número ímpar de módulos, arredonda para cima. Escolha da DCON baseada na quantidade de strings.")

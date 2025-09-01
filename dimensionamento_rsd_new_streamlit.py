import streamlit as st
from PIL import Image
import math
import os
import pandas as pd
from io import BytesIO

# ---------------------------
# Logos
# ---------------------------
logo_dsolar_path = "assets/logo_dsolar.png"
logo_advansol_path = "assets/logo_advansol.png"

logo_dsolar = Image.open(logo_dsolar_path)
logo_advansol = Image.open(logo_advansol_path)

# Exibir logo DSolar no topo
st.image(logo_dsolar, width=200)

# ---------------------------
# Título do App
# ---------------------------
st.title("Dimensionamento de RSD e Controladora de Dados")

# ---------------------------
# Datasheet + Logo Advansol
# ---------------------------
pdf_path = "assets/【PT】V3.1 datasheet20250812.pdf"
col1, col2 = st.columns([1,1])
with col1:
    if os.path.exists(pdf_path):
        with open(pdf_path, "rb") as f:
            st.download_button(
                label="📄 Baixar Datasheet Advansol",
                data=f,
                file_name="Datasheet_Advansol.pdf",
                mime="application/pdf"
            )
    else:
        st.warning("Arquivo de datasheet não encontrado no caminho especificado.")
with col2:
    st.image(logo_advansol, width=200)

# ---------------------------
# Inputs do Sistema
# ---------------------------
st.header("Entradas do Sistema")
potencia_modulo = st.number_input("Potência do módulo (W)", min_value=1, value=600)
potencia_inversor = st.number_input("Potência do inversor (kW)", min_value=1, value=75)
qtd_strings = st.number_input("Quantidade de strings conectadas", min_value=1, value=16)
modulos_por_string = st.number_input("Quantidade de módulos por string", min_value=1, value=14)
modelo_rsd = st.selectbox("Modelo de RSD", ["APT-MC-R-T2", "APT-MC-MR-T2", "APT-MC-MRO"])

# ---------------------------
# Cálculos de RSD
# ---------------------------
rsd_por_string = math.ceil(modulos_por_string / 2)
qtd_rsd_total = rsd_por_string * qtd_strings

# ---------------------------
# Cálculo da Controladora (DCON)
# ---------------------------
if modelo_rsd == "APT-MC-R-T2":
    modelo_dcon = "APT-CB-DS"
    capacidade_dcon = 12
    qtd_dcon = math.ceil(qtd_strings / capacidade_dcon)
    nota = f"A controladora {modelo_dcon} suporta até {capacidade_dcon} strings, portanto foram necessárias {qtd_dcon} unidades."
elif modelo_rsd in ["APT-MC-MR-T2", "APT-MC-MRO"]:
    if qtd_strings <= 4:
        modelo_dcon = "APT-CB-D-WIFI-S"
        capacidade_dcon = 4
        qtd_dcon = 1
        nota = f"A controladora {modelo_dcon} suporta até {capacidade_dcon} strings, portanto foi necessária 1 unidade."
    elif 5 <= qtd_strings <= 12:
        modelo_dcon = "APT-CB-D-WIFI-M"
        capacidade_dcon = 12
        qtd_dcon = 1
        nota = f"A controladora {modelo_dcon} suporta até {capacidade_dcon} strings, portanto foi necessária 1 unidade."
    elif 13 <= qtd_strings <= 20:
        modelo_dcon = "APT-CB-D-WIFI-L"
        capacidade_dcon = 20
        qtd_dcon = 1
        nota = f"A controladora {modelo_dcon} suporta até {capacidade_dcon} strings, portanto foi necessária 1 unidade."
    else:
        qtd_l = qtd_strings // 20
        resto = qtd_strings % 20
        if resto <= 4 and resto > 0:
            qtd_s = 1
            qtd_m = 0
        elif 5 <= resto <= 12:
            qtd_m = 1
            qtd_s = 0
        elif 13 <= resto <= 20:
            qtd_l += 1
            qtd_m = 0
            qtd_s = 0
        else:
            qtd_m = 0
            qtd_s = 0
        partes = []
        if qtd_l > 0:
            partes.append(f"{qtd_l}x APT-CB-D-WIFI-L")
        if qtd_m > 0:
            partes.append(f"{qtd_m}x APT-CB-D-WIFI-M")
        if qtd_s > 0:
            partes.append(f"{qtd_s}x APT-CB-D-WIFI-S")
        modelo_dcon = " + ".join(partes)
        qtd_dcon = qtd_l + qtd_m + qtd_s
        nota = f"Foi necessária a combinação: {modelo_dcon} para atender às {qtd_strings} strings."

# ---------------------------
# Saída em Streamlit
# ---------------------------
st.header("Resultado do Dimensionamento")
st.write(f"**Modelo RSD:** {modelo_rsd}")
st.write(f"**Quantidade de RSD:** {qtd_rsd_total}")
st.write(f"**Controladora de Dados:** {modelo_dcon}")
st.write(f"**Quantidade de DCON:** {qtd_dcon}")
st.write(f"**Nota explicativa:** {nota}")

# ---------------------------
# Exportar resultado para Excel
# ---------------------------
st.header("Exportar Resultado para Excel")
resultado = pd.DataFrame({
    "Modelo RSD": [modelo_rsd],
    "Quantidade RSD": [qtd_rsd_total],
    "Controladora": [modelo_dcon],
    "Quantidade DCON": [qtd_dcon],
    "Nota": [nota]
})

towrite = BytesIO()
resultado.to_excel(towrite, index=False, engine='openpyxl')
towrite.seek(0)
st.download_button(
    label="📥 Baixar Resultado em Excel",
    data=towrite,
    file_name="Dimensionamento_RSD.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

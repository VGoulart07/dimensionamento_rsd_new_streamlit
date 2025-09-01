import streamlit as st
import math

st.title("Dimensionamento de RSD e Controladora de Dados")

# --- Inputs ---
potencia_modulo = st.number_input("Potência do módulo (W)", min_value=1, value=600)
potencia_inversor = st.number_input("Potência do inversor (kW)", min_value=1, value=75)
qtd_strings = st.number_input("Quantidade de strings conectadas", min_value=1, value=16)
modulos_por_string = st.number_input("Quantidade de módulos por string", min_value=1, value=14)
modelo_rsd = st.selectbox("Modelo de RSD", ["APT-MC-R-T2", "APT-MC-MR-T2", "APT-MC-MRO"])

# --- Cálculos ---
# Quantidade de RSD: 1 RSD para cada 2 módulos (arredondando para cima)
rsd_por_string = math.ceil(modulos_por_string / 2)
qtd_rsd_total = rsd_por_string * qtd_strings

# --- Escolha da controladora ---
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
        # Para mais de 20 strings, combinar controladoras
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

# --- Saída ---
st.subheader("Resultado do Dimensionamento")
st.write(f"**Modelo RSD:** {modelo_rsd}")
st.write(f"**Quantidade de RSD:** {qtd_rsd_total}")
st.write(f"**Controladora de Dados:** {modelo_dcon}")
st.write(f"**Quantidade de DCON:** {qtd_dcon}")
st.write(f"**Nota explicativa:** {nota}")

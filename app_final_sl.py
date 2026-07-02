import streamlit as st
import pandas as pd
import conversor

st.set_page_config(page_title="Conversor DI Effa", layout="centered")

# Injetando o CSS para dar aquele toque de cor da Effa (vermelho)
st.markdown("""
    <style>
        .stButton>button { background-color: #ff2600; color: white; width: 100%; font-weight: bold; }
        .stButton>button:hover { background-color: #cc1e00; color: white; }
    </style>
""", unsafe_allow_html=True)

st.title("Conversor DI/NF-e - EFFA MOTORS")

# O st.form garante que o app só processe quando você clicar no botão
with st.form("form_conversao"):
    xml_file = st.file_uploader("Upload do Arquivo (.xml) (DI):", type="xml")
    
    col1, col2 = st.columns(2)
    with col1:
        num_di = st.text_input("Número da DI:", placeholder="Ex: 26/0000000-0")
    with col2:
        num_bl = st.text_input("Número do Processo (BL):", placeholder="Ex: BL123456789")
    
    peso_liquido = st.text_input("Peso Líquido DI:")
    
    st.write("Valores de Rateio (Custos DI):")
    c1, c2 = st.columns(2)
    with c1:
        frete = st.text_input("Frete")
        armazem = st.text_input("Armazém")
        taxa_siscomex = st.text_input("Taxa Siscomex")
    with c2:
        seguro = st.text_input("Seguro")
        despachante = st.text_input("Despachante")
        
    submit = st.form_submit_button("Gerar Dados")

if submit:
    if xml_file is not None:
        # Monta o dicionário com os valores do formulário
        # Nota: O nome das chaves deve bater com o que o seu 'aplicar_rateio' espera
        custos_usuario = {
            "armazenagem": armazem,
            "honorarios_despachante": despachante,
            "liberacao_bl": 0, # Se não tiver no form, coloque 0
            "afrmm": 0,
            "frete_nacional": 0,
            "taxa_siscomex": taxa_siscomex,
            "processo": num_bl
        }
        
        # Chama a função adaptada
        df_final = conversor.processar_di_via_web(xml_file, custos_usuario)
        
        st.success("Conversão concluída!")
        
        # Converte o df_final para Excel na memória para o botão de download
        from io import BytesIO
        buffer = BytesIO()
        df_final.to_excel(buffer, index=False)
        
        st.download_button(
            label="Baixar Planilha de Conferência",
            data=buffer.getvalue(),
            file_name="conferencia_rateio.xlsx"
        )
    else:
        st.error("Por favor, faça o upload do XML.")
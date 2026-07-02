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
        # ... (seu código que chama o conversor.processar_di_via_web permanece igual)
        custos_usuario = {
            "armazenagem": float(armazem) if armazem else 0.0,
            "honorarios_despachante": float(despachante) if despachante else 0.0,
            "taxa_siscomex": float(taxa_siscomex) if taxa_siscomex else 0.0,
            "processo": num_bl if num_bl else "SEM_PROCESSO"
        }
        df_final = conversor.processar_di_via_web(xml_file, custos_usuario)
        st.success("Conversão concluída!")
        
        # Chama a função e recebe o conteúdo (bytes)
        conteudo_xml = conversor.gerar_xml_adempiere_teste(xml_file.name, df_final, custos_usuario)
        
        # Botão de Download para o XML
        st.download_button(
            label="📥 Baixar XML para Adempiere",
            data=conteudo_xml,
            file_name=f"{xml_file.name.replace('.xml', '')}_ADMPIERE.xml",
            mime="application/xml"
        )
        
        
        # Opcional: Manter o download da conferência
        st.download_button(
            label="📊 Baixar Planilha de Conferência",
            data=df_final.to_csv(index=False).encode('utf-8'),
            file_name="conferencia.csv",
            mime="text/csv"
        )
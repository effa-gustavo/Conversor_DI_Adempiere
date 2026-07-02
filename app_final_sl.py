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
        custos_usuario = { ... }
        df_final = conversor.processar_di_via_web(xml_file, custos_usuario)
        
        st.success("Conversão concluída!")
        
        # --- AQUI ESTÁ A MUDANÇA ---
        # 1. Gerar o XML em memória (não salvar no disco)
        # Vamos usar a função que você já tem, mas garantindo que ela retorne o XML
        
        # Dica: No seu conversor.py, certifique-se de que a função 
        # gera_xml_adempiere_teste retorne o objeto XML ou o texto dele.
        
        # Se você quiser simplificar agora, podemos gerar o download do XML:
        # Como o seu conversor.py salva no disco, você pode ler esse arquivo gerado:
        
        caminho_xml_gerado = conversor.PASTA_SAIDA_XML / f"{xml_file.name.replace('.xml', '')}_adempiere_teste.xml"
        
        if caminho_xml_gerado.exists():
            with open(caminho_xml_gerado, "rb") as f:
                xml_data = f.read()
            
            st.download_button(
                label="📥 Baixar XML para Adempiere",
                data=xml_data,
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
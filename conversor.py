from pathlib import Path
from datetime import datetime
import pandas as pd
from lxml import etree

# =========================
# CONFIGURAÇÃO DE CAMINHOS (PORTÁTIL)
# =========================
BASE_DIR = Path(__file__).parent
ARQUIVO_CUSTOS = BASE_DIR / "custos_di.xlsx"

# Substituímos as pastas fixas por funções que não dependem do sistema de arquivos
# ou simplesmente comentamos as que não são usadas no fluxo web.
def gerar_log_web(nome_di, df_itens, custos):
    # Log simplificado apenas para o console do Streamlit
    print(f"Log gerado para DI {nome_di}")

def salvar_base_pecas(caminho, df):
    # Garante que salva no caminho correto
    df.to_excel(caminho, index=False)

# =========================
# FUNÇÕES BÁSICAS
# =========================
def texto(elemento, tag):
    achado = elemento.find(tag)
    if achado is None or achado.text is None:
        return ""
    return achado.text.strip()

def numero_xml(valor):
    if not valor or str(valor).strip() == "":
        return 0.0
    try:
        return float(valor)
    except ValueError:
        return 0.0

def numero_xml_casas_decimais(valor_str, casas_decimais):
    if not valor_str or str(valor_str).strip() == "":
        return 0.0
    try:
        return float(valor_str) / (10 ** casas_decimais)
    except ValueError:
        return 0.0

def dinheiro(valor):
    return f"{float(valor):.2f}"

# =========================
# BASE DE PEÇAS
# =========================
def carregar_base_pecas(caminho_excel):
    try:
        df_base = pd.read_excel(caminho_excel, sheet_name='BASE_PECAS')
    except ValueError:
        df_base = pd.DataFrame(columns=['CODIGO', 'DESCRICAO', 'NCM'])

    dicionario_pecas = {}
    for _, row in df_base.iterrows():
        chave = (str(row['DESCRICAO']).strip().upper(), str(row['NCM']).strip())
        dicionario_pecas[chave] = row['CODIGO']

    return df_base, dicionario_pecas

def salvar_base_pecas(caminho_excel, df_base):
    try:
        with pd.ExcelWriter(caminho_excel, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_base.to_excel(writer, sheet_name='BASE_PECAS', index=False)
    except PermissionError:
        print("\n" + "="*60)
        print("❌ ERRO: O EXCEL ESTÁ ABERTO!")
        print(f"Não foi possível salvar os novos códigos na planilha porque o arquivo está aberto.")
        print("Feche o arquivo no Excel e rode o script novamente para salvar a base.")
        print("="*60 + "\n")

# =========================
# LER XML DA DI
# =========================
def ler_di(caminho_xml, df_base, dicionario_pecas):
    try:
        # 1. Leitura do XML (Flexível para arquivo ou upload)
        if hasattr(caminho_xml, 'read'):
            caminho_xml.seek(0)
            tree = etree.parse(caminho_xml)
        else:
            tree = etree.parse(str(caminho_xml))
            
        root = tree.getroot()
        
        # 2. Processamento dos Itens (Agora indentado dentro da função)
        itens = []
        str_siscomex = texto(root, ".//taxaSiscomex") or texto(root, ".//taxaSiscomexDevida") or texto(root, ".//valorTaxaSiscomex")
        taxa_siscomex_total = numero_xml_casas_decimais(str_siscomex, 2)

        for adicao in root.findall(".//adicao"):
            numero_di = texto(adicao, "numeroDI")
            adicao_num = texto(adicao, "numeroAdicao")
            ncm = texto(adicao, "dadosMercadoriaCodigoNcm")
            peso_adicao = numero_xml_casas_decimais(texto(adicao, "dadosMercadoriaPesoLiquido"), 5)
            cif_adicao = numero_xml_casas_decimais(texto(adicao, "iiBaseCalculo"), 2)
            frete_internacional = numero_xml_casas_decimais(texto(adicao, "valorReaisFreteInternacional"), 2)

            mercadorias = adicao.findall("mercadoria")
            lista_mercadorias_temp = []
            valor_total_mercadorias = 0.0

            for i, mercadoria in enumerate(mercadorias, start=1):
                qtd_corrigida = numero_xml_casas_decimais(texto(mercadoria, "quantidade"), 5)
                v_un_corrigido = numero_xml_casas_decimais(texto(mercadoria, "valorUnitario"), 7)
                valor_item = qtd_corrigida * v_un_corrigido
                valor_total_mercadorias += valor_item
                lista_mercadorias_temp.append({
                    "sequencia": i, "xml_node": mercadoria, 
                    "qtd_corrigida": qtd_corrigida, "v_un_corrigido": v_un_corrigido, "valor_item": valor_item
                })

            for item in lista_mercadorias_temp:
                merc = item["xml_node"]
                fator_rateio = item["valor_item"] / valor_total_mercadorias if valor_total_mercadorias > 0 else 1.0 / max(len(lista_mercadorias_temp), 1)
                
                descricao_xml = texto(merc, "descricaoMercadoria").replace("\n", " ").replace("\r", " ").strip()
                ncm_xml = ncm.strip()
                chave_busca = (descricao_xml.upper(), ncm_xml)

                if chave_busca in dicionario_pecas:
                    codigo_produto = dicionario_pecas[chave_busca]
                else:
                    codigo_produto = df_base['CODIGO'].max() + 1 if not df_base.empty else 25001
                    dicionario_pecas[chave_busca] = codigo_produto
                    nova_linha = pd.DataFrame([{'CODIGO': codigo_produto, 'DESCRICAO': descricao_xml.upper(), 'NCM': ncm_xml}])
                    df_base = pd.concat([df_base, nova_linha], ignore_index=True)

                itens.append({
                    "numero_di": numero_di, "adicao": adicao_num, "item_sequencia": item["sequencia"],
                    "codigo_produto": codigo_produto, "ncm": ncm_xml, "descricao": descricao_xml,
                    "quantidade": item["qtd_corrigida"], "unidade": texto(merc, "unidadeMedida").strip() or "UN",
                    "valor_unitario": item["v_un_corrigido"], "peso": peso_adicao * fator_rateio,
                    "cif": cif_adicao * fator_rateio, "frete_internacional": frete_internacional * fator_rateio,
                    "taxa_siscomex": taxa_siscomex_total * (cif_adicao * fator_rateio / cif_adicao) if cif_adicao > 0 else (taxa_siscomex_total / len(mercadorias)),
                    "ii": 0.0, "ipi": 0.0, "pis": 0.0, "cofins": 0.0
                })

        return pd.DataFrame(itens), df_base, dicionario_pecas

    except Exception as e:
        print(f"Erro crítico no processamento do XML: {e}")
        return pd.DataFrame(), df_base, dicionario_pecas

# =========================
# LER PLANILHA DE CUSTOS E APLICAR RATEIO
# =========================
def ler_custos(numero_di):
    if not ARQUIVO_CUSTOS.exists():
        raise Exception("Arquivo custos_di.xlsx não encontrado na pasta principal.")
    df = pd.read_excel(ARQUIVO_CUSTOS, sheet_name="DESPESAS_DI")
    df["numero_di"] = df["numero_di"].astype(str)
    linha = df[df["numero_di"] == str(numero_di)]
    if linha.empty:
        raise Exception(f"DI {numero_di} não encontrada na planilha DESPESAS_DI.")
    return linha.iloc[0].to_dict()

def aplicar_rateio(df_itens, custos):

    if ler_custos is None:
        custos = {}
    elif not isinstance(custos, dict):
        # Se for qualquer outra coisa (tipo uma Serie ou objeto Streamlit), converte para dict
        try:
            custos = dict(custos)
        except:
            custos = {}

    peso_total = df_itens["peso"].sum()
    if peso_total <= 0:
        raise Exception("Peso total do XML está zerado.")

    # Adicionamos "taxa_siscomex" na lista para ser lida da planilha e rateada pelo peso
    despesas_por_peso = [
        "armazenagem", "honorarios_despachante", "liberacao_bl",
        "afrmm", "frete_internacional", "frete_nacional", "taxa_siscomex"
    ]

    for despesa in despesas_por_peso:
        if despesa == "frete_internacional":
            valor_total = df_itens["frete_internacional"].sum()
        else:
            # Busca o valor da despesa na planilha Excel
            valor_total = float(custos.get(despesa, 0) or 0)
            
        df_itens[f"{despesa}_rateado"] = (df_itens["peso"] / peso_total * valor_total)

    # Agora a Taxa Siscomex rateada entra corretamente na composição das despesas extras
    df_itens["custo_extra_rateado"] = (
        df_itens["armazenagem_rateado"]
        + df_itens["honorarios_despachante_rateado"]
        + df_itens["liberacao_bl_rateado"]
        + df_itens["afrmm_rateado"]
        + df_itens["frete_nacional_rateado"]
        + df_itens["frete_internacional_rateado"]
        + df_itens["taxa_siscomex_rateado"] # Alterado para puxar a versão rateada
    )

    df_itens["custo_total_estimado"] = (
        df_itens["cif"]
        + df_itens["ii"]
        + df_itens["ipi"]
        + df_itens["pis"]
        + df_itens["cofins"]
        + df_itens["custo_extra_rateado"]
    )
    return df_itens

# =========================
# GERAR LOG E PLANILHA DE CONFERÊNCIA
# =========================
def gerar_log(nome_arquivo, df_itens, custos):

    numero_di = df_itens["numero_di"].iloc[0]
    caminho_log = PASTA_LOGS / f"{nome_arquivo}_rateio_log.txt"

    linhas = [
        "LOG DE RATEIO DA DI",
        "=" * 50,
        f"Data/hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        f"DI: {numero_di}",
        f"Processo/Lote: {custos.get('processo', '')}",
        "",
        f"Quantidade de itens: {len(df_itens)}",
        f"Peso total XML: {df_itens['peso'].sum():,.2f}",
        f"CIF total XML: R$ {df_itens['cif'].sum():,.2f}",
        f"Taxa Siscomex Planilha: R$ {df_itens['taxa_siscomex_rateado'].sum():,.2f}",
        "",
        "Status: RATEIO GERADO COM SUCESSO."
    ]
    caminho_log.write_text("\n".join(linhas), encoding="utf-8")

def gerar_planilha_conferencia(nome_arquivo, df_itens):
    caminho_excel = PASTA_CONFERENCIA / f"{nome_arquivo}_conferencia_rateio.xlsx"
    df_itens.to_excel(caminho_excel, index=False)

# =========================
# GERAR XML TESTE ADEMPIERE
# =========================
def gerar_xml_adempiere_teste(nome_arquivo, df_itens, custos):

    custos = dict(custos) if isinstance(custos, dict) else {}

    numero_di = str(df_itens["numero_di"].iloc[0])
    processo = str(custos.get("processo", ""))

    ns = "http://www.portalfiscal.inf.br/nfe"
    NSMAP = {None: ns}

    nfeProc = etree.Element("nfeProc", nsmap=NSMAP, versao="4.00")
    NFe = etree.SubElement(nfeProc, "NFe")
    infNFe = etree.SubElement(NFe, "infNFe", Id=f"NFeTESTE{numero_di}", versao="4.00")

    # Ide
    ide = etree.SubElement(infNFe, "ide")
    etree.SubElement(ide, "cUF").text = "13"
    etree.SubElement(ide, "cNF").text = numero_di[-8:]
    etree.SubElement(ide, "natOp").text = "IMPORTACAO"
    etree.SubElement(ide, "mod").text = "55"
    etree.SubElement(ide, "serie").text = "7"
    etree.SubElement(ide, "nNF").text = numero_di[-6:]
    etree.SubElement(ide, "dhEmi").text = datetime.now().strftime("%Y-%m-%dT%H:%M:%S-04:00")
    etree.SubElement(ide, "dhSaiEnt").text = datetime.now().strftime("%Y-%m-%dT%H:%M:%S-04:00")
    etree.SubElement(ide, "tpNF").text = "0"
    etree.SubElement(ide, "idDest").text = "3"
    etree.SubElement(ide, "cMunFG").text = "1302603"
    etree.SubElement(ide, "tpImp").text = "2"
    etree.SubElement(ide, "tpEmis").text = "1"
    etree.SubElement(ide, "cDV").text = "0"
    etree.SubElement(ide, "tpAmb").text = "1"
    etree.SubElement(ide, "finNFe").text = "1"
    etree.SubElement(ide, "indFinal").text = "0"
    etree.SubElement(ide, "indPres").text = "3"
    etree.SubElement(ide, "indIntermed").text = "0"
    etree.SubElement(ide, "procEmi").text = "0"
    etree.SubElement(ide, "verProc").text = "4.00"

    # Emitente
    emit = etree.SubElement(infNFe, "emit")
    etree.SubElement(emit, "CNPJ").text = "06194010000191"
    etree.SubElement(emit, "xNome").text = "EVER ELETRIC APPLIANCES INDUSTRIA E COMERCIO DE VEICULOS LTD"
    enderEmit = etree.SubElement(emit, "enderEmit")
    etree.SubElement(enderEmit, "xLgr").text = "R AZALEIA"
    etree.SubElement(enderEmit, "nro").text = "273"
    etree.SubElement(enderEmit, "xCpl").text = "BL 3"
    etree.SubElement(enderEmit, "xBairro").text = "DISTRITO INDUSTRIAL II"
    etree.SubElement(enderEmit, "cMun").text = "1302603"
    etree.SubElement(enderEmit, "xMun").text = "MANAUS"
    etree.SubElement(enderEmit, "UF").text = "AM"
    etree.SubElement(enderEmit, "CEP").text = "69007800"
    etree.SubElement(enderEmit, "cPais").text = "1058"
    etree.SubElement(enderEmit, "xPais").text = "BRASIL"
    etree.SubElement(emit, "IE").text = "062004140"
    etree.SubElement(emit, "CRT").text = "3"

    # Destinatário
    dest = etree.SubElement(infNFe, "dest")
    etree.SubElement(dest, "idEstrangeiro")
    etree.SubElement(dest, "xNome").text = "FORNECEDOR EXTERIOR"
    enderDest = etree.SubElement(dest, "enderDest")
    etree.SubElement(enderDest, "xLgr").text = "EXTERIOR"
    etree.SubElement(enderDest, "nro").text = "S/N"
    etree.SubElement(enderDest, "xBairro").text = "EXTERIOR"
    etree.SubElement(enderDest, "cMun").text = "9999999"
    etree.SubElement(enderDest, "xMun").text = "EXTERIOR"
    etree.SubElement(enderDest, "UF").text = "EX"
    # ---> AJUSTE: Código do País do Fornecedor para 1600
    etree.SubElement(enderDest, "cPais").text = "1600"
    etree.SubElement(enderDest, "xPais").text = "CHINA"
    etree.SubElement(dest, "indIEDest").text = "9"

    total_produtos = 0
    total_nf = 0

    for index, linha in df_itens.iterrows():
        det = etree.SubElement(infNFe, "det", nItem=str(index + 1))
        prod = etree.SubElement(det, "prod")

        etree.SubElement(prod, "M_Product_ID").text = "4021955"

        valor_produto = round(float(linha["cif"]), 2)
        custo_total = round(float(linha["custo_total_estimado"]), 2)

        total_produtos += valor_produto
        total_nf += custo_total

        codigo_produto = str(linha.get("codigo_produto", f"{linha['adicao']}-{linha['item_sequencia']}"))
        etree.SubElement(prod, "cProd").text = codigo_produto
        etree.SubElement(prod, "cEAN").text = "SEM GTIN"
        etree.SubElement(prod, "xProd").text = str(linha["descricao"])[:120]
        etree.SubElement(prod, "NCM").text = str(linha["ncm"])
        etree.SubElement(prod, "CFOP").text = "3101"
        etree.SubElement(prod, "uCom").text = str(linha["unidade"]).strip() or "UN"
        etree.SubElement(prod, "qCom").text = f"{float(linha['quantidade']):.4f}"
        
        quantidade = float(linha["quantidade"])
        v_unitario_corrigido = (valor_produto / quantidade) if quantidade > 0 else 0.0
        
        v_un_str = f"{v_unitario_corrigido:.10f}".rstrip('0')
        if v_un_str.endswith('.'):
            v_un_str += "00"
        
        etree.SubElement(prod, "vUnCom").text = v_un_str
        etree.SubElement(prod, "vProd").text = dinheiro(valor_produto)
        etree.SubElement(prod, "cEANTrib").text = "SEM GTIN"
        etree.SubElement(prod, "uTrib").text = str(linha["unidade"]).strip() or "UN"
        etree.SubElement(prod, "qTrib").text = f"{float(linha['quantidade']):.4f}"
        etree.SubElement(prod, "vUnTrib").text = v_un_str
        
        # Despesas Acessórias do Item (Fretes + Honorários + Taxa Siscomex, etc)
        etree.SubElement(prod, "vOutro").text = dinheiro(linha["custo_extra_rateado"])
        etree.SubElement(prod, "indTot").text = "1"

        DI = etree.SubElement(prod, "DI")
        etree.SubElement(DI, "nDI").text = str(numero_di)
        etree.SubElement(DI, "dDI").text = datetime.now().strftime("%Y-%m-%d")
        etree.SubElement(DI, "xLocDesemb").text = "MANAUS"
        etree.SubElement(DI, "UFDesemb").text = "AM"
        etree.SubElement(DI, "dDesemb").text = datetime.now().strftime("%Y-%m-%d")
        etree.SubElement(DI, "tpViaTransp").text = "1"
        etree.SubElement(DI, "vAFRMM").text = dinheiro(linha.get("afrmm_rateado", 0))
        etree.SubElement(DI, "tpIntermedio").text = "1"
        etree.SubElement(DI, "cExportador").text = "FORNECEDOR"

        adi = etree.SubElement(DI, "adi")
        etree.SubElement(adi, "nAdicao").text = str(linha["adicao"])
        etree.SubElement(adi, "nSeqAdic").text = str(linha["item_sequencia"])
        etree.SubElement(adi, "cFabricante").text = "FORNECEDOR"

        imposto = etree.SubElement(det, "imposto")
        
        ICMS = etree.SubElement(imposto, "ICMS")
        ICMS40 = etree.SubElement(ICMS, "ICMS40")
        etree.SubElement(ICMS40, "orig").text = "1"
        etree.SubElement(ICMS40, "CST").text = "41"

        # ---> AJUSTE: Impostos Zerados e Tags Corrigidas para CSTs de Não Tributado/Outros
        imposto_ii = etree.SubElement(imposto, "II")
        etree.SubElement(imposto_ii, "vBC").text = "0.00"
        etree.SubElement(imposto_ii, "vDespAdu").text = "0.00"
        etree.SubElement(imposto_ii, "vII").text = "0.00"
        etree.SubElement(imposto_ii, "vIOF").text = "0.00"

        IPI = etree.SubElement(imposto, "IPI")
        etree.SubElement(IPI, "cEnq").text = "999"
        IPINT = etree.SubElement(IPI, "IPINT")
        etree.SubElement(IPINT, "CST").text = "05"

        PIS = etree.SubElement(imposto, "PIS")
        PISAliq = etree.SubElement(PIS, "PISAliq") 
        etree.SubElement(PISAliq, "CST").text = "01"
        etree.SubElement(PISAliq, "vBC").text = "0.00"
        etree.SubElement(PISAliq, "pPIS").text = "0.0000"
        etree.SubElement(PISAliq, "vPIS").text = "0.00"

        COFINS = etree.SubElement(imposto, "COFINS")
        COFINSAliq = etree.SubElement(COFINS, "COFINSAliq")
        etree.SubElement(COFINSAliq, "CST").text = "01"
        etree.SubElement(COFINSAliq, "vBC").text = "0.00"
        etree.SubElement(COFINSAliq, "pCOFINS").text = "0.0000"
        etree.SubElement(COFINSAliq, "vCOFINS").text = "0.00"

    # Totais
    total = etree.SubElement(infNFe, "total")
    ICMSTot = etree.SubElement(total, "ICMSTot")
    etree.SubElement(ICMSTot, "vBC").text = "0.00"
    etree.SubElement(ICMSTot, "vICMS").text = "0.00"
    etree.SubElement(ICMSTot, "vICMSDeson").text = "0.00"
    etree.SubElement(ICMSTot, "vFCPUFDest").text = "0.00"
    etree.SubElement(ICMSTot, "vICMSUFDest").text = "0.00"
    etree.SubElement(ICMSTot, "vICMSUFRemet").text = "0.00"
    etree.SubElement(ICMSTot, "vFCP").text = "0.00"
    etree.SubElement(ICMSTot, "vBCST").text = "0.00"
    etree.SubElement(ICMSTot, "vST").text = "0.00"
    etree.SubElement(ICMSTot, "vFCPST").text = "0.00"
    etree.SubElement(ICMSTot, "vFCPSTRet").text = "0.00"
    etree.SubElement(ICMSTot, "vProd").text = dinheiro(total_produtos)
    etree.SubElement(ICMSTot, "vFrete").text = "0.00"
    etree.SubElement(ICMSTot, "vSeg").text = "0.00"
    etree.SubElement(ICMSTot, "vDesc").text = "0.00"
    etree.SubElement(ICMSTot, "vII").text = "0.00"
    etree.SubElement(ICMSTot, "vIPI").text = "0.00"
    etree.SubElement(ICMSTot, "vIPIDevol").text = "0.00"
    etree.SubElement(ICMSTot, "vPIS").text = "0.00"
    etree.SubElement(ICMSTot, "vCOFINS").text = "0.00"
    
    # As despesas extras rateadas e a Taxa Siscomex são injetadas em vOutro
    etree.SubElement(ICMSTot, "vOutro").text = dinheiro(df_itens["custo_extra_rateado"].sum())
        
    soma_sefaz = total_produtos + df_itens["custo_extra_rateado"].sum()
    etree.SubElement(ICMSTot, "vNF").text = dinheiro(soma_sefaz)

    transp = etree.SubElement(infNFe, "transp")
    etree.SubElement(transp, "modFrete").text = "0"

    infAdic = etree.SubElement(infNFe, "infAdic")
    etree.SubElement(infAdic, "infCpl").text = (
        f"XML teste gerado via Python. DI {numero_di}. Processo {processo}."
    )

    xml_bytes = etree.tostring(
        nfeProc, 
        pretty_print=True, 
        xml_declaration=True, 
        encoding="utf-8"
    )
    return xml_bytes

# =========================
# PROCESSAR ARQUIVOS
# =========================
def main():
    arquivos = list(PASTA_ENTRADA.glob("*.xml"))

    if not arquivos:
        print("Nenhum XML encontrado na pasta entrada_di.")
        return

    print("Carregando Base de Peças do Excel...")
    df_base, dicionario_pecas = carregar_base_pecas(ARQUIVO_CUSTOS)

    for arquivo in arquivos:
        print(f"Lendo arquivo: {arquivo.name}")

        df_itens, df_base, dicionario_pecas = ler_di(arquivo, df_base, dicionario_pecas)

        if df_itens.empty:
            print("Nenhum item encontrado no XML.")
            continue

        numero_di = df_itens["numero_di"].iloc[0]
        custos = ler_custos(numero_di)
        df_itens = aplicar_rateio(df_itens, custos)
        
        gerar_log(arquivo.stem, df_itens, custos)
        gerar_planilha_conferencia(arquivo.stem, df_itens)
        gerar_xml_adempiere_teste(arquivo.stem, df_itens, custos)

    print("Salvando base atualizada de peças no Excel...")
    salvar_base_pecas(ARQUIVO_CUSTOS, df_base)
    print("Processo concluído com sucesso!")

def processar_di_via_web(uploaded_file, dados_formulario):
    # 1. Carrega base de peças (isso pode ficar igual)
    df_base, dicionario_pecas = carregar_base_pecas(ARQUIVO_CUSTOS)
    
    # 2. Processa o XML (Agora passamos o objeto 'uploaded_file' direto para o ler_di)
    # A função ler_di aceita um 'caminho_xml', mas o lxml consegue ler se passarmos o arquivo aberto
    df_itens, df_base, dicionario_pecas = ler_di(uploaded_file, df_base, dicionario_pecas)
    
    if df_itens.empty:
        return pd.DataFrame()
    
    custos = dict(dados_formulario) if isinstance(dados_formulario, dict) else {}

    # 3. Aplica os custos (dados_formulario é o dicionário que você montará no app_final_sl)
    df_itens = aplicar_rateio(df_itens, dados_formulario)
    return df_itens
    
    # 4. Atualiza a base de peças no Excel
    salvar_base_pecas(ARQUIVO_CUSTOS, df_base)
    
    # Retorna o DataFrame para o Streamlit poder criar o download
    return df_itens

if __name__ == "__main__":
    main()
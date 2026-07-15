# -*- coding: utf-8 -*-
from danfe_recognition_module import DanfeRecognizer


TARGET_TEXT = """RECEBEMOS DE NORTE TOOLS COM FERRAMENTAS E MAQUINAS UNIPESSOAL LTDA OS PRODUTOS E/OU SERVIÇOS CONSTANTES DA NOTA NF-e
FISCAL ELETRÔNICA INDICADA ABAIXO. EMISSÃO: 06/01/2026 VALOR TOTAL: R$ 44,89 DESTINATÁRIO: Edson Vagner Betlinski Da Silva - Rua Joao
Francisco Fonseca, 131 Popular Tupancireta-RS
Nº. 000.044.182
DATA DE RECEBIMENTO IDENTIFICAÇÃO E ASSINATURA DO RECEBEDOR
Série 002
IDENTIFICAÇÃO DO EMITENTE
DANFE
Documento Auxiliar da Nota
Fiscal Eletrônica
NORTE TOOLS COM FERRAMENTAS E MAQUINAS UNIPESSOAL LTDA 0 - ENTRADA 1
Av. Presidente Wilson, 5056, 1504 - SALA 07 1 - SAÍDA CHAVE DE ACESSO
Vila Independencia - 04220-001
3526 0139 9698 9000 0180 5500 2000 0441 8213 3365 6441
Sao Paulo - SP Fone/Fax: 1197882237 Nº. 000.044.182
Série 002 Consulta de autenticidade no portal nacional da NF-e
Folha 1/1 www.nfe.fazenda.gov.br/portal ou no site da Sefaz Autorizadora
NATUREZA DA OPERAÇÃO PROTOCOLO DE AUTORIZAÇÃO DE USO
Venda de mercadorias 135260050803407 - 06/01/2026 08:35:42
INSCRIÇÃO ESTADUAL INSCRIÇÃO MUNICIPAL INSCRIÇÃO ESTADUAL DO SUBST. TRIBUT. CNPJ
130188978110 39.969.890/0001-80
DESTINATÁRIO / REMETENTE
NOME / RAZÃO SOCIAL CNPJ / CPF DATA DA EMISSÃO
Edson Vagner Betlinski Da Silva 045.476.120-13 06/01/2026
ENDEREÇO BAIRRO / DISTRITO CEP DATA DA SAÍDA/ENTRADA
Rua Joao Francisco Fonseca, 131 Popular 98170-000 06/01/2026
MUNICÍPIO UF FONE / FAX INSCRIÇÃO ESTADUAL HORA DA SAÍDA/ENTRADA
Tupancireta RS 08:29:59
CÁLCULO DO IMPOSTO
BASE DE CÁLC. DO ICMS VALOR DO ICMS BASE DE CÁLC. ICMS S.T. VALOR DO ICMS SUBST. V. IMP. IMPORTAÇÃO V. ICMS UF REMET. V. FCP UF DEST. VALOR DO PIS V. TOTAL PRODUTOS
44,89 1,80 0,00 0,00 0,00 0,00 0,90 0,74 44,89
VALOR DO FRETE VALOR DO SEGURO DESCONTO OUTRAS DESPESAS VALOR TOTAL IPI V. ICMS UF DEST. V. TOT. TRIB. VALOR DA COFINS V. TOTAL DA NOTA
0,00 0,00 0,00 0,00 0,00 5,84 18,62 3,41 44,89
TRANSPORTADOR / VOLUMES TRANSPORTADOS
NOME / RAZÃO SOCIAL FRETE CÓDIGO ANTT PLACA DO VEÍCULO UF CNPJ / CPF
0-Por conta do Rem
ENDEREÇO MUNICÍPIO UF INSCRIÇÃO ESTADUAL
QUANTIDADE ESPÉCIE MARCA NUMERAÇÃO PESO BRUTO PESO LÍQUIDO
1,000 1,000
DADOS DOS PRODUTOS / SERVIÇOS
CÓDIGO PRODUTO DESCRIÇÃO DO PRODUTO / SERVIÇO NCM/SH O/CST CFOP UN QUANT V U A N L I O T R V TO A T L A O L R V D A E L S O C R B I . C C M ÁL S C V I A C L M O S R VA I L PI OR A IC L M ÍQ S . A I L P Í I Q.
E71T-GS0.8mm-1KGArame de solda revestido MIG uso sem gas - E71T-GS 83119000 100 6108 un 1,0000 44,8900 44,89 0,00 44,89 1,80 4,00
0.8mm 1KG
"""


N108_TEXT = """RECEBEMOS DE NORTE TOOLS COM FERRAMENTAS E MAQUINAS UNIPESSOAL LTDA OS PRODUTOS E/OU SERVIÇOS CONSTANTES DA NOTA NF-e
FISCAL ELETRÔNICA INDICADA ABAIXO. EMISSÃO: 02/01/2026 VALOR TOTAL: R$ 131,00 DESTINATÁRIO: Jose Edmilson de Sousa Chagas - rua h N
108, 108 Area Rural de Petrolina Petrolina-PE
Nº. 000.043.992
DATA DE RECEBIMENTO IDENTIFICAÇÃO E ASSINATURA DO RECEBEDOR
Série 002
IDENTIFICAÇÃO DO EMITENTE
DANFE
Documento Auxiliar da Nota
Fiscal Eletrônica
NORTE TOOLS COM FERRAMENTAS E MAQUINAS UNIPESSOAL LTDA 0 - ENTRADA 1
Av. Presidente Wilson, 5056, 1504 - SALA 07 1 - SAÍDA CHAVE DE ACESSO
Vila Independencia - 04220-001
3526 0139 9698 9000 0180 5500 2000 0439 9213 3320 1792
Sao Paulo - SP Fone/Fax: 1197882237 Nº. 000.043.992
Série 002 Consulta de autenticidade no portal nacional da NF-e
Folha 1/1 www.nfe.fazenda.gov.br/portal ou no site da Sefaz Autorizadora
NATUREZA DA OPERAÇÃO PROTOCOLO DE AUTORIZAÇÃO DE USO
Venda de mercadorias 135260005773840 - 02/01/2026 08:50:36
"""


KIT_ITEM_TEXT = """RECEBEMOS DE NORTE TOOLS COM FERRAMENTAS E MAQUINAS UNIPESSOAL LTDA OS PRODUTOS E/OU SERVIÇOS CONSTANTES DA NOTA NF-e
FISCAL ELETRÔNICA INDICADA ABAIXO. EMISSÃO: 05/01/2026 VALOR TOTAL: R$ 38,90 DESTINATÁRIO: Antonio Izidro Dos Santos Neto - Rua
Walfredo Gomes de Araujo, 082 Centro Boa Vista-PB
Nº. 000.044.079
IDENTIFICAÇÃO DO EMITENTE
NORTE TOOLS COM FERRAMENTAS E MAQUINAS UNIPESSOAL LTDA 0 - ENTRADA 1
Av. Presidente Wilson, 5056, 1504 - SALA 07 1 - SAÍDA CHAVE DE ACESSO
Vila Independencia - 04220-001
3526 0139 9698 9000 0180 5500 2000 0440 7913 3345 4295
NATUREZA DA OPERAÇÃO PROTOCOLO DE AUTORIZAÇÃO DE USO
Venda de mercadorias 135260036017044 - 05/01/2026 09:35:34
DADOS DOS PRODUTOS / SERVIÇOS
CÓDIGO PRODUTO DESCRIÇÃO DO PRODUTO / SERVIÇO NCM/SH O/CST CFOP UN QUANT V U A N L I O T R V TO A T L A O L R V D A E L S O C R B I . C C M ÁL S C V I A C L M O S R VA I L PI OR A IC L M ÍQ S . A I L P Í I Q.
KK002-11 Kit Jogo de Ferramentas 17 pecas com Maleta KK002-11 82060000 100 6108 un 1,0000 38,9000 38,90 0,00 38,90 1,56 4,00
Kraw
"""


MERGED_NUMERIC_TEXT = """RECEBEMOS DE NORTE TOOLS COM FERRAMENTAS E MAQUINAS UNIPESSOAL LTDA OS PRODUTOS E/OU SERVIÇOS CONSTANTES DA NOTA NF-e
FISCAL ELETRÔNICA INDICADA ABAIXO. EMISSÃO: 06/01/2026 VALOR TOTAL: R$ 1.699,89 DESTINATÁRIO: Lucas de Araujo Avila - Avenida
Guapore, 2265 Setor Central Gurupi-TO
Nº. 000.044.169
IDENTIFICAÇÃO DO EMITENTE
NORTE TOOLS COM FERRAMENTAS E MAQUINAS UNIPESSOAL LTDA 0 - ENTRADA 1
Av. Presidente Wilson, 5056, 1504 - SALA 07 1 - SAÍDA CHAVE DE ACESSO
Vila Independencia - 04220-001
3526 0139 9698 9000 0180 5500 2000 0441 6913 3365 6300
NATUREZA DA OPERAÇÃO PROTOCOLO DE AUTORIZAÇÃO DE USO
Venda de mercadorias 135260050782140 - 06/01/2026 08:34:29
DADOS DOS PRODUTOS / SERVIÇOS
CÓDIGO PRODUTO DESCRIÇÃO DO PRODUTO / SERVIÇO NCM/SH O/CST CFOP UN QUANT V U A N L I O T R V TO A T L A O L R V D A E L S O C R B I . C C M ÁL S C V I A C L M O S R VA I L PI OR A IC L M ÍQ S . A I L P Í I Q.
LIFT-255 Maquina De Solda Inversora Mig/mma/tig (lift)-255 85153900 200 6108 un 1,00001.699,8900 1.699,89 0,00 1.699,89 68,00 4,00
Bivolt Usk
"""


HEADER_ONLY_PROTOCOL_TEXT = """Nº. 000.044.689
NATUREZA DA OPERAÇÃO PROTOCOLO DE AUTORIZAÇÃO DE USO
Venda de mercadorias
"""


BROKEN_ACCESS_KEY_LAYOUT_TEXT = """RECEBEMOS DE NORTE TOOLS COMERCIO DE FERRAMENTAS E MAQUINAS UNIPESSOAL LT OS PRODUTOS CONSTANTES DA NOTA FISCAL INDICADA AO LADO NF-e
Nº 000.011.153
DATA DE RECEBIMENTO IDENTIFICACAO E ASSINATURA DO RECEBEDOR
SÉRIE003
DANFE
NORTE TOOLS
Documento Auxiliar da
COMERCIO DE
Nota Fiscal Eletrônica
FERRAMENTAS E 0: Entrada
1
1: Saída
Rua Silva Bueno, 1504, SALA 33 OU 35 Refere - Nº 000.011.153 CHAVE DE ACE 3 S 5 S 2 O 6 0239 9698 9000 0180 5500 3000 0111 5312 6643 7236
Ipiranga, Sao Paulo, SP - CEP: 04208001 Fone:
SÉRIE:003
0011978822370 Consulta de autenticidade no portal nacional da NF-e
Folha 1 d 1 www.nfe.fazenda.gov.br/portal ou no site da Sefaz Autorizadora
NATUREZA DA OPERAÇÃO PROTOCOLO DE AUTORIZAÇÃO DE USO
Venda de mercadorias 135260751564 26/02/2026 15:44:05
INSCRIÇÃO ESTADUAL INSC. ESTADUAL DO SUBST. TRIBUTÁRIO CNPJ
130188978110 39.969.890/0001-80
DESTINATÁRIO / REMETENTE
NOME/RAZÃO SOCIAL C.N.P.J / C.P.F. DATA DA EMISSÃO
Anderson Eckhardt 042.588.770-77 26/02/2026
ENDEREÇO BAIRRO/DISTRITO CEP DATA DA ENTRADA / SAÍDA
Rua Coronel Bicaco, 134 - Nao consta Centro 98640000 26/02/2026
MUNICÍPIO FONE/FAX UF INSCRIÇÃO ESTADUAL HORA DE SAÍDA
Crissiumal RS 15:44:03
FATURA/DUPLICATA
CÁLCULO DO IMPOSTO
BASE DE CÁLCULO DO ICMS VALOR DO ICMS BASE DE CÁLCULO DO ICMS SUBSTITUIÇÃO VALOR DO ICMS SUBSTITUIÇÃO VALOR TOTAL DOS PRODUTOS
199,89 8,00 0,00 0,00 199,89
VALOR DO FRETE VALOR DO SEGURO DESCONTO OUTRAS DESPESAS ACESSÓRIAS VALOR DO IPI VALOR TOTAL DA NOTA
0,00 0,00 0,00 0,00 0,00 199,89
TRANSPORTADOR/VOLUME
RAZÃO SOCIAL FRETE POR CONTA CODIGO ANTT PLACA DO VEÍCULO UF CNPJ/CPF
EBAZAR.COM.BR LTDA 2 - Terceiros 03.007.331/0122-39
ENDEREÇO MUNICÍPIO UF INSCRIÇÃO ESTADUAL
AVENIDA DAS NACOES UNIDAS 3000 3003 OSASCO SP 120519234116
QUANTIDADE ESPÉCIE MARCA NUMERAÇÃO PESO BRUTO PESO LÍQUIDO
1 2,730 2,730
INFORMAÇÕES DO LOCAL DE ENTREGA / RETIRADA
NOME/RAZÃO SOCIAL C.N.P.J / C.P.F. INSCRIÇÃO ESTADUAL
ENDEREÇO BAIRRO/DISTRITO CEP
MUNICÍPIO UF FONE/FAX
DADOS DO PRODUTO / SERVIÇOS
CÓDIGO DESCRIÇAO DOS PRODUTOS / SERVIÇOS NCM/SH CST CFOP UNID. QTD. VLR UNIT. VALOR TOTAL B. CALC. VALOR VALOR ALÍQUOTAS
PRODUTO ICMS ICMS IPI ICMS IPI
MMA-120-220V MMA-120-220V 85153190 200 6106 UN 1 199,89 199,89 199,89 8,00 0,00 4,00 0,00
"""

MULTILINE_ITEMS_WITH_TAX_NOTES_TEXT = """DADOS DOS PRODUTOS / SERVIÇOS
CÓDIGO PRODUTO DESCRIÇÃO DO PRODUTO / SERVIÇO NCM/SH O/CST CFOP UN QUANT V U A N L I O T R V TO A T L A O L R V D A E L S O C R B I . C C M ÁL S C V I A C L M O S R VA I L PI OR A IC L M ÍQ S . A I L P Í I Q.
MA-03 Mascara de Solda Automatica Tonalidade DIN11 MA-03 65069900 100 6108 un 1,0000 119,8900 119,89 0,00 119,89 4,80 4,00
Kraw
pFCPUFDest=2,00% pICMSUFDest=20,50%
pICMSInterPart=100,00% vFCPUFDest=2,40
vICMSUFDest=19,78 vICMSUFRemet=0,00 (Pedido
260111P76P8FWD)
E71T-GS0.8mm-1KG Arame de solda revestido MIG uso sem gas - E71T-GS 83119000 200 6108 un 1,0000 34,8900 34,89 0,00 34,89 1,40 4,00
0.8mm 1KG
pFCPUFDest=2,00% pICMSUFDest=20,50%
pICMSInterPart=100,00% vFCPUFDest=0,70
vICMSUFDest=5,76 vICMSUFRemet=0,00 (Pedido
260111P76P8FWD)
KK-1808 Pistola de Pintura Eletrica 650w 800ml KK-1808 Kraw - 84244100 100 6108 un 1,0000 129,8900 129,89 0,00 129,89 5,20 4,00
220v
pFCPUFDest=2,00% pICMSUFDest=20,50%
pICMSInterPart=100,00% vFCPUFDest=2,60
vICMSUFDest=21,43 vICMSUFRemet=0,00 (Pedido
260111P76P8FWD)
DADOS ADICIONAIS
"""


MULTILINE_WIRE_ITEMS_TEXT = """DADOS DOS PRODUTOS / SERVIÇOS
CÓDIGO PRODUTO DESCRIÇÃO DO PRODUTO / SERVIÇO NCM/SH O/CST CFOP UN QUANT V U A N L I O T R V TO A T L A O L R V D A E L S O C R B I . C C M ÁL S C V I A C L M O S R VA I L PI OR A IC L M ÍQ S . A I L P Í I Q.
E71T-GS0.8mm-1KG Arame de solda revestido MIG uso sem gas - E71T-GS 83119000 200 6108 un 1,0000 34,8900 34,89 0,00 34,89 1,40 4,00
0.8mm 1KG
pFCPUFDest=2,00% pICMSUFDest=19,50%
pICMSInterPart=100,00% vFCPUFDest=0,70
vICMSUFDest=5,41 vICMSUFRemet=0,00 (Pedido
260113U5C2NVEK)
E71TGS-1.0mm-1KG Arame de solda revestido MIG uso sem gas - E71T-GS 83119000 200 6108 un 1,0000 34,8900 34,89 0,00 34,89 1,40 4,00
1.0mm 1kg
pFCPUFDest=2,00% pICMSUFDest=19,50%
pICMSInterPart=100,00% vFCPUFDest=0,70
vICMSUFDest=5,41 vICMSUFRemet=0,00 (Pedido
260113U5C2NVEK)
DADOS ADICIONAIS
"""


def test_danfe_recognize_target_invoice_fields():
    recognizer = DanfeRecognizer()
    result = recognizer.recognize_from_text(TARGET_TEXT)
    item = result["items"][0]

    assert result["numero_nota"] == "000.044.182"
    assert result["chave_acesso"] == "35260139969890000180550020000441821333656441"
    assert result["natureza_operacao"] == "Venda de mercadorias"
    assert result["protocolo"] == "135260050803407 - 06/01/2026 08:35:42"
    assert "Av. Presidente Wilson" in result["emitente_endereco"]
    assert "04220-001" in result["emitente_endereco"]
    assert result["destinatario_endereco"] == "Rua Joao Francisco Fonseca, 131 Popular 98170-000 Tupancireta RS"
    assert "QUANTIDADE ESPÉCIE" not in result["destinatario_endereco"]
    assert item["valor_total"] == 44.89
    assert item["bc_icms"] == 44.89
    assert item["v_icms"] == 1.8


def test_danfe_numero_nota_ignores_address_house_number():
    recognizer = DanfeRecognizer()
    result = recognizer.recognize_from_text(N108_TEXT)

    assert result["numero_nota"] == "000.043.992"


def test_danfe_item_parser_does_not_treat_description_kit_as_unit():
    recognizer = DanfeRecognizer()
    result = recognizer.recognize_from_text(KIT_ITEM_TEXT)
    item = result["items"][0]

    assert item["codigo"] == "KK002-11"
    assert item["unidade"] == "un"
    assert item["qtd"] == 1.0
    assert item["valor_unit"] == 38.9
    assert item["valor_total"] == 38.9
    assert item["bc_icms"] == 38.9
    assert item["v_icms"] == 1.56


def test_danfe_item_parser_handles_merged_quantity_and_unit_price():
    recognizer = DanfeRecognizer()
    result = recognizer.recognize_from_text(MERGED_NUMERIC_TEXT)
    item = result["items"][0]

    assert item["codigo"] == "LIFT-255"
    assert item["unidade"] == "un"
    assert item["qtd"] == 1.0
    assert item["valor_unit"] == 1699.89
    assert item["valor_total"] == 1699.89
    assert item["bc_icms"] == 1699.89
    assert item["v_icms"] == 68.0


def test_danfe_protocol_header_without_value_becomes_empty():
    recognizer = DanfeRecognizer()
    result = recognizer.recognize_from_text(HEADER_ONLY_PROTOCOL_TEXT)

    assert result["natureza_operacao"] == "Venda de mercadorias"
    assert result["protocolo"] == ""


def test_danfe_broken_access_key_layout_is_parsed():
    recognizer = DanfeRecognizer()
    result = recognizer.recognize_from_text(BROKEN_ACCESS_KEY_LAYOUT_TEXT)
    item = result["items"][0]

    assert result["numero_nota"] == "000.011.153"
    assert result["chave_acesso"] == "35260239969890000180550030000111531266437236"
    assert result["natureza_operacao"] == "Venda de mercadorias"
    assert result["protocolo"] == "135260751564 26/02/2026 15:44:05"
    assert result["data_emissao"] == "26/02/2026"
    assert result["destinatario_nome"] == "Anderson Eckhardt"
    assert result["destinatario_cnpj_cpf"] == "042.588.770-77"
    assert result["destinatario_endereco"] == "Rua Coronel Bicaco, 134 - Nao consta Centro 98640000 Crissiumal RS"
    assert result["v_prod"] == 199.89
    assert result["v_nota"] == 199.89
    assert item["codigo"] == "MMA-120-220V"
    assert item["valor_total"] == 199.89


def test_danfe_access_key_can_fallback_to_filename():
    recognizer = DanfeRecognizer()
    text = "Nº 000.011.153\nVenda de mercadorias"
    result = recognizer.recognize_from_text(
        text,
        file_path=r"C:/Users/123/Downloads/汇总提取_03241656/5576400897_35260239969890000180550030000111531266437236-DANFE.pdf",
    )

    assert result["chave_acesso"] == "35260239969890000180550030000111531266437236"


def test_danfe_item_parser_keeps_multiline_specs_with_previous_item():
    recognizer = DanfeRecognizer()
    result = recognizer.recognize_from_text(MULTILINE_ITEMS_WITH_TAX_NOTES_TEXT)

    assert [item["codigo"] for item in result["items"]] == ["MA-03", "E71T-GS0.8mm-1KG", "KK-1808"]
    assert "Kraw" in result["items"][0]["descricao"]
    assert "0.8mm 1KG" in result["items"][1]["descricao"]
    assert "220v" in result["items"][2]["descricao"]
    assert all("pICMSInterPart" not in item["descricao"] for item in result["items"])


def test_danfe_item_parser_keeps_second_wire_code_after_multiline_specs():
    recognizer = DanfeRecognizer()
    result = recognizer.recognize_from_text(MULTILINE_WIRE_ITEMS_TEXT)

    assert [item["codigo"] for item in result["items"]] == ["E71T-GS0.8mm-1KG", "E71TGS-1.0mm-1KG"]
    assert "0.8mm 1KG" in result["items"][0]["descricao"]
    assert "1.0mm 1kg" in result["items"][1]["descricao"]
    assert result["items"][1]["valor_total"] == 34.89

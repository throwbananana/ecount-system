# -*- coding: utf-8 -*-
"""NF-e XML parser used by the DANFE recognition module."""

from __future__ import annotations

import copy
import re
import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Optional


class DanfeXmlParser:
    CNPJ_ALNUM_RE = r'[A-Z0-9]{2}\.[A-Z0-9]{3}\.[A-Z0-9]{3}/[A-Z0-9]{4}-\d{2}'

    def __init__(self):
        self.template = {
            'source_used': 'xml',
            'status_documento': '',
            'fonte_status': 'xml',
            'chave_acesso': '',
            'numero_nota': '',
            'serie': '',
            'natureza_operacao': '',
            'protocolo': '',
            'data_emissao': '',
            'data_saida': '',
            'bc_icms': 0.0,
            'v_icms': 0.0,
            'bc_icms_st': 0.0,
            'v_icms_st': 0.0,
            'v_pis': 0.0,
            'v_cofins': 0.0,
            'v_ipi': 0.0,
            'v_frete': 0.0,
            'v_seguro': 0.0,
            'v_desconto': 0.0,
            'v_outras_desp': 0.0,
            'v_prod': 0.0,
            'v_nota': 0.0,
            'valor_total': 0.0,
            'v_icms_uf_dest': 0.0,
            'v_fcp_uf_dest': 0.0,
            'v_tot_trib': 0.0,
            'emitente_nome': '',
            'emitente_cnpj': '',
            'emitente_ie': '',
            'emitente_endereco': '',
            'emitente_municipio': '',
            'emitente_uf': '',
            'emitente_cep': '',
            'destinatario_nome': '',
            'destinatario_cnpj_cpf': '',
            'destinatario_ie': '',
            'destinatario_endereco': '',
            'destinatario_municipio': '',
            'destinatario_uf': '',
            'destinatario_cep': '',
            'frete_por_conta': '',
            'transportador_nome': '',
            'transportador_cnpj_cpf': '',
            'transportador_ie': '',
            'transportador_endereco': '',
            'transportador_municipio': '',
            'transportador_uf': '',
            'placa_veiculo': '',
            'placa_uf': '',
            'antt_rntc': '',
            'quantidade_volumes': '',
            'especie': '',
            'marca_volumes': '',
            'numeracao_volumes': '',
            'peso_bruto': 0.0,
            'peso_liquido': 0.0,
            'fatura_numero': '',
            'fatura_valor_original': 0.0,
            'fatura_valor_desconto': 0.0,
            'fatura_valor_liquido': 0.0,
            'duplicatas': [],
            'formas_pagamento': [],
            'inf_complementar': '',
            'items': [],
        }

    def _new(self) -> Dict[str, Any]:
        return copy.deepcopy(self.template)

    def _clean_number(self, value: Optional[str]) -> float:
        if value in (None, ''):
            return 0.0
        value = str(value).strip()
        if ',' in value and '.' in value:
            value = value.replace('.', '').replace(',', '.')
        elif ',' in value:
            value = value.replace(',', '.')
        try:
            return float(value)
        except Exception:
            value = re.sub(r'[^\d\.\-]', '', value)
            try:
                return float(value)
            except Exception:
                return 0.0

    def _format_date(self, value: str) -> str:
        value = (value or '').strip()
        if not value:
            return ''
        base = value[:10]
        if re.match(r'\d{4}-\d{2}-\d{2}', base):
            yyyy, mm, dd = base.split('-')
            return f'{dd}/{mm}/{yyyy}'
        return base

    def _strip_ns(self, root: ET.Element) -> ET.Element:
        for elem in root.iter():
            if '}' in elem.tag:
                elem.tag = elem.tag.split('}', 1)[1]
        return root

    def _find(self, node: Optional[ET.Element], path: str) -> Optional[ET.Element]:
        if node is None:
            return None
        return node.find(path)

    def _text(self, node: Optional[ET.Element], path: str, default: str = '') -> str:
        if node is None:
            return default
        child = node.find(path)
        if child is None or child.text is None:
            return default
        return child.text.strip()

    def _join_address(self, node: Optional[ET.Element], prefix: str) -> str:
        if node is None:
            return ''
        parts = [
            self._text(node, f'{prefix}/xLgr'),
            self._text(node, f'{prefix}/nro'),
            self._text(node, f'{prefix}/xCpl'),
            self._text(node, f'{prefix}/xBairro'),
            self._text(node, f'{prefix}/xMun'),
            self._text(node, f'{prefix}/UF'),
            self._text(node, f'{prefix}/CEP'),
        ]
        return ' '.join([p for p in parts if p])

    def parse_xml_file(self, file_path: str) -> Dict[str, Any]:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as fh:
            return self.parse_xml_string(fh.read())

    def parse_xml_string(self, xml_text: str) -> Dict[str, Any]:
        root = ET.fromstring(xml_text)
        root = self._strip_ns(root)
        res = self._new()

        inf_nfe = root.find('.//infNFe')
        ide = self._find(inf_nfe, 'ide')
        emit = self._find(inf_nfe, 'emit')
        dest = self._find(inf_nfe, 'dest')
        total = self._find(inf_nfe, 'total/ICMSTot')
        transp = self._find(inf_nfe, 'transp')
        transporta = self._find(transp, 'transporta')
        veic_transp = self._find(transp, 'veicTransp')
        vol = self._find(transp, 'vol')
        cobr = self._find(inf_nfe, 'cobr')
        fat = self._find(cobr, 'fat')
        prot = root.find('.//protNFe/infProt')
        pag = self._find(inf_nfe, 'pag')
        inf_adic = self._find(inf_nfe, 'infAdic')

        res['chave_acesso'] = (inf_nfe.attrib.get('Id', '').replace('NFe', '') if inf_nfe is not None else '')
        res['numero_nota'] = self._text(ide, 'nNF')
        res['serie'] = self._text(ide, 'serie')
        res['natureza_operacao'] = self._text(ide, 'natOp')
        res['protocolo'] = self._text(prot, 'nProt')

        dh_emi = self._text(ide, 'dhEmi') or self._text(ide, 'dEmi')
        dh_saida = self._text(ide, 'dhSaiEnt') or self._text(ide, 'dSaiEnt')
        if dh_emi:
            res['data_emissao'] = self._format_date(dh_emi)
        if dh_saida:
            res['data_saida'] = self._format_date(dh_saida)

        cstat = self._text(prot, 'cStat')
        if cstat in {'100', '150'}:
            res['status_documento'] = 'autorizada'
        elif cstat in {'101', '151', '135', '155'}:
            res['status_documento'] = 'cancelada'
        elif cstat:
            res['status_documento'] = cstat

        res['emitente_nome'] = self._text(emit, 'xNome')
        res['emitente_cnpj'] = self._text(emit, 'CNPJ') or self._text(emit, 'CPF')
        res['emitente_ie'] = self._text(emit, 'IE')
        res['emitente_municipio'] = self._text(emit, 'enderEmit/xMun')
        res['emitente_uf'] = self._text(emit, 'enderEmit/UF')
        res['emitente_cep'] = self._text(emit, 'enderEmit/CEP')
        res['emitente_endereco'] = self._join_address(emit, 'enderEmit')

        res['destinatario_nome'] = self._text(dest, 'xNome')
        res['destinatario_cnpj_cpf'] = self._text(dest, 'CNPJ') or self._text(dest, 'CPF')
        res['destinatario_ie'] = self._text(dest, 'IE')
        res['destinatario_municipio'] = self._text(dest, 'enderDest/xMun')
        res['destinatario_uf'] = self._text(dest, 'enderDest/UF')
        res['destinatario_cep'] = self._text(dest, 'enderDest/CEP')
        res['destinatario_endereco'] = self._join_address(dest, 'enderDest')

        res['bc_icms'] = self._clean_number(self._text(total, 'vBC'))
        res['v_icms'] = self._clean_number(self._text(total, 'vICMS'))
        res['bc_icms_st'] = self._clean_number(self._text(total, 'vBCST'))
        res['v_icms_st'] = self._clean_number(self._text(total, 'vST'))
        res['v_pis'] = self._clean_number(self._text(total, 'vPIS'))
        res['v_cofins'] = self._clean_number(self._text(total, 'vCOFINS'))
        res['v_ipi'] = self._clean_number(self._text(total, 'vIPI'))
        res['v_frete'] = self._clean_number(self._text(total, 'vFrete'))
        res['v_seguro'] = self._clean_number(self._text(total, 'vSeg'))
        res['v_desconto'] = self._clean_number(self._text(total, 'vDesc'))
        res['v_outras_desp'] = self._clean_number(self._text(total, 'vOutro'))
        res['v_prod'] = self._clean_number(self._text(total, 'vProd'))
        res['v_nota'] = self._clean_number(self._text(total, 'vNF'))
        res['valor_total'] = res['v_nota']
        res['v_icms_uf_dest'] = self._clean_number(self._text(total, 'vICMSUFDest'))
        res['v_fcp_uf_dest'] = self._clean_number(self._text(total, 'vFCPUFDest'))
        res['v_tot_trib'] = self._clean_number(self._text(total, 'vTotTrib'))

        res['frete_por_conta'] = self._text(transp, 'modFrete')
        res['transportador_nome'] = self._text(transporta, 'xNome')
        res['transportador_cnpj_cpf'] = self._text(transporta, 'CNPJ') or self._text(transporta, 'CPF')
        res['transportador_ie'] = self._text(transporta, 'IE')
        res['transportador_endereco'] = self._text(transporta, 'xEnder')
        res['transportador_municipio'] = self._text(transporta, 'xMun')
        res['transportador_uf'] = self._text(transporta, 'UF')
        res['placa_veiculo'] = self._text(veic_transp, 'placa')
        res['placa_uf'] = self._text(veic_transp, 'UF')
        res['antt_rntc'] = self._text(veic_transp, 'RNTC')
        res['quantidade_volumes'] = self._text(vol, 'qVol')
        res['especie'] = self._text(vol, 'esp')
        res['marca_volumes'] = self._text(vol, 'marca')
        res['numeracao_volumes'] = self._text(vol, 'nVol')
        res['peso_bruto'] = self._clean_number(self._text(vol, 'pesoB'))
        res['peso_liquido'] = self._clean_number(self._text(vol, 'pesoL'))

        res['fatura_numero'] = self._text(fat, 'nFat')
        res['fatura_valor_original'] = self._clean_number(self._text(fat, 'vOrig'))
        res['fatura_valor_desconto'] = self._clean_number(self._text(fat, 'vDesc'))
        res['fatura_valor_liquido'] = self._clean_number(self._text(fat, 'vLiq'))
        for dup in (cobr.findall('dup') if cobr is not None else []):
            res['duplicatas'].append({
                'numero': self._text(dup, 'nDup'),
                'vencimento': self._text(dup, 'dVenc'),
                'valor': self._clean_number(self._text(dup, 'vDup')),
            })
        if pag is not None:
            for det in pag.findall('detPag'):
                res['formas_pagamento'].append({
                    'tPag': self._text(det, 'tPag'),
                    'xPag': self._text(det, 'xPag'),
                    'valor': self._clean_number(self._text(det, 'vPag')),
                })

        res['inf_complementar'] = self._text(inf_adic, 'infCpl')

        for det in inf_nfe.findall('det') if inf_nfe is not None else []:
            prod = det.find('prod')
            imposto = det.find('imposto')
            icms = imposto.find('ICMS') if imposto is not None else None
            icms_child = list(icms)[0] if icms is not None and list(icms) else None
            ipi = imposto.find('IPI/IPITrib') if imposto is not None else None
            pis_node = None
            cofins_node = None
            if imposto is not None:
                pis = imposto.find('PIS')
                if pis is not None and list(pis):
                    pis_node = list(pis)[0]
                cof = imposto.find('COFINS')
                if cof is not None and list(cof):
                    cofins_node = list(cof)[0]
            res['items'].append({
                'codigo': self._text(prod, 'cProd'),
                'descricao': self._text(prod, 'xProd'),
                'ncm': self._text(prod, 'NCM'),
                'cfop': self._text(prod, 'CFOP'),
                'cst_csosn': self._text(icms_child, 'CST') or self._text(icms_child, 'CSOSN'),
                'cest': self._text(prod, 'CEST'),
                'ean': self._text(prod, 'cEAN'),
                'ean_tributavel': self._text(prod, 'cEANTrib'),
                'unidade': self._text(prod, 'uCom'),
                'qtd': self._clean_number(self._text(prod, 'qCom')),
                'valor_unit': self._clean_number(self._text(prod, 'vUnCom')),
                'valor_total': self._clean_number(self._text(prod, 'vProd')),
                'unidade_tributavel': self._text(prod, 'uTrib'),
                'qtd_tributavel': self._clean_number(self._text(prod, 'qTrib')),
                'valor_unit_tributavel': self._clean_number(self._text(prod, 'vUnTrib')),
                'bc_icms': self._clean_number(self._text(icms_child, 'vBC')),
                'aliq_icms': self._clean_number(self._text(icms_child, 'pICMS')),
                'v_icms': self._clean_number(self._text(icms_child, 'vICMS')),
                'aliq_ipi': self._clean_number(self._text(ipi, 'pIPI')),
                'v_ipi': self._clean_number(self._text(ipi, 'vIPI')),
                'aliq_pis': self._clean_number(self._text(pis_node, 'pPIS')),
                'v_pis': self._clean_number(self._text(pis_node, 'vPIS')),
                'aliq_cofins': self._clean_number(self._text(cofins_node, 'pCOFINS')),
                'v_cofins': self._clean_number(self._text(cofins_node, 'vCOFINS')),
            })
        return res

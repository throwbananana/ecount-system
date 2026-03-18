# -*- coding: utf-8 -*-
"""
文档识别模块 - 巴西NF-e (DANFE) 识别升级版

升级点：
1. 支持 XML 优先解析（同名 XML 或直接导入 XML）。
2. 扩展运输、支付/duplicatas、商品税务细项字段。
3. 对 PDF 继续保留文本抽取 + 启发式兜底。
4. CNPJ 匹配兼容未来字母数字格式。
"""

from __future__ import annotations

import copy
import json
import re
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

try:
    from danfe_xml_parser import DanfeXmlParser
except ImportError:  # pragma: no cover - 便于补丁逐步落地
    DanfeXmlParser = None


class DanfeRecognizer:
    """巴西 NF-e (DANFE) 文档识别器。"""

    CNPJ_ALNUM_RE = r'[A-Z0-9]{2}\.[A-Z0-9]{3}\.[A-Z0-9]{3}/[A-Z0-9]{4}-\d{2}'
    CPF_RE = r'\d{3}\.\d{3}\.\d{3}-\d{2}'

    def __init__(self):
        self.patterns = {
            'chave_acesso': r'CHAVE DE ACESSO\s*([\d\s]{44,80})',
            'numero_nota': r'\bN[º°o]?\s*\.?\s*([\d][\d\.]*)\b',
            'serie': r'S[ÉE]RIE\s*(\d+)',
            'natureza_operacao': r'NATUREZA DA OPERAÇÃO\s*\n?\s*(.*?)(?:\n|PROTOCOLO)',
            'protocolo': r'PROTOCOLO DE AUTORIZAÇÃO DE USO\s*\n?\s*([\d\s\-/: ]+)',
            'data_emissao': r'DATA DA EMISS[ÃA]O\s*(\d{2}/\d{2}/\d{4})',
            'data_saida': r'DATA DA SA[ÍI]DA/ENTRADA\s*(\d{2}/\d{2}/\d{4})',
            'bc_icms': r'BASE DE CÁLC\.? DO ICMS\s*([\d\.,]+)',
            'v_icms': r'VALOR DO ICMS\s*([\d\.,]+)',
            'bc_icms_st': r'BASE DE CÁLC\.? ICMS (?:S\.T\.|ST)\s*([\d\.,]+)',
            'v_icms_st': r'VALOR DO ICMS (?:SUBST\.|ST)\s*([\d\.,]+)',
            'v_pis': r'VALOR DO PIS\s*([\d\.,]+)',
            'v_cofins': r'VALOR DA COFINS\s*([\d\.,]+)',
            'v_ipi': r'VALOR DO IPI\s*([\d\.,]+)',
            'v_frete': r'VALOR DO FRETE\s*([\d\.,]+)',
            'v_seguro': r'VALOR DO SEGURO\s*([\d\.,]+)',
            'v_desconto': r'VALOR DO DESCONTO\s*([\d\.,]+)',
            'v_outras_desp': r'(?:OUTRAS DESPESAS ACESSÓRIAS|OUTRAS DESPESAS)\s*([\d\.,]+)',
            'v_prod': r'V\. TOTAL PRODUTOS\s*([\d\.,]+)',
            'v_nota': r'V\. TOTAL DA NOTA\s*([\d\.,]+)|VALOR TOTAL[:\s]*R?\$?\s*([\d\.,]+)',
            'v_icms_uf_dest': r'V\. ICMS UF DEST\.?\s*([\d\.,]+)',
            'v_fcp_uf_dest': r'V\. FCP UF DEST\.?\s*([\d\.,]+)',
            'v_tot_trib': r'V\. TOT\. TRIB\.?\s*([\d\.,]+)|VALOR APROXIMADO DOS TRIBUTOS\s*[:\s]*R?\$?\s*([\d\.,]+)',
            'peso_bruto': r'PESO BRUTO\s*\n?\s*([\d\.,]+)',
            'peso_liquido': r'PESO L[ÍI]QUIDO\s*\n?\s*([\d\.,]+)',
            'frete_por_conta': r'FRETE POR CONTA\s*\n?\s*([^\n]+)',
        }

        self.units = ['un', 'unid', 'und', 'pç', 'pc', 'kg', 'lt', 'mt', 'cx', 'jg', 'rl', 'pr', 'kt', 'kit']
        self.doc_re = re.compile(rf'(?:{self.CNPJ_ALNUM_RE}|{self.CPF_RE})', re.I)
        self.result_template = {
            'source_used': '',
            'status_documento': '',
            'fonte_status': '',
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
        self.numeric_keys = {k for k, v in self.result_template.items() if isinstance(v, float)}
        self.xml_parser = DanfeXmlParser() if DanfeXmlParser else None

    def _new_result(self) -> Dict[str, Any]:
        return copy.deepcopy(self.result_template)

    def clean_number(self, value: Any) -> float:
        if value is None or value == '':
            return 0.0
        s = str(value).strip()
        if '/' in s and len(s) >= 8:
            return 0.0
        s = s.replace('.', '').replace(',', '.')
        try:
            return float(s)
        except ValueError:
            s = re.sub(r'[^\d\.\-]', '', s)
            try:
                return float(s)
            except Exception:
                return 0.0

    def _first_non_empty_group(self, match: re.Match) -> str:
        for grp in match.groups():
            if grp is not None and str(grp).strip():
                return str(grp).strip()
        raw = match.group(0) if match else ''
        return raw.strip() if raw else ''

    def _extract_decimal_numbers(self, line: str) -> List[float]:
        nums = re.findall(r'\d{1,3}(?:\.\d{3})*,\d{2}', line)
        return [self.clean_number(n) for n in nums]

    def _normalize_doc(self, value: str) -> str:
        value = (value or '').strip().upper()
        return re.sub(r'\s+', ' ', value)

    def _find_first_doc(self, text: str) -> str:
        match = self.doc_re.search(text or '')
        return self._normalize_doc(match.group(0)) if match else ''

    def _extract_chave_acesso(self, text: str, lines: List[str]) -> str:
        for i, line in enumerate(lines):
            if 'CHAVE DE ACESSO' in line.upper():
                block = ' '.join(lines[i:i + 4])
                candidates = re.findall(r'(?:\d[\s]{0,2}){44,60}', block)
                for cand in candidates:
                    digits = re.sub(r'\D', '', cand)
                    if len(digits) >= 44:
                        return digits[:44]
        text_block_match = re.search(r'CHAVE\s+DE\s+ACESSO([\s\S]{0,260})', text, re.I)
        if text_block_match:
            digits = re.sub(r'\D', '', text_block_match.group(1))
            if len(digits) >= 44:
                return digits[:44]
        return ''

    def _extract_emitente_nome(self, text: str, lines: List[str]) -> str:
        m = re.search(r'RECEBEMOS\s+DE\s+(.+?)\s+OS\s+PRODUTOS', text, re.I)
        if m:
            return m.group(1).strip()
        skip_terms = ['DANFE', 'DOCUMENTO AUXILIAR', 'FISCAL ELETRÔNICA', 'CHAVE DE ACESSO', 'CONSULTA DE AUTENTICIDADE', 'FOLHA', 'Nº.', 'SÉRIE']
        for i, line in enumerate(lines[:60]):
            if 'IDENTIFICAÇÃO DO EMITENTE' in line.upper():
                for cand in lines[i + 1:i + 14]:
                    upper = cand.upper()
                    if any(term in upper for term in skip_terms):
                        continue
                    cleaned = re.sub(r'\s+[01]\s*-\s*ENTRADA.*$', '', cand, flags=re.I).strip()
                    cleaned = re.sub(r'\s+[01]\s*-\s*SA[ÍI]DA.*$', '', cleaned, flags=re.I).strip()
                    if cleaned and len(cleaned) >= 4:
                        return cleaned
                break
        return ''

    def _extract_emitente_endereco(self, lines: List[str]) -> str:
        for i, line in enumerate(lines[:60]):
            if 'IDENTIFICAÇÃO DO EMITENTE' in line.upper():
                addr_parts = []
                for cand in lines[i + 1:i + 14]:
                    upper = cand.upper()
                    if any(tag in upper for tag in ['CHAVE DE ACESSO', 'Nº.', 'SÉRIE', 'CNPJ', 'INSCRIÇÃO']):
                        break
                    if any(tag in upper for tag in ['DANFE', 'DOCUMENTO AUXILIAR', 'FISCAL ELETRÔNICA']):
                        continue
                    if re.search(r'\d{5}-\d{3}', cand) or any(x in upper for x in ['VILA', 'RUA', 'AV.', 'AVENIDA', 'SÃO', 'SAO']):
                        addr_parts.append(cand.strip())
                if addr_parts:
                    return ' '.join(addr_parts).strip()
                break
        return ''

    def _extract_emitente_docs(self, text: str) -> Tuple[str, str]:
        cnpj = self._find_first_doc(text)
        ie = ''
        ie_match = re.search(r'INSCRI[ÇC][ÃA]O\s+ESTADUAL\s*\n?\s*([\d\.]{8,20})', text, re.I)
        if ie_match:
            ie = re.sub(r'\D', '', ie_match.group(1))
        return cnpj, ie

    def _extract_destinatario_fields(self, text: str, lines: List[str]) -> Dict[str, str]:
        out = {
            'destinatario_nome': '',
            'destinatario_cnpj_cpf': '',
            'destinatario_ie': '',
            'destinatario_endereco': '',
            'destinatario_municipio': '',
            'destinatario_uf': '',
            'destinatario_cep': '',
            'data_emissao': '',
            'data_saida': '',
        }
        section_idx = -1
        for i, line in enumerate(lines):
            up = line.upper()
            if 'DESTINAT' in up and 'REMETENTE' in up:
                section_idx = i
                break
        if section_idx == -1:
            fallback = re.search(r'DESTINAT[ÁA]RIO:\s*(.+?)(?:\s+-\s+|\n|$)', text, re.I)
            if fallback:
                out['destinatario_nome'] = fallback.group(1).strip()
            return out

        window = lines[section_idx: section_idx + 24]
        for idx, row in enumerate(window):
            up = row.upper()
            if 'NOME / RAZ' in up and 'CNPJ / CPF' in up and idx + 1 < len(window):
                detail = window[idx + 1].strip()
                doc = self._find_first_doc(detail)
                if doc:
                    out['destinatario_cnpj_cpf'] = doc
                    out['destinatario_nome'] = detail[:detail.find(doc)].strip() if detail.find(doc) > 0 else detail
                elif not out['destinatario_nome']:
                    out['destinatario_nome'] = detail
                date_match = re.search(r'(\d{2}/\d{2}/\d{4})$', detail)
                if date_match:
                    out['data_emissao'] = date_match.group(1)
                    if out['destinatario_nome'].endswith(date_match.group(1)):
                        out['destinatario_nome'] = out['destinatario_nome'][:-10].strip()
            if 'ENDERE' in up and idx + 1 < len(window):
                out['destinatario_endereco'] = window[idx + 1].strip()
                cep_match = re.search(r'\b\d{5}-\d{3}\b', out['destinatario_endereco'])
                if cep_match:
                    out['destinatario_cep'] = cep_match.group(0)
                date_match = re.search(r'(\d{2}/\d{2}/\d{4})$', out['destinatario_endereco'])
                if date_match:
                    out['data_saida'] = date_match.group(1)
                    out['destinatario_endereco'] = out['destinatario_endereco'][:date_match.start()].strip()
            if 'MUNIC' in up and 'UF' in up and idx + 1 < len(window):
                city_line = re.sub(r'\b\d{2}:\d{2}:\d{2}\b.*$', '', window[idx + 1]).strip()
                if city_line:
                    tokens = city_line.split()
                    if len(tokens) >= 2:
                        out['destinatario_uf'] = tokens[-1] if len(tokens[-1]) == 2 else out['destinatario_uf']
                        out['destinatario_municipio'] = ' '.join(tokens[:-1]) if len(tokens[-1]) == 2 else city_line
                    else:
                        out['destinatario_municipio'] = city_line
            if 'INSCRIÇÃO ESTADUAL' in up and idx + 1 < len(window) and not out['destinatario_ie']:
                ie_match = re.search(r'\b(\d{8,14})\b', window[idx + 1])
                if ie_match:
                    out['destinatario_ie'] = ie_match.group(1)
        if not out['data_emissao']:
            em_match = re.search(r'EMISS[ÃA]O:\s*(\d{2}/\d{2}/\d{4})', text, re.I)
            if em_match:
                out['data_emissao'] = em_match.group(1)
        if not out['data_saida']:
            saida_match = re.search(r'DATA\s+DA\s+SA[ÍI]DA/ENTRADA\s*\n?\s*.*?(\d{2}/\d{2}/\d{4})', text, re.I)
            if saida_match:
                out['data_saida'] = saida_match.group(1)
        return out

    def _extract_transport_fields(self, text: str, lines: List[str]) -> Dict[str, Any]:
        out = {
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
        }
        section_idx = -1
        for i, line in enumerate(lines):
            up = line.upper()
            if 'TRANSPORTADOR' in up or 'VOLUMES TRANSPORTADOS' in up:
                section_idx = i
                break
        if section_idx != -1:
            window = lines[section_idx: section_idx + 18]
            joined = '\n'.join(window)
            if not out['frete_por_conta']:
                fp = re.search(r'FRETE POR CONTA\s*\n?\s*([^\n]+)', joined, re.I)
                if fp:
                    out['frete_por_conta'] = fp.group(1).strip()
            for idx, row in enumerate(window):
                up = row.upper()
                if 'RAZÃO SOCIAL' in up and idx + 1 < len(window):
                    detail = window[idx + 1].strip()
                    doc = self._find_first_doc(detail)
                    if doc:
                        out['transportador_cnpj_cpf'] = doc
                        out['transportador_nome'] = detail[:detail.find(doc)].strip() if detail.find(doc) > 0 else detail
                    elif not out['transportador_nome']:
                        out['transportador_nome'] = detail
                if 'ENDEREÇO' in up and idx + 1 < len(window):
                    out['transportador_endereco'] = window[idx + 1].strip()
                if 'MUNIC' in up and 'UF' in up and idx + 1 < len(window):
                    city_line = window[idx + 1].strip()
                    city_tokens = city_line.split()
                    if len(city_tokens) >= 2 and len(city_tokens[-1]) == 2:
                        out['transportador_uf'] = city_tokens[-1]
                        out['transportador_municipio'] = ' '.join(city_tokens[:-1])
                    else:
                        out['transportador_municipio'] = city_line
                if 'INSCRIÇÃO ESTADUAL' in up and idx + 1 < len(window):
                    ie_match = re.search(r'\b(\d{8,14})\b', window[idx + 1])
                    if ie_match:
                        out['transportador_ie'] = ie_match.group(1)
                if 'PLACA DO VEÍCULO' in up and idx + 1 < len(window):
                    plate_line = window[idx + 1].strip()
                    plate_match = re.search(r'\b[A-Z]{3}[\- ]?\d[A-Z0-9]\d{2}\b|\b[A-Z]{3}[\- ]?\d{4}\b', plate_line, re.I)
                    if plate_match:
                        out['placa_veiculo'] = plate_match.group(0).replace(' ', '')
                    uf_match = re.search(r'\b([A-Z]{2})\b', plate_line)
                    if uf_match:
                        out['placa_uf'] = uf_match.group(1)
                if 'CÓDIGO ANTT' in up or 'RNTC' in up:
                    code_match = re.search(r'([A-Z0-9\-]{4,})', row)
                    if code_match:
                        out['antt_rntc'] = code_match.group(1)
            qty_match = re.search(r'QUANTIDADE\s*\n?\s*([^\n]+)', joined, re.I)
            if qty_match:
                out['quantidade_volumes'] = qty_match.group(1).strip()
            especie_match = re.search(r'ESPÉCIE\s*\n?\s*([^\n]+)', joined, re.I)
            if especie_match:
                out['especie'] = especie_match.group(1).strip()
            marca_match = re.search(r'MARCA\s*\n?\s*([^\n]+)', joined, re.I)
            if marca_match:
                out['marca_volumes'] = marca_match.group(1).strip()
            numero_match = re.search(r'NUMERAÇÃO\s*\n?\s*([^\n]+)', joined, re.I)
            if numero_match:
                out['numeracao_volumes'] = numero_match.group(1).strip()
        return out

    def _extract_payment_fields(self, text: str, lines: List[str]) -> Dict[str, Any]:
        out: Dict[str, Any] = {
            'fatura_numero': '',
            'fatura_valor_original': 0.0,
            'fatura_valor_desconto': 0.0,
            'fatura_valor_liquido': 0.0,
            'duplicatas': [],
            'formas_pagamento': [],
        }
        section_text = text
        if 'COBRANÇA' in text.upper():
            idx = text.upper().find('COBRANÇA')
            section_text = text[idx: idx + 2000]
        fat_match = re.search(r'FATURA\s*[:\n ]+([A-Z0-9\-/.]+)', section_text, re.I)
        if fat_match:
            out['fatura_numero'] = fat_match.group(1).strip()
        valor_original = re.search(r'VALOR\s+ORIGINAL\s*[:\n ]+([\d\.,]+)', section_text, re.I)
        if valor_original:
            out['fatura_valor_original'] = self.clean_number(valor_original.group(1))
        valor_desc = re.search(r'VALOR\s+DO\s+DESCONTO\s*[:\n ]+([\d\.,]+)', section_text, re.I)
        if valor_desc:
            out['fatura_valor_desconto'] = self.clean_number(valor_desc.group(1))
        valor_liq = re.search(r'VALOR\s+L[ÍI]QUIDO\s*[:\n ]+([\d\.,]+)', section_text, re.I)
        if valor_liq:
            out['fatura_valor_liquido'] = self.clean_number(valor_liq.group(1))
        dup_pattern = re.compile(r'([A-Z0-9\-/.]{2,})\s+(\d{2}/\d{2}/\d{4})\s+([\d\.,]+)')
        seen = set()
        for number, due_date, value in dup_pattern.findall(section_text):
            key = (number, due_date, value)
            if key in seen:
                continue
            seen.add(key)
            out['duplicatas'].append({'numero': number, 'vencimento': due_date, 'valor': self.clean_number(value)})
        for line in lines:
            up = line.upper()
            if 'PAGAMENTO' in up and any(ch.isdigit() for ch in line):
                val_match = re.search(r'([\d\.,]+)$', line)
                if val_match:
                    label = re.sub(r'([\d\.,]+)$', '', line).strip(' :-')
                    out['formas_pagamento'].append({'descricao': label, 'valor': self.clean_number(val_match.group(1))})
        return out

    def _fill_tax_summary_from_table(self, lines: List[str], res: Dict[str, Any]) -> None:
        for i, line in enumerate(lines):
            upper = line.upper()
            if 'V. TOTAL PRODUTOS' in upper and i + 1 < len(lines):
                nums = self._extract_decimal_numbers(lines[i + 1])
                if len(nums) >= 9:
                    if res.get('bc_icms', 0.0) <= 0:
                        res['bc_icms'] = nums[0]
                    if res.get('v_icms', 0.0) <= 0:
                        res['v_icms'] = nums[1]
                    if res.get('bc_icms_st', 0.0) <= 0:
                        res['bc_icms_st'] = nums[2]
                    if res.get('v_icms_st', 0.0) <= 0:
                        res['v_icms_st'] = nums[3]
                    if res.get('v_fcp_uf_dest', 0.0) <= 0:
                        res['v_fcp_uf_dest'] = nums[6]
                    if res.get('v_pis', 0.0) <= 0:
                        res['v_pis'] = nums[7]
                    if res.get('v_prod', 0.0) <= 0:
                        res['v_prod'] = nums[8]
            if 'V. TOTAL DA NOTA' in upper and i + 1 < len(lines):
                nums = self._extract_decimal_numbers(lines[i + 1])
                if nums:
                    if res.get('v_frete', 0.0) <= 0 and len(nums) >= 1:
                        res['v_frete'] = nums[0]
                    if res.get('v_seguro', 0.0) <= 0 and len(nums) >= 2:
                        res['v_seguro'] = nums[1]
                    if res.get('v_desconto', 0.0) <= 0 and len(nums) >= 3:
                        res['v_desconto'] = nums[2]
                    if res.get('v_outras_desp', 0.0) <= 0 and len(nums) >= 4:
                        res['v_outras_desp'] = nums[3]
                    if res.get('v_ipi', 0.0) <= 0 and len(nums) >= 5:
                        res['v_ipi'] = nums[4]
                    if res.get('v_icms_uf_dest', 0.0) <= 0 and len(nums) >= 6:
                        res['v_icms_uf_dest'] = nums[5]
                    if res.get('v_tot_trib', 0.0) <= 0 and len(nums) >= 7:
                        res['v_tot_trib'] = nums[6]
                    if res.get('v_cofins', 0.0) <= 0 and len(nums) >= 8:
                        res['v_cofins'] = nums[7]
                    if res.get('v_nota', 0.0) <= 0:
                        res['v_nota'] = nums[-1]
                break

    def _serialize_for_sheet(self, value: Any) -> str:
        if isinstance(value, (list, dict)):
            return json.dumps(value, ensure_ascii=False)
        return value if value is not None else ''

    def _merge_results(self, base: Dict[str, Any], override: Dict[str, Any]) -> Dict[str, Any]:
        merged = copy.deepcopy(base)
        for key, value in (override or {}).items():
            if key not in merged:
                merged[key] = copy.deepcopy(value)
                continue
            if key in ('items', 'duplicatas', 'formas_pagamento'):
                if value:
                    merged[key] = copy.deepcopy(value)
                continue
            if key in self.numeric_keys:
                if value not in (None, '', 0, 0.0):
                    merged[key] = value
                continue
            if isinstance(value, str):
                if value.strip():
                    merged[key] = value.strip()
                continue
            if value not in (None, ''):
                merged[key] = copy.deepcopy(value)
        if merged.get('v_nota', 0.0) and not merged.get('valor_total', 0.0):
            merged['valor_total'] = merged['v_nota']
        return merged

    def _extract_items_from_text(self, lines: List[str]) -> List[Dict[str, Any]]:
        items: List[Dict[str, Any]] = []
        start_idx = -1
        for i, line in enumerate(lines):
            if any(kw in line.upper() for kw in ['DADOS DOS PRODUTOS', 'PRODUTOS / SERVIÇOS']):
                start_idx = i
                break
        if start_idx == -1:
            return items

        buffer_desc: List[str] = []
        pending_code = ''
        header_terms = ['CÓDIGO', 'DESCRIÇÃO', 'NCM/SH', 'O/CST', 'CFOP', 'UN', 'QUANT', 'VALOR', 'UNIT', 'TOTAL', 'B.CÁLC', 'ALÍQ']
        for line in lines[start_idx + 1:]:
            if any(kw in line.upper() for kw in ['DADOS ADICIONAIS', 'RESERVADO', 'INFORMAÇÕES COMPLEMENTARES']):
                break
            parts = line.split()
            if not parts:
                continue
            if any(term in line.upper() for term in ['CÓDIGO PRODUTO', 'PFCPUFDEST', 'PICMSUFDEST', 'VICMSUFDEST', 'VICMSUFREMET', '(PEDIDO']):
                continue
            u_idx = -1
            for idx, part in enumerate(parts):
                p_clean = re.sub(r'[\d\.,]', '', part).lower().strip()
                if p_clean in self.units and not any(c.isdigit() for c in part):
                    if idx + 1 < len(parts) and self.clean_number(parts[idx + 1]) >= 0:
                        u_idx = idx
                        break
            ncm_idx, cfop_idx = -1, -1
            cst_idx = -1
            if u_idx == -1:
                for idx, part in enumerate(parts):
                    raw = re.sub(r'\D', '', part)
                    if len(raw) == 8 and ncm_idx == -1:
                        ncm_idx = idx
                        continue
                    if ncm_idx != -1 and len(raw) in (2, 3) and cst_idx == -1:
                        cst_idx = idx
                        continue
                    if ncm_idx != -1 and len(raw) == 4:
                        cfop_idx = idx
                        break
                if cfop_idx != -1 and cfop_idx + 1 < len(parts):
                    if any(c.isalpha() for c in parts[cfop_idx + 1]):
                        u_idx = cfop_idx + 1
                    else:
                        u_idx = cfop_idx
            if u_idx == -1:
                if not any(term in parts[0].upper() for term in header_terms):
                    if not pending_code and len(parts[0]) > 3 and (any(c.isdigit() for c in parts[0]) or '-' in parts[0]):
                        pending_code = parts[0]
                        buffer_desc.append(' '.join(parts[1:]))
                    else:
                        buffer_desc.append(line)
                continue

            desc = ' '.join(buffer_desc).strip()
            if ncm_idx == -1:
                ncm_idx = next((idx for idx, p in enumerate(parts[:u_idx + 1]) if len(re.sub(r'\D', '', p)) == 8), -1)
            ncm = re.sub(r'\D', '', parts[ncm_idx])[:8] if ncm_idx >= 0 else ''
            if cfop_idx == -1 and ncm_idx != -1:
                for idx in range(ncm_idx + 1, min(len(parts), u_idx + 1)):
                    raw = re.sub(r'\D', '', parts[idx])
                    if len(raw) == 4:
                        cfop_idx = idx
                        break
            if cst_idx == -1 and ncm_idx != -1:
                for idx in range(ncm_idx + 1, min(len(parts), u_idx + 1)):
                    raw = re.sub(r'\D', '', parts[idx])
                    if len(raw) in (2, 3):
                        cst_idx = idx
                        break
            code = pending_code
            if not code:
                for p in parts[:u_idx]:
                    if len(p) > 2 and p != ncm and (any(c.isdigit() for c in p) or '-' in p):
                        code = p
                        break
            if not desc:
                if ncm_idx > 0:
                    desc = ' '.join(parts[1:ncm_idx]).strip()
                elif u_idx > 1:
                    desc = ' '.join(parts[1:u_idx]).strip()
                else:
                    desc = line.strip()
            nums: List[float] = []
            idx_nums: List[Tuple[int, float]] = []
            for idx, p in enumerate(parts):
                if any(c.isdigit() for c in p):
                    val = self.clean_number(p)
                    raw = p.replace('.', '').replace(',', '')
                    if raw.isdigit() and len(raw) in (4, 8):
                        continue
                    if idx > u_idx or (cfop_idx != -1 and idx > cfop_idx):
                        if val > 0 or '0,00' in p or p == '0':
                            nums.append(val)
                            idx_nums.append((idx, val))
            item = {
                'codigo': code,
                'descricao': desc,
                'ncm': ncm,
                'cfop': re.sub(r'\D', '', parts[cfop_idx]) if cfop_idx != -1 else '',
                'cst_csosn': re.sub(r'\D', '', parts[cst_idx]) if cst_idx != -1 else '',
                'cest': '',
                'ean': '',
                'ean_tributavel': '',
                'unidade': parts[u_idx] if u_idx < len(parts) and any(c.isalpha() for c in parts[u_idx]) else 'un',
                'qtd': 0.0,
                'valor_unit': 0.0,
                'valor_total': 0.0,
                'unidade_tributavel': '',
                'qtd_tributavel': 0.0,
                'valor_unit_tributavel': 0.0,
                'bc_icms': 0.0,
                'aliq_icms': 0.0,
                'v_icms': 0.0,
                'aliq_ipi': 0.0,
                'v_ipi': 0.0,
                'aliq_pis': 0.0,
                'v_pis': 0.0,
                'aliq_cofins': 0.0,
                'v_cofins': 0.0,
            }
            if len(nums) >= 3:
                item['qtd'], item['valor_unit'], item['valor_total'] = nums[0], nums[1], nums[2]
            if len(nums) >= 5:
                item['bc_icms'], item['v_icms'] = nums[3], nums[4]
            if len(nums) >= 6:
                item['aliq_icms'] = nums[5]
            items.append(item)
            buffer_desc = []
            pending_code = ''
        return items

    def recognize_from_text(self, text: str) -> Dict[str, Any]:
        res = self._new_result()
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        res['source_used'] = 'pdf_text'
        for key, pattern in self.patterns.items():
            match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
            if not match:
                continue
            val = self._first_non_empty_group(match)
            if key in self.numeric_keys:
                res[key] = self.clean_number(val)
            elif key == 'chave_acesso':
                res[key] = re.sub(r'\D', '', val)
            else:
                res[key] = val
        if not res['chave_acesso'] or len(res['chave_acesso']) < 44:
            res['chave_acesso'] = self._extract_chave_acesso(text, lines)
        if not res['numero_nota'] or res['numero_nota'].strip('.') == '':
            note_match = re.search(r'\bN[º°o]?\s*\.?\s*([\d][\d\.]*)\b', text, re.I)
            if note_match:
                res['numero_nota'] = note_match.group(1).strip()
        emitente_nome = self._extract_emitente_nome(text, lines)
        if emitente_nome:
            res['emitente_nome'] = emitente_nome
        if not res.get('emitente_endereco'):
            emitente_addr = self._extract_emitente_endereco(lines)
            if emitente_addr:
                res['emitente_endereco'] = emitente_addr
                cep_match = re.search(r'\b\d{5}-\d{3}\b', emitente_addr)
                if cep_match:
                    res['emitente_cep'] = cep_match.group(0)
        emit_cnpj, emit_ie = self._extract_emitente_docs(text)
        if emit_cnpj:
            res['emitente_cnpj'] = emit_cnpj
        if emit_ie:
            res['emitente_ie'] = emit_ie
        dest_data = self._extract_destinatario_fields(text, lines)
        for k, v in dest_data.items():
            if v and not res.get(k):
                res[k] = v
        self._fill_tax_summary_from_table(lines, res)
        if res.get('v_nota', 0.0) <= 0:
            total_match = re.search(r'VALOR\s+TOTAL\s*:\s*R\$\s*([\d\.,]+)', text, re.I)
            if total_match:
                res['v_nota'] = self.clean_number(total_match.group(1))
        res['valor_total'] = res['v_nota']
        transport_data = self._extract_transport_fields(text, lines)
        res = self._merge_results(res, transport_data)
        payment_data = self._extract_payment_fields(text, lines)
        res = self._merge_results(res, payment_data)
        if 'INFORMAÇÕES COMPLEMENTARES' in text.upper():
            idx = text.upper().find('INFORMAÇÕES COMPLEMENTARES')
            inf_part = text[idx + len('INFORMAÇÕES COMPLEMENTARES'):]
            inf_end_idx = len(inf_part)
            for kw in ['RESERVADO AO FISCO', 'DADOS DOS PRODUTOS', 'CÁLCULO DO ISSQN']:
                pos = inf_part.upper().find(kw)
                if pos != -1 and pos < inf_end_idx:
                    inf_end_idx = pos
            res['inf_complementar'] = inf_part[:inf_end_idx].strip()
        res['items'] = self._extract_items_from_text(lines)
        return res

    def recognize_from_xml(self, xml_text: str) -> Dict[str, Any]:
        if not self.xml_parser:
            raise RuntimeError('danfe_xml_parser.py 不可用，无法执行 XML 解析')
        res = self.xml_parser.parse_xml_string(xml_text)
        res['source_used'] = 'xml'
        if res.get('v_nota', 0.0) and not res.get('valor_total', 0.0):
            res['valor_total'] = res['v_nota']
        return self._merge_results(self._new_result(), res)

    def recognize_document(self, text: Optional[str] = None, xml_text: Optional[str] = None) -> Dict[str, Any]:
        text = text or ''
        xml_text = xml_text or ''
        if xml_text and text:
            pdf_res = self.recognize_from_text(text)
            xml_res = self.recognize_from_xml(xml_text)
            merged = self._merge_results(pdf_res, xml_res)
            merged['source_used'] = 'xml+pdf'
            return merged
        if xml_text:
            return self.recognize_from_xml(xml_text)
        return self.recognize_from_text(text)

    def to_comprehensive_dataframe(self, results: List[Dict[str, Any]]) -> pd.DataFrame:
        rows = []
        for res in results:
            base = {
                '文件路径': res.get('file_path', ''),
                '来源': res.get('source_used', ''),
                '单据状态': res.get('status_documento', ''),
                '状态来源': res.get('fonte_status', ''),
                'Access Key (Chave)': res.get('chave_acesso', ''),
                '发票号码': res.get('numero_nota', ''),
                '系列 (Série)': res.get('serie', ''),
                '业务性质': res.get('natureza_operacao', ''),
                '日期': res.get('data_emissao', ''),
                '出库日期': res.get('data_saida', ''),
                '发行人': res.get('emitente_nome', ''),
                '发行人CNPJ': res.get('emitente_cnpj', ''),
                '发行人IE': res.get('emitente_ie', ''),
                '发行人地址': res.get('emitente_endereco', ''),
                '发行人城市': res.get('emitente_municipio', ''),
                '发行人UF': res.get('emitente_uf', ''),
                '发行人CEP': res.get('emitente_cep', ''),
                '收件人': res.get('destinatario_nome', ''),
                '收件人ID (CNPJ/CPF)': res.get('destinatario_cnpj_cpf', ''),
                '收件人IE': res.get('destinatario_ie', ''),
                '收件人地址': res.get('destinatario_endereco', ''),
                '收件人城市': res.get('destinatario_municipio', ''),
                '收件人UF': res.get('destinatario_uf', ''),
                '收件人CEP': res.get('destinatario_cep', ''),
                'ICMS底数': res.get('bc_icms', 0.0),
                'ICMS金额': res.get('v_icms', 0.0),
                'ICMS ST底数': res.get('bc_icms_st', 0.0),
                'ICMS ST金额': res.get('v_icms_st', 0.0),
                'PIS金额': res.get('v_pis', 0.0),
                'COFINS金额': res.get('v_cofins', 0.0),
                'IPI金额': res.get('v_ipi', 0.0),
                '运费': res.get('v_frete', 0.0),
                '保险': res.get('v_seguro', 0.0),
                '折扣': res.get('v_desconto', 0.0),
                '其他费用': res.get('v_outras_desp', 0.0),
                '商品总计': res.get('v_prod', 0.0),
                '发票总额': res.get('v_nota', 0.0),
                'ICMS UF Dest金额': res.get('v_icms_uf_dest', 0.0),
                'FCP UF Dest金额': res.get('v_fcp_uf_dest', 0.0),
                '总税贡献 (Trib)': res.get('v_tot_trib', 0.0),
                '运费承担': res.get('frete_por_conta', ''),
                '运输商': res.get('transportador_nome', ''),
                '运输商ID': res.get('transportador_cnpj_cpf', ''),
                '运输商IE': res.get('transportador_ie', ''),
                '运输商地址': res.get('transportador_endereco', ''),
                '运输商城市': res.get('transportador_municipio', ''),
                '运输商UF': res.get('transportador_uf', ''),
                '车牌': res.get('placa_veiculo', ''),
                '车牌UF': res.get('placa_uf', ''),
                'ANTT/RNTC': res.get('antt_rntc', ''),
                '件数': res.get('quantidade_volumes', ''),
                '包装种类': res.get('especie', ''),
                '包装品牌': res.get('marca_volumes', ''),
                '包装编号': res.get('numeracao_volumes', ''),
                '毛重 (Peso Bruto)': res.get('peso_bruto', 0.0),
                '净重 (Peso Líquido)': res.get('peso_liquido', 0.0),
                'Fatura编号': res.get('fatura_numero', ''),
                'Fatura原值': res.get('fatura_valor_original', 0.0),
                'Fatura折扣': res.get('fatura_valor_desconto', 0.0),
                'Fatura净值': res.get('fatura_valor_liquido', 0.0),
                'DuplicatasJSON': self._serialize_for_sheet(res.get('duplicatas', [])),
                'PagamentoJSON': self._serialize_for_sheet(res.get('formas_pagamento', [])),
                '补充信息': res.get('inf_complementar', ''),
            }
            if res.get('items'):
                for item in res['items']:
                    row = base.copy()
                    row.update({
                        '商品代码': item.get('codigo', ''),
                        '商品描述': item.get('descricao', ''),
                        'NCM': item.get('ncm', ''),
                        'CFOP': item.get('cfop', ''),
                        'CST/CSOSN': item.get('cst_csosn', ''),
                        'CEST': item.get('cest', ''),
                        'EAN': item.get('ean', ''),
                        'EAN Tributável': item.get('ean_tributavel', ''),
                        '单位': item.get('unidade', ''),
                        '数量': item.get('qtd', 0.0),
                        '单价': item.get('valor_unit', 0.0),
                        '商品总价': item.get('valor_total', 0.0),
                        'Trib单位': item.get('unidade_tributavel', ''),
                        'Trib数量': item.get('qtd_tributavel', 0.0),
                        'Trib单价': item.get('valor_unit_tributavel', 0.0),
                        '项目ICMS底数': item.get('bc_icms', 0.0),
                        '项目ICMS税率': item.get('aliq_icms', 0.0),
                        '项目ICMS金额': item.get('v_icms', 0.0),
                        '项目IPI税率': item.get('aliq_ipi', 0.0),
                        '项目IPI金额': item.get('v_ipi', 0.0),
                        '项目PIS税率': item.get('aliq_pis', 0.0),
                        '项目PIS金额': item.get('v_pis', 0.0),
                        '项目COFINS税率': item.get('aliq_cofins', 0.0),
                        '项目COFINS金额': item.get('v_cofins', 0.0),
                    })
                    rows.append(row)
            else:
                row = base.copy()
                row.update({'商品描述': '无明细项'})
                rows.append(row)
        return pd.DataFrame(rows)

    def to_standard_voucher(self, results: List[Dict[str, Any]]) -> pd.DataFrame:
        rows = []
        for res in results:
            d = res.get('data_emissao', '') or ''
            date_str = f"{d[6:]}{d[3:5]}{d[:2]}" if len(d) == 10 else d
            numero_nota = str(res.get('numero_nota', ''))
            nota_digits = re.sub(r'\D', '', numero_nota)
            seq = nota_digits[-4:] if nota_digits else numero_nota[-4:]
            summary = f"NF-e{numero_nota} | {res.get('natureza_operacao', '')[:20]} | {res.get('emitente_nome', '')}"
            rows.append({
                '凭证日期': date_str,
                '序号': seq,
                '会计凭证No.': numero_nota,
                '摘要': summary,
                '类型': '3',
                '科目编码': '',
                '往来单位编码': res.get('emitente_cnpj', ''),
                '往来单位名': res.get('emitente_nome', ''),
                '金额': res.get('v_nota', 0.0),
                '外币金额': 0.0,
                '汇率': 1.0,
                '部门': '',
                '备注/附加信息': f"Key:{str(res.get('chave_acesso', ''))[-4:]}; ICMS:{res.get('v_icms', 0.0)}",
            })
            for item in res.get('items', []):
                item_desc = (item.get('descricao') or item.get('codigo') or '').strip()
                rows.append({
                    '凭证日期': date_str,
                    '序号': seq,
                    '会计凭证No.': numero_nota,
                    '摘要': f"[Item] {item_desc[:100]}",
                    '类型': '3',
                    '科目编码': item.get('codigo', ''),
                    '往来单位编码': res.get('emitente_cnpj', ''),
                    '往来单位名': res.get('emitente_nome', ''),
                    '金额': item.get('valor_total', 0.0),
                    '外币金额': item.get('qtd', 0.0),
                    '汇率': item.get('valor_unit', 0.0),
                    '部门': item.get('unidade', ''),
                    '备注/附加信息': f"NCM:{item.get('ncm', '')}; CFOP:{item.get('cfop', '')}",
                })
        return pd.DataFrame(rows)

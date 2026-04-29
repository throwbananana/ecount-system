# -*- coding: utf-8 -*-
"""
文档识别模块 - 巴西NF-e (DANFE) 识别 (全字段增强版)
支持提取发票抬头、发行人、收件人、税金汇总、物流及商品明细等全量字段。
"""

import os
import re
import unicodedata
import pandas as pd
from typing import Dict, List, Any, Tuple

class DanfeRecognizer:
    """巴西NF-e (DANFE) 文档识别器 - 能够识别发票所有维度的数据"""

    def __init__(self):
        # 1. 核心字段正则表达式模式 (支持多种变体和换行)
        self.patterns = {
            'chave_acesso': r'CHAVE DE ACESSO\s*([\d\s]{44,80})',
            'numero_nota': r'(?:\bN[º°o]\s*\.?|\bN\.)\s*([\d][\d\.]{5,})\b',
            'serie': r'S[ÉE]RIE\s*(\d+)',
            'natureza_operacao': r'NATUREZA DA OPERAÇÃO\s*\n?\s*(.*?)(?:\n|PROTOCOLO)',
            'protocolo': r'PROTOCOLO DE AUTORIZAÇÃO DE USO\s*\n?\s*([\d\s\-/: ]+)',
            'data_emissao': r'DATA DA EMISS[ÃA]O\s*(\d{2}/\d{2}/\d{4})',
            'data_saida': r'DATA DA SA[ÍI]DA/ENTRADA\s*(\d{2}/\d{2}/\d{4})',
            
            # 税金汇总字段
            'bc_icms': r'BASE DE CÁLC\. DO ICMS\s*([\d\.,]+)',
            'v_icms': r'VALOR DO ICMS\s*([\d\.,]+)',
            'bc_icms_st': r'BASE DE CÁLC\. ICMS (?:S\.T\.|ST)\s*([\d\.,]+)',
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
            'v_icms_uf_dest': r'V\. ICMS UF DEST\.\s*([\d\.,]+)',
            'v_fcp_uf_dest': r'V\. FCP UF DEST\.\s*([\d\.,]+)',
            'v_tot_trib': r'V\. TOT\. TRIB\.\s*([\d\.,]+)|VALOR APROXIMADO DOS TRIBUTOS\s*[:\s]*R?\$?\s*([\d\.,]+)',
            
            # 物流及辅助字段
            'peso_bruto': r'PESO BRUTO\s*\n?\s*([\d\.,]+)',
            'peso_liquido': r'PESO LÍQUIDO\s*\n?\s*([\d\.,]+)',
            'cnpj_emitente': r'CNPJ\s*(?:/\s*CPF)?\s*\n?\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})',
            'ie_emitente': r'INSCRIÇÃO ESTADUAL\s*\n?\s*([\d\.]+)',
        }
        
        self.units = ['un', 'unid', 'und', 'pç', 'pc', 'kg', 'lt', 'mt', 'cx', 'jg', 'rl', 'pr', 'kt', 'kit']

    def clean_number(self, value: Any) -> float:
        """清理巴西格式数字: 1.234,56 -> 1234.56"""
        if value is None or value == "": return 0.0
        s = str(value).strip()
        if '/' in s and len(s) >= 8: return 0.0
        s = s.replace('.', '').replace(',', '.')
        try: return float(s)
        except ValueError:
            s = re.sub(r'[^\d\.\-]', '', s)
            try: return float(s)
            except: return 0.0

    def _first_non_empty_group(self, match: re.Match) -> str:
        """兼容多分组正则，优先返回第一个非空分组。"""
        for grp in match.groups():
            if grp is not None and str(grp).strip():
                return str(grp).strip()
        raw = match.group(0) if match else ""
        return raw.strip() if raw else ""

    def _extract_decimal_numbers(self, line: str) -> List[float]:
        """提取行内巴西小数格式数字。"""
        nums = re.findall(r'\d{1,3}(?:\.\d{3})*,\d{2}', line)
        return [self.clean_number(n) for n in nums]

    def _merge_description(self, original: str, fragment: str) -> str:
        """将续行文本合并到商品描述，避免重复拼接。"""
        original = (original or "").strip()
        fragment = (fragment or "").strip()
        if not fragment:
            return original
        if not original:
            return fragment

        original_norm = re.sub(r'\s+', ' ', original)
        fragment_norm = re.sub(r'\s+', ' ', fragment)
        if fragment_norm in original_norm:
            return original
        return f"{original} {fragment}".strip()

    def _normalize_label(self, value: str) -> str:
        """归一化 PDF 文本标签，兼容无重音、标点分裂和多空格。"""
        if not value:
            return ""
        normalized = unicodedata.normalize("NFKD", value)
        normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
        normalized = normalized.upper()
        normalized = re.sub(r'[^A-Z0-9]+', ' ', normalized)
        return re.sub(r'\s+', ' ', normalized).strip()

    def _extract_chave_acesso_from_filename(self, file_path: str) -> str:
        """从文件名兜底提取 44 位 access key。"""
        if not file_path:
            return ""
        match = re.search(r'_(\d{44})-DANFE\.pdf$', os.path.basename(file_path), re.I)
        return match.group(1) if match else ""

    def _find_item_markers(self, parts: List[str]) -> Tuple[int, int, int]:
        """定位商品行中的 NCM、CFOP 与单位列。"""
        ncm_idx, cfop_idx = -1, -1
        for idx, part in enumerate(parts):
            if len(part) == 8 and part.isdigit():
                ncm_idx = idx
            if len(part) == 4 and part.isdigit() and ncm_idx != -1:
                cfop_idx = idx
                break

        u_idx = -1
        unit_search_start = cfop_idx + 1 if cfop_idx != -1 else 0
        for idx in range(unit_search_start, len(parts)):
            part = parts[idx]
            part_clean = re.sub(r'[\d\.,]', '', part).lower().strip()
            if part_clean in self.units and not any(c.isdigit() for c in part):
                if idx + 1 < len(parts) and self.clean_number(parts[idx + 1]) >= 0:
                    u_idx = idx
                    break

        if u_idx == -1 and cfop_idx != -1 and cfop_idx + 1 < len(parts):
            if any(c.isalpha() for c in parts[cfop_idx + 1]):
                u_idx = cfop_idx + 1
            else:
                u_idx = cfop_idx

        return ncm_idx, cfop_idx, u_idx

    def _looks_like_measurement_token(self, token: str) -> bool:
        """判断 token 是否更像规格值而不是商品编码。"""
        token = (token or "").strip().lower()
        if not token:
            return False
        return bool(re.fullmatch(r'\d+(?:[\.,]\d+)?(?:mm|cm|m|kg|g|mg|ml|l|v|w|kw|cv|hp)', token))

    def _looks_like_product_code(self, token: str) -> bool:
        """判断 token 是否更像商品编码。"""
        token = (token or "").strip()
        if len(token) < 3 or '=' in token or token.endswith(')'):
            return False
        if self._looks_like_measurement_token(token):
            return False
        if re.fullmatch(r'\d+(?:[\.,]\d+)?', token):
            return False

        has_alpha = any(c.isalpha() for c in token)
        has_digit = any(c.isdigit() for c in token)
        if not (has_alpha and has_digit):
            return False
        if any(sep in token for sep in ['-', '_', '/']):
            return True
        return token.upper() == token and len(token) >= 5

    def _is_item_metadata_line(self, line: str, line_label: str) -> bool:
        """识别商品行之间的税务/订单补充文本。"""
        compact = re.sub(r'\s+', '', line_label)
        metadata_terms = [
            "PFCPUFDEST", "PICMSUFDEST", "PICMSINTERPART", "VFCPUFDEST",
            "VICMSUFDEST", "VICMSUFREMET", "PEDIDO"
        ]
        if any(term in compact for term in metadata_terms):
            return True
        return bool(re.fullmatch(r'[A-Za-z0-9\-_/\.]+\)', line.strip()))

    def _extract_numero_nota(self, text: str, lines: List[str]) -> str:
        """提取发票号码，避免误抓地址中的 N 108 或商品编码中的 N60。"""
        pattern = re.compile(r'(?:\bN[º°o]\s*\.?|\bN\.)\s*([\d][\d\.]{5,})\b', re.I)
        candidates = []

        for idx, line in enumerate(lines[:40]):
            for match in pattern.finditer(line):
                val = match.group(1).strip()
                digits = re.sub(r'\D', '', val)
                if len(digits) >= 6:
                    candidates.append((len(digits), idx, val))

        if candidates:
            candidates.sort(key=lambda item: (-item[0], item[1]))
            return candidates[0][2]

        text_match = pattern.search(text)
        if text_match:
            return text_match.group(1).strip()
        return ""

    def _extract_chave_acesso(self, text: str, lines: List[str]) -> str:
        """提取 44 位 access key。"""
        for i, line in enumerate(lines):
            if not re.search(r'CHAVE\s+DE\s+ACE(?:SSO)?', line, re.I):
                continue

            raw_candidates = []
            split_match = re.split(r'CHAVE\s+DE\s+ACE(?:SSO)?', line, flags=re.I, maxsplit=1)
            same_line_tail = split_match[-1].strip() if len(split_match) > 1 else ""
            if same_line_tail:
                raw_candidates.append(same_line_tail)
            raw_candidates.extend(lines[i + 1:i + 4])

            for cand in raw_candidates:
                cand_digits = re.sub(r'\D', '', cand)
                if len(cand_digits) >= 44:
                    return cand_digits[:44]
                exact_match = re.search(r'(?<!\d)((?:\d\s*){44})(?!\d)', cand)
                if exact_match:
                    digits = re.sub(r'\D', '', exact_match.group(1))
                    if len(digits) == 44:
                        return digits

            combined = " ".join(raw_candidates)
            compact = re.sub(r'\D', '', combined)
            if len(compact) >= 44:
                return compact[:44]

        text_block_match = re.search(r'CHAVE\s+DE\s+ACE(?:SSO)?([\s\S]{0,260})', text, re.I)
        if text_block_match:
            digits = re.sub(r'\D', '', text_block_match.group(1))
            if len(digits) >= 44:
                return digits[:44]
        return ""

    def _extract_natureza_protocolo(self, text: str, lines: List[str]) -> Tuple[str, str]:
        """从业务性质/协议号区块提取有效内容，避免返回表头。"""
        natureza = ""
        protocolo = ""
        protocol_pattern = re.compile(r'(\d{10,}\s*(?:-\s*)?\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})')

        for i, line in enumerate(lines):
            line_label = self._normalize_label(line)
            if "NATUREZA DA OPERACAO" not in line_label:
                continue

            candidates = []
            candidates.extend(lines[i + 1:i + 4])

            for cand in candidates:
                cand = cand.strip()
                if not cand:
                    continue
                cand_label = self._normalize_label(cand)
                proto_match = protocol_pattern.search(cand)
                if proto_match:
                    protocolo = proto_match.group(1).strip()
                    left = cand[:proto_match.start()].strip(" -")
                    if left and "PROTOCOLO DE AUTORIZACAO DE USO" not in self._normalize_label(left):
                        natureza = left
                    break
                if "PROTOCOLO DE AUTORIZACAO DE USO" in cand_label:
                    continue
                if not natureza and not any(
                    token in cand_label
                    for token in ["INSCRICAO ESTADUAL", "CHAVE DE ACESSO", "CHAVE DE ACE", "CONSULTA DE AUTENTICIDADE"]
                ):
                    natureza = cand

            if natureza or protocolo:
                break

        if not protocolo:
            text_proto = protocol_pattern.search(text)
            if text_proto:
                protocolo = text_proto.group(1).strip()

        return natureza, protocolo

    def _extract_emitente_nome(self, text: str, lines: List[str]) -> str:
        """提取发行人名称，避免误识别成 DANFE。"""
        m = re.search(r'RECEBEMOS\s+DE\s+(.+?)\s+OS\s+PRODUTOS', text, re.I)
        if m:
            return m.group(1).strip()

        skip_terms = [
            "DANFE", "DOCUMENTO AUXILIAR", "FISCAL ELETRÔNICA", "CHAVE DE ACESSO",
            "CONSULTA DE AUTENTICIDADE", "FOLHA", "Nº.", "SÉRIE"
        ]
        for i, line in enumerate(lines[:60]):
            if "IDENTIFICAÇÃO DO EMITENTE" in line.upper():
                for cand in lines[i+1:i+14]:
                    upper = cand.upper()
                    if any(term in upper for term in skip_terms):
                        continue
                    cleaned = re.sub(r'\s+[01]\s*-\s*ENTRADA.*$', '', cand, flags=re.I).strip()
                    cleaned = re.sub(r'\s+[01]\s*-\s*SA[ÍI]DA.*$', '', cleaned, flags=re.I).strip()
                    if cleaned and len(cleaned) >= 4:
                        return cleaned
                break
        return ""

    def _extract_emitente_endereco(self, lines: List[str]) -> str:
        """提取发行人地址信息。"""
        for i, line in enumerate(lines[:60]):
            if "IDENTIFICAÇÃO DO EMITENTE" in line.upper():
                addr_parts = []
                for cand in lines[i + 1:i + 14]:
                    upper = cand.upper()
                    if any(tag in upper for tag in ["DANFE", "DOCUMENTO AUXILIAR", "FISCAL ELETRÔNICA"]):
                        continue

                    cleaned = re.sub(r'\s+[01]\s*-\s*ENTRADA.*$', '', cand, flags=re.I).strip()
                    cleaned = re.sub(r'\s+[01]\s*-\s*SA[ÍI]DA.*$', '', cleaned, flags=re.I).strip()
                    cleaned = re.sub(r'\s+CHAVE DE ACESSO.*$', '', cleaned, flags=re.I).strip()
                    cleaned = re.sub(r'\s+CONSULTA DE AUTENTICIDADE.*$', '', cleaned, flags=re.I).strip()
                    cleaned = re.sub(r'\s+FONE/FAX:.*$', '', cleaned, flags=re.I).strip()
                    cleaned = re.sub(r'\s+N[º°o]?\.\s*[\d\.]+.*$', '', cleaned, flags=re.I).strip()
                    cleaned = re.sub(r'\s+S[ÉE]RIE\s+\d+.*$', '', cleaned, flags=re.I).strip()

                    if not cleaned:
                        if any(tag in upper for tag in ["CHAVE DE ACESSO", "Nº.", "SÉRIE", "CNPJ", "INSCRIÇÃO"]):
                            break
                        continue
                    if re.fullmatch(r'(?:\d\s*){44}', cleaned):
                        continue

                    cleaned_upper = cleaned.upper()
                    if (
                        re.search(r'\d{5}-\d{3}', cleaned)
                        or any(x in cleaned_upper for x in ["VILA", "RUA", "AV.", "AVENIDA", "SÃO", "SAO", "SALA"])
                        or (addr_parts and re.search(r'[A-Za-zÀ-ÿ\s]+-\s*[A-Z]{2}\b', cleaned))
                    ):
                        addr_parts.append(cleaned)
                if addr_parts:
                    return " ".join(addr_parts).strip()
                break
        return ""

    def _extract_emitente_docs(self, text: str) -> Tuple[str, str]:
        """提取发行人 CNPJ 与 IE。"""
        cnpj = ""
        ie = ""
        cnpj_pat = r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}'

        row_match = re.search(
            r'INSCRI[ÇC][ÃA]O\s+ESTADUAL\s+INSCRI[ÇC][ÃA]O\s+MUNICIPAL\s+'
            r'INSCRI[ÇC][ÃA]O\s+ESTADUAL\s+DO\s+SUBST\.\s+TRIBUT\.\s+CNPJ\s*\n\s*([^\n]+)',
            text,
            re.I
        )
        if row_match:
            row = row_match.group(1)
            cnpj_match = re.search(cnpj_pat, row)
            if cnpj_match:
                cnpj = cnpj_match.group(0)
            ie_match = re.search(r'\b(\d{8,14})\b', row)
            if ie_match:
                ie = ie_match.group(1)

        if not cnpj:
            cnpj_match = re.search(cnpj_pat, text)
            if cnpj_match:
                cnpj = cnpj_match.group(0)

        if not ie:
            ie_match = re.search(r'INSCRI[ÇC][ÃA]O\s+ESTADUAL\s*\n?\s*([\d\.]{8,20})', text, re.I)
            if ie_match:
                ie = re.sub(r'\D', '', ie_match.group(1))
        return cnpj, ie

    def _extract_destinatario_fields(self, text: str, lines: List[str]) -> Dict[str, str]:
        """提取收件人区块，避免被顶部摘要行误触发。"""
        out = {
            'destinatario_nome': '',
            'destinatario_cnpj_cpf': '',
            'destinatario_ie': '',
            'destinatario_endereco': '',
            'data_emissao': '',
            'data_saida': ''
        }

        section_idx = -1
        for i, line in enumerate(lines):
            up = self._normalize_label(line)
            if "DESTINATARIO" in up and "REMETENTE" in up:
                section_idx = i
                break

        if section_idx == -1:
            fallback = re.search(r'DESTINAT[ÁA]RIO:\s*(.+?)(?:\s+-\s+|\n|$)', text, re.I)
            if fallback:
                out['destinatario_nome'] = fallback.group(1).strip()
            return out

        section_end = min(len(lines), section_idx + 24)
        for idx in range(section_idx + 1, min(len(lines), section_idx + 40)):
            up = self._normalize_label(lines[idx])
            if any(tag in up for tag in [
                "INFORMACOES DO LOCAL DE ENTREGA",
                "CALCULO DO IMPOSTO",
                "TRANSPORTADOR / VOLUMES TRANSPORTADOS",
                "TRANSPORTADOR VOLUME",
                "DADOS DOS PRODUTOS SERVICOS",
                "DADOS DO PRODUTO SERVICOS",
            ]):
                section_end = idx
                break

        window = lines[section_idx: section_end]
        for idx, row in enumerate(window):
            up = self._normalize_label(row)

            if "NOME RAZAO SOCIAL" in up and "DATA DA EMISSAO" in up and idx + 1 < len(window):
                detail = window[idx + 1].strip()
                line_match = re.match(r'(.+?)\s+(\d{2,3}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2})\s+(\d{2}/\d{2}/\d{4})$', detail)
                if line_match:
                    if not out['destinatario_nome']:
                        out['destinatario_nome'] = line_match.group(1).strip()
                    out['destinatario_cnpj_cpf'] = line_match.group(2).strip()
                    out['data_emissao'] = line_match.group(3).strip()
                else:
                    id_match = re.search(r'(\d{2,3}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2})', detail)
                    date_match = re.search(r'(\d{2}/\d{2}/\d{4})$', detail)
                    if id_match:
                        out['destinatario_cnpj_cpf'] = id_match.group(1)
                        if not out['destinatario_nome']:
                            out['destinatario_nome'] = detail[:id_match.start()].strip()
                    elif date_match and not out['destinatario_nome']:
                        out['destinatario_nome'] = detail[:date_match.start()].strip()
                    elif not out['destinatario_nome']:
                        out['destinatario_nome'] = detail
                    if date_match:
                        out['data_emissao'] = date_match.group(1)

            if "ENDERECO" in up and "CEP" in up and idx + 1 < len(window):
                out['destinatario_endereco'] = window[idx + 1].strip()
                date_match = re.search(r'(\d{2}/\d{2}/\d{4})$', out['destinatario_endereco'])
                if date_match:
                    out['data_saida'] = date_match.group(1)
                    out['destinatario_endereco'] = out['destinatario_endereco'][:date_match.start()].strip()

            if "MUNICIPIO" in up and "UF" in up and idx + 1 < len(window):
                city_line = re.sub(r'\b\d{2}:\d{2}:\d{2}\b.*$', '', window[idx + 1]).strip()
                city_label = self._normalize_label(city_line)
                if city_line and not any(tag in city_label for tag in ["QUANTIDADE", "TRANSPORTADOR", "CALCULO DO IMPOSTO"]):
                    if out['destinatario_endereco']:
                        out['destinatario_endereco'] += f" {city_line}"
                    else:
                        out['destinatario_endereco'] = city_line

            if "INSCRICAO ESTADUAL" in up and idx + 1 < len(window) and not out['destinatario_ie']:
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

    def _fill_tax_summary_from_table(self, lines: List[str], res: Dict[str, Any]) -> None:
        """从税表数值行兜底提取总额与税值。"""
        for i, line in enumerate(lines):
            upper = self._normalize_label(line)
            if any(tag in upper for tag in ["V TOTAL PRODUTOS", "VALOR TOTAL DOS PRODUTOS"]) and i + 1 < len(lines):
                nums = self._extract_decimal_numbers(lines[i + 1])
                if len(nums) >= 9:
                    if res.get('bc_icms', 0.0) <= 0: res['bc_icms'] = nums[0]
                    if res.get('v_icms', 0.0) <= 0: res['v_icms'] = nums[1]
                    if res.get('bc_icms_st', 0.0) <= 0: res['bc_icms_st'] = nums[2]
                    if res.get('v_icms_st', 0.0) <= 0: res['v_icms_st'] = nums[3]
                    if res.get('v_fcp_uf_dest', 0.0) <= 0: res['v_fcp_uf_dest'] = nums[6]
                    if res.get('v_pis', 0.0) <= 0: res['v_pis'] = nums[7]
                    if res.get('v_prod', 0.0) <= 0: res['v_prod'] = nums[8]
                elif len(nums) >= 5:
                    if res.get('bc_icms', 0.0) <= 0: res['bc_icms'] = nums[0]
                    if res.get('v_icms', 0.0) <= 0: res['v_icms'] = nums[1]
                    if res.get('bc_icms_st', 0.0) <= 0: res['bc_icms_st'] = nums[2]
                    if res.get('v_icms_st', 0.0) <= 0: res['v_icms_st'] = nums[3]
                    if res.get('v_prod', 0.0) <= 0: res['v_prod'] = nums[4]

            if any(tag in upper for tag in ["V TOTAL DA NOTA", "VALOR TOTAL DA NOTA"]) and i + 1 < len(lines):
                nums = self._extract_decimal_numbers(lines[i + 1])
                if nums:
                    if res.get('v_frete', 0.0) <= 0 and len(nums) >= 1: res['v_frete'] = nums[0]
                    if res.get('v_seguro', 0.0) <= 0 and len(nums) >= 2: res['v_seguro'] = nums[1]
                    if res.get('v_desconto', 0.0) <= 0 and len(nums) >= 3: res['v_desconto'] = nums[2]
                    if res.get('v_outras_desp', 0.0) <= 0 and len(nums) >= 4: res['v_outras_desp'] = nums[3]
                    if res.get('v_ipi', 0.0) <= 0 and len(nums) >= 5: res['v_ipi'] = nums[4]
                    if res.get('v_icms_uf_dest', 0.0) <= 0 and len(nums) >= 6: res['v_icms_uf_dest'] = nums[5]
                    if res.get('v_tot_trib', 0.0) <= 0 and len(nums) >= 7: res['v_tot_trib'] = nums[6]
                    if res.get('v_cofins', 0.0) <= 0 and len(nums) >= 8: res['v_cofins'] = nums[7]
                    if res.get('v_nota', 0.0) <= 0: res['v_nota'] = nums[-1]
                break

    def recognize_from_text(self, text: str, file_path: str = "") -> Dict[str, Any]:
        """从OCR文本中识别全量数据"""
        res = {
            'chave_acesso': '', 'numero_nota': '', 'serie': '', 'natureza_operacao': '',
            'protocolo': '', 'data_emissao': '', 'data_saida': '',
            'bc_icms': 0.0, 'v_icms': 0.0, 'bc_icms_st': 0.0, 'v_icms_st': 0.0,
            'v_pis': 0.0, 'v_cofins': 0.0, 'v_ipi': 0.0, 'v_frete': 0.0,
            'v_seguro': 0.0, 'v_desconto': 0.0, 'v_outras_desp': 0.0,
            'v_prod': 0.0, 'v_nota': 0.0, 'v_icms_uf_dest': 0.0, 'v_fcp_uf_dest': 0.0,
            'v_tot_trib': 0.0, 'valor_total': 0.0,
            'emitente_nome': '', 'emitente_cnpj': '', 'emitente_ie': '', 'emitente_endereco': '',
            'destinatario_nome': '', 'destinatario_cnpj_cpf': '', 'destinatario_ie': '', 'destinatario_endereco': '',
            'peso_bruto': 0.0, 'peso_liquido': 0.0, 'inf_complementar': '',
            'items': []
        }

        lines = [line.strip() for line in text.split('\n') if line.strip()]

        # 1. 基础正则抓取（支持多分组）
        for key, pattern in self.patterns.items():
            match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
            if not match:
                continue

            val = self._first_non_empty_group(match)
            if any(k in key for k in ['v_', 'bc_', 'peso_']):
                res[key] = self.clean_number(val)
            elif key == 'chave_acesso':
                res[key] = re.sub(r'\D', '', val)
            else:
                res[key] = val

        # 2. 关键字段兜底：access key / 发行人 / 收件人 / 日期 / 税表
        extracted_key = self._extract_chave_acesso(text, lines)
        if extracted_key:
            res['chave_acesso'] = extracted_key
        elif file_path:
            res['chave_acesso'] = self._extract_chave_acesso_from_filename(file_path)

        extracted_note = self._extract_numero_nota(text, lines)
        if extracted_note:
            res['numero_nota'] = extracted_note

        natureza, protocolo = self._extract_natureza_protocolo(text, lines)
        if natureza:
            res['natureza_operacao'] = natureza
        if protocolo:
            res['protocolo'] = protocolo
        elif str(res.get('protocolo', '')).strip().upper() == 'PROTOCOLO DE AUTORIZAÇÃO DE USO':
            res['protocolo'] = ''

        emitente_nome = self._extract_emitente_nome(text, lines)
        if emitente_nome:
            res['emitente_nome'] = emitente_nome
        if not res.get('emitente_endereco'):
            emitente_addr = self._extract_emitente_endereco(lines)
            if emitente_addr:
                res['emitente_endereco'] = emitente_addr

        emit_cnpj, emit_ie = self._extract_emitente_docs(text)
        if emit_cnpj:
            res['emitente_cnpj'] = emit_cnpj
            res['cnpj_emitente'] = emit_cnpj
        if emit_ie:
            res['emitente_ie'] = emit_ie
            res['ie_emitente'] = emit_ie

        dest_data = self._extract_destinatario_fields(text, lines)
        for k, v in dest_data.items():
            if v and not res.get(k):
                res[k] = v

        if not res['data_emissao']:
            em_match = re.search(r'EMISS[ÃA]O:\s*(\d{2}/\d{2}/\d{4})', text, re.I)
            if em_match:
                res['data_emissao'] = em_match.group(1)

        # 税表兜底（修复 v_nota / v_tot_trib 等）
        self._fill_tax_summary_from_table(lines, res)

        if res.get('v_nota', 0.0) <= 0:
            total_match = re.search(r'VALOR\s+TOTAL\s*:\s*R\$\s*([\d\.,]+)', text, re.I)
            if total_match:
                res['v_nota'] = self.clean_number(total_match.group(1))

        res['valor_total'] = res['v_nota']

        # 4. 补充信息 (INFORMAÇÕES COMPLEMENTARES)
        if "INFORMAÇÕES COMPLEMENTARES" in text:
            inf_part = text.split("INFORMAÇÕES COMPLEMENTARES")[1]
            inf_end_idx = len(inf_part)
            for kw in ["RESERVADO AO FISCO", "DADOS DOS PRODUTOS", "CÁLCULO DO ISSQN"]:
                idx = inf_part.find(kw)
                if idx != -1 and idx < inf_end_idx: inf_end_idx = idx
            res['inf_complementar'] = inf_part[:inf_end_idx].strip()

        # 5. 商品明细行 (稳定识别逻辑)
        start_idx = -1
        for i, line in enumerate(lines):
            if any(kw in line.upper() for kw in ["DADOS DOS PRODUTOS", "PRODUTOS / SERVIÇOS"]):
                start_idx = i; break
        
        if start_idx != -1:
            buffer_desc, pending_code = [], ""
            skip_order_tail = False
            header_terms = [
                "CODIGO", "DESCRICAO", "NCM SH", "O CST", "CST", "CFOP", "UN", "UNID",
                "QUANT", "QTD", "VALOR", "UNIT", "TOTAL", "B CALC", "BCALC", "ALIQ",
                "PRODUTO", "ICMS", "IPI", "ALIQUOTAS"
            ]
            for line in lines[start_idx+1:]:
                line_label = self._normalize_label(line)
                if any(kw in line_label for kw in ["DADOS ADICIONAIS", "RESERVADO", "INFORMACOES COMPLEMENTARES"]): break
                parts = line.split()
                if not parts:
                    continue

                if skip_order_tail:
                    if ")" in line:
                        skip_order_tail = False
                    continue

                if any(term in line_label for term in ["CODIGO PRODUTO", "VALOR TOTAL"]):
                    continue
                if self._is_item_metadata_line(line, line_label):
                    if "PEDIDO" in re.sub(r'\s+', '', line_label) and ")" not in line:
                        skip_order_tail = True
                    continue

                ncm_idx, cfop_idx, u_idx = self._find_item_markers(parts)
                if u_idx != -1:
                    desc = " ".join(buffer_desc).strip()
                    ncm_idx = next((idx for idx, p in enumerate(parts[:u_idx + 1]) if len(p) >= 8 and p[:8].isdigit()), -1)
                    ncm = parts[ncm_idx][:8] if ncm_idx >= 0 else ""
                    
                    # 确定代码：排除 NCM 后的第一个合适字符串
                    code = pending_code
                    if not code:
                        for p in parts[:u_idx]:
                            if len(p) > 2 and p != ncm and (any(c.isdigit() for c in p) or '-' in p):
                                code = p; break

                    if not desc:
                        desc_parts = []
                        if ncm_idx > 0:
                            desc_parts = parts[1:ncm_idx]
                        elif u_idx > 1:
                            desc_parts = parts[1:u_idx]

                        merged = parts[0] if parts else ""
                        merged_match = re.match(r'^(.+?)([A-ZÁ-Ý][a-zá-ÿ].*)$', merged)
                        if merged_match:
                            if not code or code == merged:
                                code = merged_match.group(1).strip()
                            desc_parts = [merged_match.group(2).strip()] + desc_parts
                        elif not code and merged:
                            code = merged

                        desc = " ".join(x for x in desc_parts if x).strip()
                        if not desc:
                            desc = line.strip()
                    
                    item = {
                        'codigo': code, 'descricao': desc, 'ncm': ncm, 'unidade': parts[u_idx] if u_idx < len(parts) and any(c.isalpha() for c in parts[u_idx]) else "un",
                        'qtd': 0.0, 'valor_unit': 0.0, 'valor_total': 0.0, 'bc_icms': 0.0, 'v_icms': 0.0
                    }
                    
                    # 数值列优先从单位之后的整段尾串提取，兼容整数数量与“1,00001.699,8900”这类黏连 token。
                    tail_start = u_idx + 1 if u_idx < len(parts) else len(parts)
                    tail_text = " ".join(parts[tail_start:])
                    raw_nums = re.findall(r'\d+(?:\.\d{3})*,\d{2,4}|\b\d+\b', tail_text)
                    nums = [self.clean_number(num) for num in raw_nums]
                    
                    if len(nums) >= 3:
                        item['qtd'], item['valor_unit'], item['valor_total'] = nums[0], nums[1], nums[2]
                        remainder = nums[3:]
                        if len(remainder) >= 2:
                            if len(remainder) >= 3 and remainder[0] == 0.0 and remainder[1] > 0:
                                item['bc_icms'], item['v_icms'] = remainder[1], remainder[2]
                            else:
                                item['bc_icms'], item['v_icms'] = remainder[0], remainder[1]
                    
                    res['items'].append(item)
                    buffer_desc, pending_code = [], ""
                else:
                    if pending_code or buffer_desc:
                        buffer_desc.append(line)
                    elif res.get('items') and not self._looks_like_product_code(parts[0]) and not any(
                        term in self._normalize_label(parts[0]) for term in header_terms
                    ):
                        res['items'][-1]['descricao'] = self._merge_description(res['items'][-1].get('descricao', ''), line)
                    elif not pending_code and self._looks_like_product_code(parts[0]):
                        pending_code = parts[0]
                        remainder = " ".join(parts[1:]).strip()
                        if remainder:
                            buffer_desc.append(remainder)
                    elif not any(term in self._normalize_label(parts[0]) for term in header_terms):
                        buffer_desc.append(line)
        return res

    def to_comprehensive_dataframe(self, results: List[Dict[str, Any]]) -> pd.DataFrame:
        """转换为包含所有请求字段的详尽DataFrame"""
        rows = []
        for res in results:
            base = {
                '文件路径': res.get('file_path', ''), 'Access Key (Chave)': res.get('chave_acesso', ''),
                '发票号码': res.get('numero_nota', ''), '系列 (Série)': res.get('serie', ''),
                '业务性质': res.get('natureza_operacao', ''), '日期': res.get('data_emissao', ''),
                '出库日期': res.get('data_saida', ''), '发行人': res.get('emitente_nome', ''),
                '发行人CNPJ': res.get('emitente_cnpj', ''), '发行人IE': res.get('emitente_ie', ''),
                '发行人地址': res.get('emitente_endereco', ''), '收件人': res.get('destinatario_nome', ''),
                '收件人ID (CNPJ/CPF)': res.get('destinatario_cnpj_cpf', ''), '收件人IE': res.get('destinatario_ie', ''),
                '收件人地址': res.get('destinatario_endereco', ''), 'ICMS底数': res.get('bc_icms', 0.0),
                'ICMS金额': res.get('v_icms', 0.0), 'ICMS ST底数': res.get('bc_icms_st', 0.0),
                'ICMS ST金额': res.get('v_icms_st', 0.0), 'PIS金额': res.get('v_pis', 0.0),
                'COFINS金额': res.get('v_cofins', 0.0), 'IPI金额': res.get('v_ipi', 0.0),
                '运费': res.get('v_frete', 0.0), '折扣': res.get('v_desconto', 0.0),
                '其他费用': res.get('v_outras_desp', 0.0), '商品总计': res.get('v_prod', 0.0),
                '发票总额': res.get('v_nota', 0.0), 'ICMS UF Dest金额': res.get('v_icms_uf_dest', 0.0),
                'FCP UF Dest金额': res.get('v_fcp_uf_dest', 0.0), '总税贡献 (Trib)': res.get('v_tot_trib', 0.0),
                '毛重 (Peso Bruto)': res.get('peso_bruto', 0.0), '净重 (Peso Líquido)': res.get('peso_liquido', 0.0),
                '补充信息': res.get('inf_complementar', '')
            }
            if res.get('items'):
                for item in res['items']:
                    row = base.copy()
                    row.update({
                        '商品代码': item.get('codigo', ''), '商品描述': item.get('descricao', ''),
                        'NCM': item.get('ncm', ''), '单位': item.get('unidade', ''),
                        '数量': item.get('qtd', 0.0), '单价': item.get('valor_unit', 0.0),
                        '商品总价': item.get('valor_total', 0.0), '项目ICMS金额': item.get('v_icms', 0.0)
                    })
                    rows.append(row)
            else:
                row = base.copy()
                row.update({'商品描述': '无明细项'})
                rows.append(row)
        return pd.DataFrame(rows)

    def to_standard_voucher(self, results: List[Dict[str, Any]]) -> pd.DataFrame:
        """标准凭证格式导出"""
        rows = []
        for res in results:
            d = res['data_emissao']
            date_str = f"{d[6:]}{d[3:5]}{d[:2]}" if len(d) == 10 else d
            numero_nota = str(res.get('numero_nota', ''))
            nota_digits = re.sub(r'\D', '', numero_nota)
            seq = nota_digits[-4:] if nota_digits else numero_nota[-4:]
            summary = f"NF-e{numero_nota} | {res['natureza_operacao'][:20]} | {res['emitente_nome']}"
            main_row = {
                '凭证日期': date_str, '序号': seq, '会计凭证No.': numero_nota,
                '摘要': summary, '类型': '3', '科目编码': '', '往来单位编码': res['emitente_cnpj'],
                '往来单位名': res['emitente_nome'], '金额': res['v_nota'], '外币金额': 0.0, '汇率': 1.0,
                '部门': '', '备注/附加信息': f"Key:{res['chave_acesso'][-4:]}; ICMS:{res['v_icms']}"
            }
            rows.append(main_row)
            for item in res.get('items', []):
                item_desc = (item.get('descricao') or item.get('codigo') or '').strip()
                rows.append({
                    '凭证日期': date_str, '序号': seq, '会计凭证No.': numero_nota,
                    '摘要': f"[Item] {item_desc[:100]}", '类型': '3', '科目编码': item['codigo'],
                    '往来单位编码': res['emitente_cnpj'], '往来单位名': res['emitente_nome'],
                    '金额': item['valor_total'], '外币金额': item['qtd'], '汇率': item['valor_unit'],
                    '部门': item['unidade'], '备注/附加信息': f"NCM:{item['ncm']}"
                })
        return pd.DataFrame(rows)

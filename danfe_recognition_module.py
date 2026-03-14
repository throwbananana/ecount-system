# -*- coding: utf-8 -*-
"""
文档识别模块 - 巴西NF-e (DANFE) 识别 (全字段增强版)
支持提取发票抬头、发行人、收件人、税金汇总、物流及商品明细等全量字段。
"""

import re
import pandas as pd
from typing import Dict, List, Any, Tuple

class DanfeRecognizer:
    """巴西NF-e (DANFE) 文档识别器 - 能够识别发票所有维度的数据"""

    def __init__(self):
        # 1. 核心字段正则表达式模式 (支持多种变体和换行)
        self.patterns = {
            'chave_acesso': r'CHAVE DE ACESSO\s*([\d\s]{44,80})',
            'numero_nota': r'\bN[º°o]?\s*\.?\s*([\d][\d\.]+)\b',
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

    def _extract_chave_acesso(self, text: str, lines: List[str]) -> str:
        """提取 44 位 access key。"""
        for i, line in enumerate(lines):
            if "CHAVE DE ACESSO" in line.upper():
                block = " ".join(lines[i:i+4])
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
        return ""

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
                    if any(tag in upper for tag in ["CHAVE DE ACESSO", "Nº.", "SÉRIE", "CNPJ", "INSCRIÇÃO"]):
                        break
                    if any(tag in upper for tag in ["DANFE", "DOCUMENTO AUXILIAR", "FISCAL ELETRÔNICA"]):
                        continue
                    if re.search(r'\d{5}-\d{3}', cand) or any(x in upper for x in ["VILA", "RUA", "AV.", "AVENIDA", "SÃO", "SAO"]):
                        addr_parts.append(cand.strip())
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
            up = line.upper()
            if "DESTINAT" in up and "REMETENTE" in up:
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

            if "NOME / RAZ" in up and "CNPJ / CPF" in up and "DATA DA EMISS" in up and idx + 1 < len(window):
                detail = window[idx + 1].strip()
                line_match = re.match(
                    r'(.+?)\s+(\d{2,3}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2})\s+(\d{2}/\d{2}/\d{4})$',
                    detail
                )
                if line_match and not out['destinatario_nome']:
                    out['destinatario_nome'] = line_match.group(1).strip()
                    out['destinatario_cnpj_cpf'] = line_match.group(2).strip()
                    out['data_emissao'] = line_match.group(3).strip()
                else:
                    id_match = re.search(r'(\d{2,3}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2})', detail)
                    if id_match and not out['destinatario_nome']:
                        out['destinatario_cnpj_cpf'] = id_match.group(1)
                        out['destinatario_nome'] = detail[:id_match.start()].strip()
                    elif not out['destinatario_nome']:
                        out['destinatario_nome'] = detail

            if "ENDERE" in up and "CEP" in up and idx + 1 < len(window):
                out['destinatario_endereco'] = window[idx + 1].strip()
                date_match = re.search(r'(\d{2}/\d{2}/\d{4})$', out['destinatario_endereco'])
                if date_match:
                    out['data_saida'] = date_match.group(1)
                    out['destinatario_endereco'] = out['destinatario_endereco'][:date_match.start()].strip()

            if "MUNIC" in up and "UF" in up and idx + 1 < len(window):
                city_line = re.sub(r'\b\d{2}:\d{2}:\d{2}\b.*$', '', window[idx + 1]).strip()
                if city_line:
                    if out['destinatario_endereco']:
                        out['destinatario_endereco'] += f" {city_line}"
                    else:
                        out['destinatario_endereco'] = city_line

            if "INSCRIÇÃO ESTADUAL" in up and idx + 1 < len(window) and not out['destinatario_ie']:
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
            upper = line.upper()
            if "V. TOTAL PRODUTOS" in upper and i + 1 < len(lines):
                nums = self._extract_decimal_numbers(lines[i + 1])
                if len(nums) >= 9:
                    if res.get('bc_icms', 0.0) <= 0: res['bc_icms'] = nums[0]
                    if res.get('v_icms', 0.0) <= 0: res['v_icms'] = nums[1]
                    if res.get('bc_icms_st', 0.0) <= 0: res['bc_icms_st'] = nums[2]
                    if res.get('v_icms_st', 0.0) <= 0: res['v_icms_st'] = nums[3]
                    if res.get('v_fcp_uf_dest', 0.0) <= 0: res['v_fcp_uf_dest'] = nums[6]
                    if res.get('v_pis', 0.0) <= 0: res['v_pis'] = nums[7]
                    if res.get('v_prod', 0.0) <= 0: res['v_prod'] = nums[8]

            if "V. TOTAL DA NOTA" in upper and i + 1 < len(lines):
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

    def recognize_from_text(self, text: str) -> Dict[str, Any]:
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
        if not res['chave_acesso'] or len(res['chave_acesso']) < 44:
            res['chave_acesso'] = self._extract_chave_acesso(text, lines)

        if not res['numero_nota'] or res['numero_nota'].strip('.') == '':
            note_match = re.search(r'\bN[º°o]?\s*\.?\s*([\d][\d\.]+)\b', text, re.I)
            if note_match:
                res['numero_nota'] = note_match.group(1).strip()

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
            header_terms = ["CÓDIGO", "DESCRIÇÃO", "NCM/SH", "O/CST", "CFOP", "UN", "QUANT", "VALOR", "UNIT", "TOTAL", "B.CÁLC", "ALÍQ"]
            for line in lines[start_idx+1:]:
                if any(kw in line.upper() for kw in ["DADOS ADICIONAIS", "RESERVADO", "INFORMAÇÕES COMPLEMENTARES"]): break
                parts = line.split()
                if not parts or any(term in line.upper() for term in [
                    "CÓDIGO PRODUTO", "VALOR TOTAL", "PFCPUFDEST", "PICMSUFDEST",
                    "VICMSUFDEST", "VICMSUFREMET", "(PEDIDO"
                ]):
                    continue
                
                u_idx = -1
                for idx, p in enumerate(parts):
                    p_clean = re.sub(r'[\d\.,]', '', p).lower().strip()
                    if p_clean in self.units and not any(c.isdigit() for c in p):
                        if idx + 1 < len(parts) and self.clean_number(parts[idx+1]) >= 0:
                            u_idx = idx; break
                
                # 如果没找到单位，尝试通过 NCM (8位) + CFOP (4位) 定位
                ncm_idx, cfop_idx = -1, -1
                if u_idx == -1:
                    for idx, p in enumerate(parts):
                        if len(p) == 8 and p.isdigit(): ncm_idx = idx
                        if len(p) == 4 and p.isdigit() and ncm_idx != -1: 
                            cfop_idx = idx; break
                    if cfop_idx != -1 and cfop_idx + 1 < len(parts):
                        # 假设 CFOP 后通常跟着单位或直接是数量
                        # 如果下一位是字母，则是单位；如果是数字，则是数量
                        if any(c.isalpha() for c in parts[cfop_idx+1]):
                            u_idx = cfop_idx + 1
                        else:
                            # 插入一个虚拟单位索引以利用后续逻辑
                            u_idx = cfop_idx 

                if u_idx != -1:
                    desc = " ".join(buffer_desc).strip()
                    ncm_idx = next((idx for idx, p in enumerate(parts[:u_idx+1]) if len(p) >= 8 and p[:8].isdigit()), -1)
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
                    
                    # 终极数值序列提取：过滤掉 NCM(8位) 和 CFOP(4位)
                    nums = []
                    for p in parts:
                        # 只看可能是数值的部分
                        if any(c.isdigit() for c in p):
                            val = self.clean_number(p)
                            p_clean = p.replace('.', '').replace(',', '')
                            # 过滤条件：排除 NCM (8位) 和 CFOP (4位) 且该位确实是纯数字的情况
                            if p_clean.isdigit():
                                if len(p_clean) == 8 or len(p_clean) == 4:
                                    continue
                            
                            # 财务数值通常在单位/CFOP之后
                            if val > 0 or "0,00" in p or p == "0":
                                # 确保这个数值在单位索引之后
                                try:
                                    p_idx = parts.index(p)
                                    if p_idx > u_idx or (u_idx == cfop_idx and p_idx > cfop_idx):
                                        nums.append(val)
                                except: pass
                    
                    if len(nums) >= 3:
                        item['qtd'], item['valor_unit'], item['valor_total'] = nums[0], nums[1], nums[2]
                        if len(nums) >= 5: item['bc_icms'], item['v_icms'] = nums[3], nums[4]
                    
                    res['items'].append(item)
                    buffer_desc, pending_code = [], ""
                else:
                    if not pending_code and len(parts[0]) > 3 and (any(c.isdigit() for c in parts[0]) or '-' in parts[0]):
                        pending_code = parts[0]
                        buffer_desc.append(" ".join(parts[1:]))
                    elif not any(term in parts[0].upper() for term in header_terms):
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

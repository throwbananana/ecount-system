# -*- coding: utf-8 -*-
"""
摘要智能识别引擎
根据摘要内容自动识别并填充凭证字段
"""

import os
import re
import json
import pandas as pd
from typing import Dict, Optional, Any
from base_data_manager import BaseDataManager

try:
    from zhipuai import ZhipuAI
    HAS_ZHIPUAI = True
except ImportError:
    HAS_ZHIPUAI = False

DEFAULT_API_KEY = os.environ.get("ZHIPU_API_KEY", "")
# 限制 AI 提示中带入的科目数量，避免上下文过长导致性能问题
MAX_AI_SUBJECTS = 300


class SummaryIntelligence:
    """摘要智能识别器"""

    def __init__(self, base_data_mgr: Optional[BaseDataManager] = None, default_values: Optional[Dict] = None, api_key: str = DEFAULT_API_KEY):
        """初始化

        Args:
            base_data_mgr: 基础数据管理器
            default_values: 默认值字典
            api_key: 智谱AI API Key (兼容旧逻辑，优先使用 default_values 中的配置)
        """
        self.base_data_mgr = base_data_mgr or BaseDataManager()
        self.default_values = default_values or {}
        self.api_key = self.default_values.get("ai_api_key") or api_key  # 优先从配置取，空值回退环境变量
        self.ai_provider = self.default_values.get("ai_provider", "zhipu") # zhipu / lm_studio
        self.ai_base_url = self.default_values.get("ai_base_url", "http://localhost:1234/v1")
        self.ai_model_name = self.default_values.get("ai_model_name", "local-model")
        self.recognition_priority = self.default_values.get("recognition_priority", "default") # default / profit_loss / balance_sheet
        
        self.ai_client = None
        
        # 特殊科目缓存
        self.special_accounts = {
            "销售费用": None,
            "管理费用": None,
            "财务费用": None
        }
        
        self._init_ai_client()

        self._init_recognition_rules()
        self._load_base_data_cache()

    def _post_adjust_for_priority(self, summary: str, code: str) -> str:
        """根据优先级策略对AI/缓存科目做纠偏"""
        if not code:
            return code
        if self.recognition_priority == "profit_loss":
            try:
                if str(code).startswith(("1", "2")):
                    expense_keywords = ["报销", "费用", "租金", "油费", "路费", "机票", "差旅", "住宿", "补贴", "薪", "工资", "福利", "补发", "罚款", "运费"]
                    if any(kw in summary for kw in expense_keywords):
                        replacement = self.special_accounts.get("管理费用") or self.special_accounts.get("销售费用") or self.special_accounts.get("财务费用")
                        if replacement:
                            return replacement
            except Exception:
                return code
        return code

    def _get_expense_parent_limit(self) -> str:
        """获取费用科目上级限制：6601 / 6602 / 6603 / UNLIMITED。"""
        raw = str(self.default_values.get("费用科目上级限制", "") or "").strip().upper()
        if raw in ("6601", "6602", "6603"):
            return raw
        return "UNLIMITED"

    def _apply_expense_parent_constraint(self, account_code: Optional[str]) -> Optional[str]:
        """按配置约束费用科目归属到指定上级科目。"""
        code = str(account_code or "").strip()
        if not code:
            return account_code

        limit = self._get_expense_parent_limit()
        if limit == "UNLIMITED":
            return code
        if not code.startswith("6"):
            return code
        if code.startswith(limit):
            return code

        if limit == "6601":
            target_name = "销售费用"
        elif limit == "6602":
            target_name = "管理费用"
        else:
            target_name = "财务费用"
        target_code = str(self.special_accounts.get(target_name) or "").strip()
        if target_code:
            return target_code
        return limit

    def _init_ai_client(self):
        """初始化 AI 客户端"""
        print(f"初始化 AI 客户端: Provider={self.ai_provider}, URL={self.ai_base_url}, Model={self.ai_model_name}, Priority={self.recognition_priority}")
        
        if self.ai_provider == "lm_studio":
            try:
                # 尝试导入标准 OpenAI 库
                from openai import OpenAI
                # LM Studio 不需要真实的 API Key，但 SDK 必须填一个
                effective_key = self.api_key if self.api_key else "lm-studio"
                self.ai_client = OpenAI(base_url=self.ai_base_url, api_key=effective_key)
                print("[OK] LM Studio (OpenAI) 客户端初始化成功")
            except ImportError:
                print("[WARN] 未检测到 openai 库，LM Studio 本地功能不可用。请运行 'pip install openai'")
            except Exception as e:
                print(f"[ERR] AI 客户端初始化失败: {e}")
        else:
            # 默认智谱
            if not self.api_key:
                print("[WARN] 未配置 ZHIPU_API_KEY，已禁用远程摘要 AI。请在环境变量或设置中提供密钥。")
                self.ai_client = None
                return
            if HAS_ZHIPUAI:
                try:
                    self.zhipu_client = ZhipuAI(api_key=self.api_key)
                    self.ai_client = self.zhipu_client # 兼容，后续统一用 ai_client 封装
                    print("[OK] ZhipuAI 客户端初始化成功")
                except Exception as e:
                    print(f"[ERR] AI 客户端初始化失败: {e}")
            else:
                print("[WARN] 未检测到 zhipuai 库，智谱 AI 功能将不可用。请运行 'pip install zhipuai'")

    def update_config(self, provider=None, api_key=None, base_url=None, model_name=None, recognition_priority=None):
        """更新 AI 配置"""
        if provider: self.ai_provider = provider
        if api_key: self.api_key = api_key
        if base_url: self.ai_base_url = base_url
        if model_name: self.ai_model_name = model_name
        if recognition_priority: self.recognition_priority = recognition_priority
        self._init_ai_client()

    def update_api_key(self, api_key: str):
        """兼容旧接口"""
        self.update_config(api_key=api_key)

    def _init_recognition_rules(self):
        """初始化识别规则（从数据库加载）"""
        
        self.business_type_rules = {}
        self.account_keyword_rules = []
        self.dept_keyword_rules = {}

        # 尝试从数据库加载规则
        if self.base_data_mgr:
            try:
                all_rules = self.base_data_mgr.get_recognition_rules()
                
                for rule in all_rules:
                    try:
                        keywords = json.loads(rule["keywords"]) if rule["keywords"] else []
                    except Exception:
                        keywords = []

                    normalized_keywords = []
                    for kw in keywords:
                        if kw is None:
                            continue
                        kw_str = str(kw).strip()
                        if kw_str:
                            normalized_keywords.append(kw_str.lower())

                    if rule["rule_type"] == "business":
                        # 构建业务类型规则
                        self.business_type_rules[rule["name"]] = {
                            "keywords": normalized_keywords,
                            "account": rule["account_code"],
                            "type": rule["transaction_type"],
                            "summary_code": rule["summary_code"]
                        }
                    
                    elif rule["rule_type"] == "account":
                        # 构建科目关键词规则 (格式: (关键词列表, 目标科目代码))
                        self.account_keyword_rules.append((normalized_keywords, rule["account_code"]))
                    
                    elif rule["rule_type"] == "department":
                        # 构建部门关键词规则
                        for kw in normalized_keywords:
                            self.dept_keyword_rules[kw] = rule["dept_code"]
                            
            except Exception as e:
                print(f"从数据库加载识别规则失败: {e}")

        # 金额关键词识别
        self.amount_keywords = {
            "元": 1,
            "美元": 1,
            "USD": 1,
            "千": 1000,
            "万": 10000,
        }

    def _load_base_data_cache(self):
        """加载基础数据到缓存"""
        try:
            def _parse_match_items(raw):
                if not raw:
                    return []
                if isinstance(raw, list):
                    items = raw
                else:
                    try:
                        items = json.loads(raw)
                    except Exception:
                        items = []
                cleaned = []
                for i in items:
                    if i is None:
                        continue
                    s = str(i).strip()
                    if s:
                        cleaned.append(s)
                return cleaned

            # 加载往来单位
            self.partners = {}
            partners = self.base_data_mgr.query("business_partner")
            for p in partners:
                if p.get("code") and p.get("name"):
                    self.partners[p["name"]] = p["code"]
                    # 添加简化名称匹配
                    simplified = re.sub(r'[（\(].*?[）\)]', '', p["name"]).strip()
                    if simplified != p["name"]:
                        self.partners[simplified] = p["code"]
                    # 映射匹配项 -> 代码
                    for alias in _parse_match_items(p.get("match_items")):
                        self.partners[alias] = p["code"]
            self.partner_names_sorted = sorted(self.partners.keys(), key=len, reverse=True)

            # 加载科目编码
            self.accounts = {}
            accounts = self.base_data_mgr.query("account_subject")
            for a in accounts:
                code_name = a.get("code_name", "").strip()
                if not code_name:
                    continue
                
                # 尝试解析 "[1001] 现金" 或 "1001 现金" 或 "1001-现金" 格式
                # 匹配开头是数字，可能被 [] 或 () 包裹
                match = re.match(r'^[\[\(]?(\d+)[\]\)]?[\s\-]*(.*)', code_name)
                if match:
                    code = match.group(1)
                    name = match.group(2).strip()
                    
                    # 映射：名称 -> 代码
                    if name:
                        self.accounts[name] = code
                    
                    # 映射：代码 -> 代码
                    self.accounts[code] = code
                    
                    # 映射：完整字符串 -> 代码 (以防摘要里写的是全称)
                    self.accounts[code_name] = code

                    # 映射：额外匹配项 -> 代码
                    for alias in _parse_match_items(a.get("match_items")):
                        self.accounts[alias] = code
                    
                    # 识别特殊科目
                    if "销售费用" in name:
                        self.special_accounts["销售费用"] = code
                    elif "管理费用" in name:
                        self.special_accounts["管理费用"] = code
                    elif "财务费用" in name:
                        self.special_accounts["财务费用"] = code

                else:
                    # 如果无法解析出代码，尝试直接用整个字符串作为名称（虽然可能无法获得代码，但保持兼容）
                    pass

            # 加载部门
            self.departments = {}
            departments = self.base_data_mgr.query("department")
            for d in departments:
                if d.get("code") and d.get("name"):
                    self.departments[d["name"]] = d["code"]
            self.departments_lower = {
                str(name).lower(): code
                for name, code in self.departments.items()
                if name
            }

        except Exception as e:
            print(f"加载基础数据缓存失败: {e}")
            self.partners = {}
            self.accounts = {}
            self.departments = {}

    def refresh_cache(self):
        """刷新基础数据缓存（删除基础数据后需要调用此方法）"""
        print("正在刷新基础数据缓存...")
        old_counts = {
            "partners": len(self.partners),
            "accounts": len(self.accounts),
            "departments": len(self.departments)
        }

        self._load_base_data_cache()

        new_counts = {
            "partners": len(self.partners),
            "accounts": len(self.accounts),
            "departments": len(self.departments)
        }

        print(f"缓存刷新完成:")
        print(f"  - 往来单位: {old_counts['partners']} -> {new_counts['partners']}")
        print(f"  - 科目编码: {old_counts['accounts']} -> {new_counts['accounts']}")
        print(f"  - 部门: {old_counts['departments']} -> {new_counts['departments']}")

        return {
            "old": old_counts,
            "new": new_counts
        }

    def recognize(self, summary: str, original_data: Optional[Dict] = None, use_ai: bool = False, use_foreign_currency: bool = False) -> Dict[str, Any]:
        """
        智能识别摘要并返回字段映射

        Args:
            summary: 摘要内容
            original_data: 原始数据行
            use_ai: 是否启用AI深度识别
            use_foreign_currency: 是否启用了外币模式（界面选项）

        Returns:
            字段映射字典
        """
        result = {}

        # ... (保留原有逻辑) ...

        # 从original_data中提取其他字段
        if original_data:
            # 从日期字段识别
            date_value = self._extract_from_original(
                original_data,
                ["日期", "凭证日期", "date", "Date", "Fecha", "fecha", "Fecha_DT", "fecha_dt", "FechaDT"]
            )
            if date_value:
                recognized_date = self._recognize_date(str(date_value))
                if recognized_date:
                    result["凭证日期"] = recognized_date

            # 从金额字段识别
            amount_value = self._extract_from_original(original_data, ["金额", "amount", "Amount", "amt"])
            if amount_value:
                recognized_amount = self._recognize_amount(str(amount_value))
                if recognized_amount:
                    result["金额"] = recognized_amount

            # 从汇率字段识别
            rate_value = self._extract_from_original(original_data, ["汇率", "rate", "Rate", "exchange_rate"])
            if rate_value:
                recognized_rate = self._recognize_exchange_rate(str(rate_value))
                if recognized_rate:
                    result["汇率"] = recognized_rate

            # 从外币金额字段识别
            foreign_amount = self._extract_from_original(original_data, ["外币金额", "外币", "foreign_amount"])
            if foreign_amount:
                recognized_foreign = self._recognize_amount(str(foreign_amount))
                if recognized_foreign:
                    result["外币金额"] = recognized_foreign

        # 从摘要识别（如果摘要存在）
        if summary and isinstance(summary, str):
            # 1. 识别业务类型
            business_type = self._recognize_business_type(summary)
            if business_type:
                result.update(business_type)

            # 2. 识别往来单位
            partner = self._recognize_partner(summary)
            if partner:
                result["往来单位编码"] = partner["code"]
                result["往来单位名"] = partner["name"]
                # 如果识别到往来单位，不再强制默认应收账款 1122
                pass

            # 3. 识别科目 (优先规则匹配，失败则尝试AI)
            account = self._recognize_account(summary)
            
            # 如果规则匹配失败，且启用了AI，尝试AI识别
            if not account and use_ai and self.ai_client:
                try:
                    account = self._recognize_account_with_ai(summary)
                except Exception as e:
                    print(f"AI识别失败: {e}")

            # --- 费用科目智能调整 (基于币种) ---
            # 逻辑：如果是费用(6开头)，且非财务费用 -> 外币归销售费用，本币归管理费用
            if account and str(account).startswith("6"):
                fin_exp = self.special_accounts.get("财务费用")
                
                # 如果不是财务费用（财务费用不调整）
                if str(account) != str(fin_exp):
                    is_foreign = use_foreign_currency or ("外币金额" in result)
                    
                    target_acc = None
                    if is_foreign:
                        target_acc = self.special_accounts.get("销售费用")
                    else:
                        target_acc = self.special_accounts.get("管理费用")
                    
                    # 如果找到了对应的目标费用科目，则进行替换
                    # 注意：只有当原科目是模糊识别（如AI识别出的"办公费"）且我们想强制归类时才替换
                    # 但根据需求描述 "如果是外币，则费用引用销售费用"，似乎是强制性的
                    if target_acc:
                        account = target_acc
            # ------------------------------------

            account = self._apply_expense_parent_constraint(account)

            if account and "科目编码" not in result:
                result["科目编码"] = account

            # 4. 识别部门
            department = self._recognize_department(summary)
            if department:
                result["部门"] = department

            # 5. 从摘要中识别金额（如果之前未识别到）
            if "金额" not in result:
                amount = self._recognize_amount(summary)
                if amount:
                    result["金额"] = amount

            # 6. 从摘要中识别日期（如果之前未识别到）
            if "凭证日期" not in result:
                date = self._recognize_date(summary)
                if date:
                    result["凭证日期"] = date

            # 7. 保留原始摘要
            result["摘要"] = summary[:200]  # 限制200字

            # 8. 现金业务强制覆盖（高优先级，覆盖缓存/规则/AI）
            # 需求：现金存BAC/现金取BAC等现金相关业务，主科目应为1001，往来单位编码应为100102。
            cash_override = self._recognize_cash_business_override(summary)
            if cash_override:
                result.update(cash_override)
            
            # 应用默认值（最低优先级，在最后补缺）
            for key, value in self.default_values.items():
                if value and key not in result:  # 只在键不存在时填充默认值
                    result[key] = value

        return result

    def _recognize_cash_business_override(self, summary: str) -> Dict[str, str]:
        """识别现金相关业务并返回强制覆盖字段。"""
        if not summary:
            return {}

        s = str(summary).lower().replace(" ", "")

        # 直接命中：现金存/取 BAC
        direct_hits = (
            "现金存bac",
            "现金取bac",
            "存现金bac",
            "取现金bac",
        )
        if any(k in s for k in direct_hits):
            return {"科目编码": "1001", "往来单位编码": "100102"}

        # 泛化命中：现金 + 存/取 + BAC/ST（覆盖“涉及到现金的相关业务”）
        has_cash = ("现金" in s) or ("cash" in s)
        has_action = any(k in s for k in ("存", "取", "deposit", "withdraw"))
        has_bank = ("bac" in s) or ("st" in s)
        if has_cash and has_action and has_bank:
            return {"科目编码": "1001", "往来单位编码": "100102"}

        return {}

    def _normalize_summary(self, summary: str) -> str:
        """标准化摘要（去除日期、数字等可变信息，提高缓存命中率）"""
        if not summary:
            return ""
        
        s = summary
        
        # 1. 去除日期格式 (2025-01-01, 2025/01/01, 20250101)
        s = re.sub(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', '', s)
        s = re.sub(r'\d{8}', '', s)
        
        # 2. 去除月份/日期描述 (1月, 10月份, 1日, 15号)
        s = re.sub(r'\d{1,2}\s*[月日号]', '', s)
        s = re.sub(r'\d{1,2}\s*月份', '', s)
        
        # 3. 去除金额/数量相关的数字 (保留少许文本特征，但去掉纯数字)
        # 去除带小数点的数字
        s = re.sub(r'\d+\.\d+', '', s)
        # 去除纯数字 (连续2位以上，避免误删型号中的单个数字，视情况调整)
        s = re.sub(r'\d{2,}', '', s)
        
        # 4. 去除常见无意义符号
        s = re.sub(r'\bno\.?\b|[.#-]', ' ', s, flags=re.IGNORECASE)
        
        # 5. 去除多余空格
        s = re.sub(r'\s+', ' ', s).strip()
        
        return s

    def _recognize_account_with_ai(self, summary: str) -> Optional[str]:
        """调用AI识别科目 (含本地缓存)"""
        if not self.accounts:
            return None
            
        # 0. 标准化摘要 (去噪)
        normalized_summary = self._normalize_summary(summary)

        # 1. 优先检查本地缓存 (使用标准化后的摘要)
        cached_code = self.base_data_mgr.get_cached_recognition(normalized_summary)
        if cached_code:
            adjusted = self._post_adjust_for_priority(summary, cached_code)
            print(f"✨ [缓存命中] 摘要: '{summary[:10]}...' (标准化: '{normalized_summary}') -> 科目: {adjusted}")
            return adjusted

        # 1.1 模糊缓存匹配（允许摘要轻微差异时命中）
        fuzzy_code = self.base_data_mgr.get_cached_recognition_fuzzy(normalized_summary)
        if fuzzy_code:
            adjusted = self._post_adjust_for_priority(summary, fuzzy_code)
            print(f"✨ [模糊缓存命中] 摘要: '{summary[:10]}...' (标准化: '{normalized_summary}') -> 科目: {adjusted}")
            return adjusted

        # 2. 调用核心 AI 逻辑
        code = self.recognize_account_with_ai_core(summary)

        # 2.1 费用优先纠偏：若优先级为 profit_loss 且识别为资产/负债类，但摘要明显是费用/报销/支出，则尝试改用费用科目
        if code and self.recognition_priority == "profit_loss":
            try:
                if str(code).startswith(("1", "2")):
                    expense_keywords = ["报销", "费用", "租金", "油费", "路费", "机票", "差旅", "住宿", "补贴", "薪", "工资", "福利", "补发", "罚款", "运费"]
                    if any(kw in summary for kw in expense_keywords):
                        # 优先管理费用 -> 销售费用 -> 财务费用
                        replacement = self.special_accounts.get("管理费用") or self.special_accounts.get("销售费用") or self.special_accounts.get("财务费用")
                        if replacement:
                            print(f"⚖️ 费用优先纠偏: {summary[:10]}... 由 {code} 调整为 {replacement}")
                            code = replacement
            except Exception:
                pass

        # 3. 保存到本地缓存 (使用标准化后的摘要)
        if code:
            self.base_data_mgr.save_cached_recognition(normalized_summary, code)
            print(f"🤖 [AI识别] 摘要: '{summary[:10]}...' -> 科目: {code} (已缓存为: '{normalized_summary}')")
            return code
        
        return None

    def recognize_account_with_ai_core(self, summary: str) -> Optional[str]:
        """仅执行 AI 识别逻辑，不查缓存也不存缓存"""
        if not self.accounts or not self.ai_client:
            return None

        # 构建上下文：取所有科目名称和代码
        subject_list = []
        seen = set()
        sorted_keys = sorted(self.accounts.keys(), key=len, reverse=True)
        
        count = 0
        for key in sorted_keys:
            code = self.accounts[key]
            if code in seen: continue
            if key == code: continue
                
            subject_list.append(f"{code}:{key}")
            seen.add(code)
            count += 1
            if count >= MAX_AI_SUBJECTS: break  # 控制上下文长度，避免请求过大
        
        context_str = "\n".join(subject_list)

        # 构建优先级提示
        priority_hint = ""
        if self.recognition_priority == "profit_loss":
            priority_hint = "7. 【优先规则】如果存在多个合理的科目选择，请优先选择“损益类”科目（如费用、收入），而非资产负债类科目（如应付、预提）。例如：付工资优先选销售费用/管理费用，而非应付职工薪酬。"
        elif self.recognition_priority == "balance_sheet":
            priority_hint = "7. 【优先规则】如果存在多个合理的科目选择，请优先选择“资产负债类”科目（如应付、预提），而非损益类科目。例如：付工资优先选应付职工薪酬，而非费用。"
        expense_parent_limit = self._get_expense_parent_limit()
        expense_parent_hint = ""
        if expense_parent_limit in ("6601", "6602", "6603"):
            expense_parent_hint = (
                f"8. 【费用上级限制】若判断为费用类科目（6开头），请仅选择上级为 {expense_parent_limit} 的费用科目。"
            )

        prompt = f"""
你是一个资深会计。请根据以下“交易摘要”和“可选科目表”，判断该交易最应该计入哪个会计科目。

交易摘要：{summary}

可选科目表（格式为 代码:名称）：
{context_str}

要求：
1. 分析摘要的语义。
2. 从可选科目表中选择一个最匹配的科目。
3. 【重要】优先选择最具体的“子科目”（末级科目），而不是上级科目。例如，如果有“办公费”和“管理费用”，应优先选“办公费”。
4. 只返回该科目的【数字编码】。
5. 如果没有匹配的，返回"None"。
6. 不要解释，只返回代码。
{priority_hint}
{expense_parent_hint}
"""
        try:
            # 统一调用接口
            if self.ai_provider == "lm_studio":
                # OpenAI / LM Studio 调用方式
                try:
                    response = self.ai_client.chat.completions.create(
                        model=self.ai_model_name,
                        messages=[
                            {"role": "system", "content": "You are a helpful accounting assistant."},
                            {"role": "user", "content": prompt}
                        ],
                        temperature=0.1,
                        timeout=30,
                    )
                except TypeError:
                    # 兼容旧版 SDK 不支持 timeout 参数的情况
                    response = self.ai_client.chat.completions.create(
                        model=self.ai_model_name,
                        messages=[
                            {"role": "system", "content": "You are a helpful accounting assistant."},
                            {"role": "user", "content": prompt}
                        ],
                        temperature=0.1,
                    )
                result_text = response.choices[0].message.content.strip()
            else:
                # 智谱调用方式 (假设 zhipu_client.chat.completions.create 接口与 openai 类似，但可能有些许不同)
                # 智谱SDK新版已经兼容 OpenAI 格式，但为了稳妥沿用原逻辑
                model_name = "glm-4-flash"
                if hasattr(self.ai_client, "chat") and hasattr(self.ai_client.chat, "completions"):
                    try:
                        response = self.ai_client.chat.completions.create(
                            model=model_name,
                            messages=[{"role": "user", "content": prompt}],
                            temperature=0.1,
                            timeout=30,
                        )
                    except TypeError:
                        response = self.ai_client.chat.completions.create(
                            model=model_name,
                            messages=[{"role": "user", "content": prompt}],
                            temperature=0.1,
                        )
                    result_text = response.choices[0].message.content.strip()
                else:
                    return None

            match = re.search(r'\d{4,}', result_text)
            if match:
                return match.group(0)
        except Exception as e:
            print(f"AI Core 识别失败 ({self.ai_provider}): {e}")
            
        return None

    def _extract_from_original(self, original_data: Dict, possible_keys: list) -> Any:
        """从原始数据中提取字段值（支持多种可能的键名）"""
        for key in possible_keys:
            if key in original_data:
                value = original_data[key]
                if value and not (isinstance(value, float) and pd.isna(value)):
                    return value
        return None

    def _recognize_business_type(self, summary: str) -> Dict:
        """识别业务类型"""
        summary_lower = summary.lower()

        for business_type, rule in self.business_type_rules.items():
            for keyword in rule["keywords"]:
                if keyword in summary_lower:
                    result = {
                        "摘要编码": rule.get("summary_code"),
                        "类型": rule.get("type"),
                    }
                    account = rule.get("account")
                    # 银行类科目(1001/1002)需满足转账/存入等触发词
                    if account and str(account) in ["1001", "1002"] and not self._contains_bank_transfer(summary):
                        account = None
                    if account:
                        result["科目编码"] = account
                    return result

        return {}

    def _recognize_partner(self, summary: str) -> Optional[Dict]:
        """识别往来单位"""
        # 从缓存中匹配
        partner_names = getattr(self, "partner_names_sorted", list(self.partners.keys()))
        for partner_name in partner_names:
            if partner_name in summary:
                return {
                    "code": self.partners.get(partner_name, ""),
                    "name": partner_name
                }

        # 模糊匹配常见客户名称模式
        # 如：公司名、商店名、个人名等
        patterns = [
            r'([A-Z\s]+(?:CO|CA|COMPANY|CORP|LTD))',  # 英文公司名
            r'([\u4e00-\u9fa5]{2,20}(?:公司|商店|贸易|进出口))',  # 中文公司名
        ]

        for pattern in patterns:
            match = re.search(pattern, summary, re.IGNORECASE)
            if match:
                potential_name = match.group(1).strip()
                return {
                    "code": "",
                    "name": potential_name
                }

        return None

    def _recognize_account(self, summary: str) -> Optional[str]:
        """识别科目编码"""
        if not summary:
            return None

        summary_lower = summary.lower()

        # 0. 缓存优先（允许手工/AI 缓存覆盖默认规则）
        if self.base_data_mgr:
            normalized = self._normalize_summary(summary)
            cached = self.base_data_mgr.get_cached_recognition(normalized)
            if cached:
                return cached
            fuzzy_cached = self.base_data_mgr.get_cached_recognition_fuzzy(normalized)
            if fuzzy_cached:
                return fuzzy_cached

        # 小工具：按名称列表寻找科目编码
        def _find_account_by_names(names: list) -> Optional[str]:
            for name in names:
                if name in self.accounts:
                    return self.accounts[name]
            return None

        # 1. 高优先级关键词规则（从数据库加载）
        # self.account_keyword_rules 结构: [(keywords_list, fallback_code), ...]
        for keywords, fallback in self.account_keyword_rules:
            if any(k in summary_lower for k in keywords):
                # 数据库现在直接存储 fallback code (如 "100102")，
                # 不再存储 "target_names" 列表去反查名称，简化逻辑。
                # 如果未来需要“按名称查找”功能，需扩展表结构。
                # 目前直接返回配置的科目代码。
                return fallback

        # 2. 优先匹配完整名称或长名称（防止"办公费"误匹配到"费"）
        # 按长度降序排序所有键（包含名称、代码、全称）
        sorted_keys = sorted(self.accounts.keys(), key=len, reverse=True)

        for key in sorted_keys:
            # 忽略纯数字键的直接文本包含匹配（防止"2025"年匹配到科目2025），除非它被特殊符号包裹
            if key.isdigit():
                continue

            if key in summary:
                candidate = self.accounts[key]
                if str(candidate) in ["1001", "1002"] and not self._contains_bank_transfer(summary):
                    continue
                return candidate

        # 3. 尝试匹配被特殊符号包裹的代码，如 [1001] 或 (1001)
        code_match = re.search(r'[\[\(](\d{4,})[\]\)]', summary)
        if code_match:
            code = code_match.group(1)
            if code in self.accounts:
                candidate = self.accounts[code]
                if str(candidate) in ["1001", "1002"] and not self._contains_bank_transfer(summary):
                    candidate = None
                if candidate:
                    return candidate

        # 4. 尝试匹配独立的数字代码（前后非数字）
        # 例如 "科目 1001 "
        digit_matches = re.finditer(r'(?<!\d)(\d{4,})(?!\d)', summary)
        for match in digit_matches:
            code = match.group(1)
            if code in self.accounts:
                candidate = self.accounts[code]
                if str(candidate) in ["1001", "1002"] and not self._contains_bank_transfer(summary):
                    continue
                return candidate

        return None

    def _contains_bank_transfer(self, summary: str) -> bool:
        """判断摘要是否包含明显的转账/存入银行线索"""
        if not summary:
            return False
        s = summary.lower()
        keywords = [
            # 仅当明确出现“转到/至/入”或具体银行简称时，才视为银行转账
            "转到", "转至", "转入", "划转", "划入", "划到",
            "存入", "存到", "汇入", "汇到", "打入",
            "转bac", "转st", "存bac", "存st", "转入bac", "转入st",
            "转到bac", "转到st", "bac", "st"
        ]
        return any(k in s for k in keywords)

    def _recognize_department(self, summary: str) -> Optional[str]:
        """识别部门"""
        summary_lower = summary.lower()
        dept_map = getattr(self, "departments_lower", None)
        if dept_map:
            for dept_name, dept_code in dept_map.items():
                if dept_name in summary_lower:
                    return dept_code
        else:
            for dept_name, dept_code in self.departments.items():
                if dept_name and str(dept_name).lower() in summary_lower:
                    return dept_code

        # 匹配常见部门关键词（从数据库加载）
        for keyword, code in self.dept_keyword_rules.items():
            if keyword in summary_lower:
                return code

        return None

    def _recognize_amount(self, summary: str) -> Optional[float]:
        """识别金额"""
        if summary is None:
            return None

        # 先尝试直接按纯数字解析，避免 1000 被错误截断成 100
        raw_text = str(summary).strip()
        plain = raw_text.replace(",", "")
        if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", plain):
            try:
                return float(plain)
            except ValueError:
                pass

        # 匹配金额模式：数字 + 单位
        patterns = [
            r'([-+]?\d+(?:\.\d+)?)\s*元',
            r'([-+]?\d+(?:\.\d+)?)\s*美元',
            r'USD\s*([-+]?\d+(?:,\d{3})*(?:\.\d+)?)',
            r'¥\s*([-+]?\d+(?:,\d{3})*(?:\.\d+)?)',
            r'\$\s*([-+]?\d+(?:,\d{3})*(?:\.\d+)?)',
            r'([-+]?\d+(?:,\d{3})*(?:\.\d+)?)',  # 带逗号/长数字
        ]

        for pattern in patterns:
            match = re.search(pattern, raw_text)
            if match:
                amount_str = match.group(1).replace(',', '')
                try:
                    amount = float(amount_str)
                    # 检查单位倍数
                    if '万' in raw_text:
                        amount *= 10000
                    elif '千' in raw_text:
                        amount *= 1000
                    return amount
                except ValueError:
                    continue

        return None

    def _recognize_date(self, summary: str) -> Optional[str]:
        """识别日期"""
        if not summary:
            return None

        # 预处理：标准化文本
        summary_clean = str(summary).strip()

        # 匹配日期模式
        patterns = [
            r'(\d{4})[年/-](\d{1,2})[月/-](\d{1,2})',  # 2025年1月15日
            r'(\d{4})(\d{2})(\d{2})',  # 20250115
            r'(\d{1,2})-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*-(\d{4})',  # 21-Nov-2025
            r'([A-Za-z]{3,9})\s+(\d{1,2}),\s*(\d{4})',  # Nov 21, 2025
        ]

        month_map = {
            "jan": "01", "feb": "02", "mar": "03", "apr": "04", "may": "05", "jun": "06",
            "jul": "07", "aug": "08", "sep": "09", "oct": "10", "nov": "11", "dec": "12",
        }

        for pattern in patterns:
            match = re.search(pattern, summary_clean, re.IGNORECASE)
            if match:
                # 兼容 dd-MMM-YYYY 和 MMM dd, YYYY
                if len(match.groups()) == 3 and match.group(2).isalpha() and match.group(1).isdigit():
                    day = match.group(1).zfill(2)
                    mon_key = match.group(2).lower()[:3]
                    month = month_map.get(mon_key, "")
                    year = match.group(3)
                elif len(match.groups()) == 3 and match.group(1).isalpha():
                    mon_key = match.group(1).lower()[:3]
                    month = month_map.get(mon_key, "")
                    day = match.group(2).zfill(2)
                    year = match.group(3)
                else:
                    year = match.group(1)
                    month = match.group(2).zfill(2)
                    day = match.group(3).zfill(2)

                try:
                    y_int = int(year)
                    # 只接受合理年份，避免将票据号/Token误判为日期
                    if y_int < 1900 or y_int > 2100:
                        continue
                except ValueError:
                    continue
                if not month:
                    continue
                return f"{year}{month}{day}"

        return None

    def _recognize_exchange_rate(self, value: str) -> Optional[float]:
        """识别汇率"""
        if not value:
            return None

        # 去除空格和常见符号
        value = str(value).strip().replace(',', '').replace('，', '')

        # 尝试转换为浮点数
        try:
            rate = float(value)
            # 汇率通常在0.0001到10000之间
            if 0.0001 <= rate <= 10000:
                return rate
        except (ValueError, TypeError):
            pass

        return None

    def batch_recognize(self, data_list: list) -> list:
        """
        批量识别

        Args:
            data_list: 数据列表，每个元素应包含'摘要'字段

        Returns:
            增强后的数据列表
        """
        results = []
        for row in data_list:
            if isinstance(row, dict):
                summary = row.get("摘要", "")
                recognized = self.recognize(summary, row)
                # 合并识别结果，原始数据优先
                enhanced_row = {**recognized, **row}
                results.append(enhanced_row)
            else:
                results.append(row)

        return results

    def calculate_ai_similarity(self, summary1: str, summary2: str) -> float:
        """使用 AI 计算两个摘要的语义相似度 (0.0 - 1.0)"""
        if not self.ai_client:
            return 0.0
            
        prompt = f"""
请作为一名会计专家，判断以下两个“交易摘要”是否极有可能指代同一笔业务。

摘要1：{summary1}
摘要2：{summary2}

注意：
1. 摘要1通常是内部记录（可能包含中文和简写）。
2. 摘要2通常是银行流水（可能包含英文、拼音、交易单号或手续费说明）。
3. 即使文字不完全相同，如果它们提到的时间、金额、对方单位或业务性质（如刷卡、转账、手续费）一致，则相似度较高。

请只返回一个 0.0 到 1.0 之间的数字作为相似度评分（1.0 表示完全确定是同一笔，0.0 表示完全无关）。
不要输出任何其他解释文字。
"""
        try:
            if self.ai_provider == "lm_studio":
                response = self.ai_client.chat.completions.create(
                    model=self.ai_model_name,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.0,
                    timeout=20
                )
                result_text = response.choices[0].message.content.strip()
            else:
                # 智谱或其它
                response = self.ai_client.chat.completions.create(
                    model="glm-4-flash",
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.1,
                )
                result_text = response.choices[0].message.content.strip()

            # 提取数字 (支持 0.85, 85%, 或 "评分: 0.8")
            result_text = result_text.replace('%', '')
            match = re.search(r'(\d+\.\d+)', result_text)
            if not match:
                match = re.search(r'(\d+)', result_text)
            
            if match:
                score = float(match.group(1))
                # 如果返回的是百分制 (如 80)，转为 0.8
                if score > 1.0:
                    score = score / 100.0
                return min(max(score, 0.0), 1.0)
        except Exception as e:
            print(f"AI 相似度计算失败: {e}")
            
        return 0.0

    def close(self):

        """关闭数据库连接"""
        if self.base_data_mgr:
            self.base_data_mgr.close()


# 便捷函数
def recognize_summary(summary: str) -> Dict[str, Any]:
    """便捷的摘要识别函数"""
    recognizer = SummaryIntelligence()
    result = recognizer.recognize(summary)
    recognizer.close()
    return result

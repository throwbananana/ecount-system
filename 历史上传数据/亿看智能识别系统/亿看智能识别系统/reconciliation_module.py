# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import difflib
import re
from datetime import datetime, timedelta
from itertools import combinations
from typing import Dict, List, Any, Optional, Tuple
from summary_intelligence import SummaryIntelligence

class StandardReconciler:
    """
    Standard Reconciler
    Accepts two DataFrames in the "Standard Template Format" (as defined in Template.xlsx)
    and performs reconciliation matching.
    """

    def __init__(self, base_data_mgr=None, summary_intelligence=None):
        self.base_data_mgr = base_data_mgr
        self.summary_intelligence = summary_intelligence
        self.log_callback = print

    def set_logger(self, callback):
        self.log_callback = callback

    def log(self, msg):
        if self.log_callback:
            self.log_callback(msg)

    def _get_client_mapping(self) -> Dict[str, str]:
        """Fetch Local Code -> Yikan Code mapping from Base Data Manager"""
        mapping = {}
        if not self.base_data_mgr:
            return mapping
        
        try:
            # Query business_partner table for local_code
            partners = self.base_data_mgr.query("business_partner")
            for p in partners:
                yikan_code = p.get('code')
                local_code = p.get('local_code')
                if yikan_code and local_code:
                    # Normalize local code
                    l_norm = str(local_code).strip().upper()
                    if l_norm:
                        mapping[l_norm] = str(yikan_code).strip()
        except Exception as e:
            self.log(f"Error fetching client mapping: {e}")
        return mapping

    def parse_standard_df(self, df: pd.DataFrame, source_type: str) -> pd.DataFrame:
        """
        Parse a Standard Format DataFrame into an internal format for reconciliation.
        Standard Columns: 
            '凭证日期', '序号', '会计凭证No.', '摘要', '科目编码', '往来单位编码', '金额', '外币金额', '类型'
        
        Internal Columns:
            'Date', 'Doc', 'Code', 'Desc', 'Debit', 'Credit', 'Amount' (Net), 'Mapped_Code' (for Local)
        
        source_type: 'Local' or 'Yikan'
        """
        # Ensure columns exist
        needed = ['凭证日期', '序号', '会计凭证No.', '摘要', '往来单位编码', '金额', '类型', '外币金额']
        for c in needed:
            if c not in df.columns:
                # Try synonyms or loose matching if strict check fails?
                # For now, assume standard format.
                if c == '外币金额':
                    df[c] = 0.0
                else:
                    df[c] = "" # Fill missing with empty string
        
        out = pd.DataFrame()
        raw_dates = df['凭证日期'].astype(str).str.strip()
        
        # Pre-process: Remove suffixes like " -1", " -2" common in Yikan (YYYY/MM/DD -X)
        # ONLY apply this to Yikan data to avoid stripping years from Local data (e.g. "21-12-2025" -> "21-12")
        if source_type == 'Yikan':
            # Regex: optional space, hyphen, digits at end of string
            raw_dates = raw_dates.str.replace(r'\s*-\d+$', '', regex=True)
        
        # Date Parsing Strategy
        # Local: often DD/MM/YYYY (e.g. 13/11/2025) -> dayfirst=True
        # Yikan: often YYYY/MM/DD (e.g. 2025/10/03) -> dayfirst=False (default safe for ISO)
        use_dayfirst = (source_type == 'Local')
        
        parsed_dates = pd.to_datetime(raw_dates, dayfirst=use_dayfirst, errors='coerce')
        
        # Fallback: Try opposite dayfirst strategy for NaT values (in case format is mixed or wrong assumption)
        if parsed_dates.isna().any():
            # Try filling NaTs with the other strategy
            fallback_dates = pd.to_datetime(raw_dates, dayfirst=not use_dayfirst, errors='coerce')
            parsed_dates = parsed_dates.fillna(fallback_dates)

        if parsed_dates.isna().any():
            numeric_dates = pd.to_numeric(raw_dates, errors='coerce')
            if numeric_dates.notna().any():
                excel_dates = pd.to_datetime(
                    numeric_dates,
                    unit='D',
                    origin='1899-12-30',
                    errors='coerce',
                )
                parsed_dates = parsed_dates.fillna(excel_dates)
        out['Date'] = parsed_dates
        if len(parsed_dates):
            ok_count = int(parsed_dates.notna().sum())
            if ok_count == 0:
                sample = raw_dates.dropna().astype(str).head(5).tolist()
                if sample:
                    self.log(f"[{source_type}] 凭证日期解析全部失败，样例: {sample}")
            elif ok_count < len(parsed_dates):
                bad_count = len(parsed_dates) - ok_count
                sample = raw_dates[parsed_dates.isna()].dropna().astype(str).head(5).tolist()
                if sample:
                    self.log(f"[{source_type}] 凭证日期解析失败 {bad_count}/{len(parsed_dates)} 行，样例: {sample}")
        
        # Doc logic: 序号 usually contains the Order No / Pedido. 会计凭证No is Voucher No.
        # Reconciliation uses Order No.
        out['Doc'] = df['序号'].astype(str).str.strip()
        
        out['Code'] = df['往来单位编码'].astype(str).str.strip()
        out['Desc'] = df['摘要'].astype(str).str.strip()
        
        # Calculate Debit/Credit
        # Logic update: If '借方' and '贷方' are present, trust them over '金额'
        # This prevents errors where user maps '金额' to 'Debit' column, making Credits 0.
        
        has_debit_credit = '借方' in df.columns and '贷方' in df.columns
        
        if has_debit_credit:
            # Use Debit/Credit columns
            d_vals = pd.to_numeric(df['借方'], errors='coerce').fillna(0)
            c_vals = pd.to_numeric(df['贷方'], errors='coerce').fillna(0)
            
            debits = d_vals.tolist()
            credits = c_vals.tolist()
            
            # Recalculate Amount and Type
            amount_vals = d_vals - c_vals
            
            # Infer Type if missing (Optional, but good for consistency)
            # If Type is already there, we might not need to overwrite, but Amount should be synced.
            # Let's keep Type if valid? No, if we recalculated Amount, we should trust D/C.
            
        else:
            # Use Amount + Type logic
            def get_amount(row):
                f_amt = row.get('外币金额')
                l_amt = row.get('金额')
                try:
                    f_val = float(f_amt) if f_amt and str(f_amt).strip() != "" else 0.0
                except: f_val = 0.0
                
                try:
                    l_val = float(l_amt) if l_amt and str(l_amt).strip() != "" else 0.0
                except: l_val = 0.0
                
                # Prefer Foreign if non-zero, else Local
                return f_val if abs(f_val) > 0.001 else l_val

            amount_vals = df.apply(get_amount, axis=1)
            types = df['类型'].astype(str).str.strip()
            
            debits = []
            credits = []
            
            for amt, t in zip(amount_vals, types):
                d = 0.0
                c = 0.0
                if t in ['3', '1', '借']: # Debit
                    d = amt
                elif t in ['4', '2', '贷']: # Credit
                    c = amt
                else:
                    d = amt 
                
                debits.append(d)
                credits.append(c)
                
        out['Debit'] = debits
        out['Credit'] = credits
        out['Amount'] = out['Debit'] - out['Credit']
        
        # Original row reference
        out['Orig_Row_Idx'] = df.index
        
        return out

    @staticmethod
    def _get_direction(debit: float, credit: float, tol: float = 0.001) -> str:
        if abs(debit) <= tol and abs(credit) <= tol:
            return ""
        if abs(debit) > abs(credit):
            return "D"
        if abs(credit) > abs(debit):
            return "C"
        return ""

    def _direction_matches(self, l_row: pd.Series, y_rows: pd.DataFrame) -> bool:
        l_dir = self._get_direction(l_row.get("Debit", 0), l_row.get("Credit", 0))
        y_dir = self._get_direction(y_rows["Debit"].sum(), y_rows["Credit"].sum())
        if not l_dir or not y_dir:
            return True
        return l_dir == y_dir

    def _build_match_row(self, row_l: pd.Series, y_rows: pd.DataFrame, reason: str, status: str) -> Dict[str, Any]:
        l_amt = row_l['Amount']
        y_sum = y_rows['Amount'].sum()
        diff = l_amt - y_sum
        return {
            '匹配状态': status,
            '匹配原因': reason,
            '差额': diff,
            # Local Info
            '当地_日期': row_l['Date'],
            '当地_单号': row_l['Doc'],
            '当地_编码': row_l['Code'],
            '当地_映射编码': row_l['Mapped_Code'],
            '当地_借方': row_l['Debit'],
            '当地_贷方': row_l['Credit'],
            # Yikan Info
            '亿看_日期': ", ".join([d.strftime('%Y-%m-%d') for d in y_rows['Date'] if pd.notna(d)]),
            '亿看_单号': ", ".join(y_rows['Doc'].astype(str).unique()),
            '亿看_编码': ", ".join(y_rows['Code'].astype(str).unique()),
            '亿看_借方': y_rows['Debit'].sum(),
            '亿看_贷方': y_rows['Credit'].sum(),
        }

    def map_columns_smart(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Intelligently map columns from a raw DataFrame to standard reconciliation columns.
        Target: ['凭证日期', '序号', '摘要', '往来单位编码', '金额', '类型']
        """
        # Define synonyms mapping
        synonyms = {
            '凭证日期': ['日期', 'Date', 'Fecha', '记账日期', '制单日期'],
            '序号': ['序号', '单据号', 'No.', 'Doc', 'Document', 'Pedido', 'Order', '债权债务号码', '债权债务', '债券', '债券编号', '债券代码'],
            '摘要': ['摘要', 'Desc', 'Description', '备注', 'Note', 'Concepto', 'Detalle'],
            '往来单位编码': ['往来', '单位', 'Code', 'Client', 'Customer', 'Vendor', 'Cliente', 'Proveedor', 'Codig', 'Codigo'],
            '金额': ['金额', 'Amount', 'Monto', 'Total', 'Neto', 'Importe'],
            '类型': ['类型', 'Type', 'Tipo', '借贷', 'DrCr', 'DC'],
            # Optional specific debit/credit
            '借方': ['借方', 'Debit', 'Debito', 'Dr', 'Cargo'],
            '贷方': ['贷方', 'Credit', 'Credito', 'Cr', 'Abono'],
        }

        mapping = {}
        used_cols = set()

        # Helper to find best match
        def find_best_col(target_key):
            best_col = None
            best_score = 0.0
            
            for col in df.columns:
                if col in used_cols: continue
                
                # Exact match synonyms
                col_norm = str(col).strip().lower()
                for syn in synonyms.get(target_key, []):
                    syn_norm = str(syn).strip().lower()
                    if syn_norm == col_norm:
                        return col # Found exact
                    if syn_norm in col_norm:
                        score = 0.8
                        if score > best_score:
                            best_score = score
                            best_col = col
            return best_col

        # Map basic columns
        for key in ['凭证日期', '序号', '摘要', '往来单位编码', '类型']:
            col = find_best_col(key)
            if col:
                mapping[key] = col
                used_cols.add(col)

        # Map Amount/Debit/Credit special logic
        debit_col = find_best_col('借方')
        credit_col = find_best_col('贷方')
        
        if debit_col and credit_col:
            # We have split columns, need to synthesize '金额' and '类型' if missing
            # But the Standard format expects '金额' + '类型' OR just '金额' (signed).
            # We will construct a new DF in Standard Format
            pass
        else:
            # Look for single Amount column
            amt_col = find_best_col('金额')
            if amt_col:
                mapping['金额'] = amt_col
                used_cols.add(amt_col)

        # Construct new DF
        new_df = pd.DataFrame()
        
        # Copy mapped columns
        for std_col, src_col in mapping.items():
            new_df[std_col] = df[src_col]

        # Handle split Debit/Credit if found and '金额' not mapped
        if '金额' not in new_df.columns and debit_col and credit_col:
            # Synthesize
            # Local Convention: Debit is Positive, Credit is Negative usually for AR?
            # Or we use Type 3/4.
            # Let's populate '金额' and '类型'
            d_vals = pd.to_numeric(df[debit_col], errors='coerce').fillna(0)
            c_vals = pd.to_numeric(df[credit_col], errors='coerce').fillna(0)
            
            # Net Amount
            # Assume Net = Debit - Credit
            new_df['金额'] = d_vals - c_vals
            # We can also populate '外币金额' if needed, but for now just Amount
            # Type is implicitly handled by Amount sign if we leave Type empty?
            # parse_standard_df uses Type if present.
            
            types = []
            amts = []
            for d, c in zip(d_vals, c_vals):
                if abs(d) > 0.001:
                    types.append('3') # Debit
                    amts.append(d)
                elif abs(c) > 0.001:
                    types.append('4') # Credit
                    amts.append(c)
                else:
                    types.append('')
                    amts.append(0.0)
            new_df['类型'] = types
            new_df['金额'] = amts

        # Fill missing required columns with defaults
        required = ['凭证日期', '序号', '会计凭证No.', '摘要', '往来单位编码', '金额', '类型']
        for req in required:
            if req not in new_df.columns:
                new_df[req] = ""

        return new_df

    def analyze_mismatches_with_ai(self, df_local_unmatched: pd.DataFrame, df_yikan_unmatched: pd.DataFrame) -> List[Dict]:
        """
        Use AI to find potential semantic matches in unmatched records.
        Returns a list of suggested matches.
        """
        if not self.summary_intelligence or not self.summary_intelligence.ai_client:
            self.log("AI Client not available for analysis.")
            return []

        suggestions = []
        
        # Optimize: Group by Code
        # We only compare Local and Yikan items that share the same (or mapped) code.
        # But unmatched means maybe code was wrong too?
        # Let's assume Code mapping is handled by strategy 1/2.
        # Here we look for deep semantic matches (e.g. diff description, small amt diff).
        
        # Group Yikan by Code
        yikan_groups = {}
        for idx, row in df_yikan_unmatched.iterrows():
            code_val = row.get('往来单位编码', '')
            if pd.isna(code_val): code_val = ""
            code = str(code_val).strip().upper()
            if code == "NAN": code = "" # Extra safety
            
            if code not in yikan_groups: yikan_groups[code] = []
            yikan_groups[code].append(row)

        self.log(f"AI Analysis: Local Unmatched Count={len(df_local_unmatched)}")
        self.log(f"AI Analysis: Yikan Unmatched Groups Count={len(yikan_groups)}")
        self.log(f"AI Analysis: Yikan Group Keys: {list(yikan_groups.keys())}")

        # Iterate Local Unmatched
        skipped_no_code = 0
        skipped_no_candidates = 0
        skipped_amt_filter = 0
        
        for idx, l_row in df_local_unmatched.iterrows():
            # Use mapped code if available (passed from reconcile)
            l_code_val = l_row.get('AI_Mapped_Code', l_row.get('往来单位编码', ''))
            if pd.isna(l_code_val): l_code_val = ""
            l_code = str(l_code_val).strip().upper()
            if l_code == "NAN": l_code = ""
            
            potential_y_rows = list(yikan_groups.get(l_code, []))
            
            # Fallback/Expansion: Always include Yikan items with EMPTY code (likely Bank Fees, etc.)
            # or if the code mismatch is the issue.
            if "" in yikan_groups:
                # Avoid duplicates if l_code is "" (already included)
                if l_code != "":
                    potential_y_rows.extend(yikan_groups[""])
            
            # --- GLOBAL FALLBACK: If still no candidates, search ALL unmatched Yikan rows ---
            # This addresses "I want analysis even if no [code] match found"
            if not potential_y_rows:
                self.log(f"Debug: No Code match for '{l_code}'. Switching to GLOBAL AMOUNT SEARCH across {len(df_yikan_unmatched)} rows.")
                # We can't just list all rows, we should iterate all and let the Amount Filter do its job.
                # However, to be efficient, let's just pass the dataframe rows.
                potential_y_rows = [r for _, r in df_yikan_unmatched.iterrows()]
            # --------------------------------------------------------------------------------
            
            l_desc = str(l_row.get('摘要', ''))
            l_amt = l_row.get('金额', 0)
            l_date = str(l_row.get('凭证日期', ''))

            # Filter candidates by Amount proximity (e.g. within 10%)
            candidates = []
            # Keep track of all differences for fallback
            all_candidates_with_diff = []
            
            for y_row in potential_y_rows:
                y_amt = y_row.get('金额', 0)
                try:
                    l_val = float(l_amt)
                    y_val = float(y_amt)
                    diff = abs(l_val - y_val)
                    all_candidates_with_diff.append((diff, y_row))
                    
                    # Relaxed tolerance for AI: 20% or +/- 10.0
                    if diff < 10.0 or (l_val != 0 and abs((l_val - y_val)/l_val) < 0.2):
                        candidates.append(y_row)
                except Exception as e:
                    # Log conversion errors but continue processing other rows
                    # self.log(f"Warning: Amount conversion failed for comparison. Local={l_amt}, Yikan={y_amt}. Error: {e}")
                    pass
            
            # --- FORCE CANDIDATES: If no candidates passed strict filter, take Top 5 closest ---
            if not candidates and all_candidates_with_diff:
                self.log(f"Debug: No strict amount match. Taking Top 5 closest by amount.")
                # Sort by difference ascending
                all_candidates_with_diff.sort(key=lambda x: x[0])
                candidates = [x[1] for x in all_candidates_with_diff[:5]]
            # -----------------------------------------------------------------------------------
            
            if not candidates: 
                skipped_amt_filter += 1
                # self.log(f"Debug: No amount match candidates for Code={l_code}, Amt={l_amt}")
                continue
            
            # Prepare Prompt
            candidates_str = ""
            for i, cand in enumerate(candidates[:5]): # Limit to 5
                candidates_str += f"[{i}] Date:{cand.get('凭证日期')} Doc:{cand.get('序号')} Amt:{cand.get('金额')} Desc:{cand.get('摘要')}\n"

            prompt = f"""
            Compare this Local transaction with candidates from Yikan system.
            Local: Date={l_date}, Doc={l_row.get('序号')}, Amt={l_amt}, Desc={l_desc}
            
            Candidates:
            {candidates_str}
            
            Are any of these likely the same transaction despite small differences?
            Return index of best match or -1 if none. Return strictly just the number.
            """
            
            self.log(f"Debug: Full Prompt Sent to AI:\n{prompt}") # <--- ADDED LOG

            try:
                # Call AI
                # Use simple completion
                self.log(f"Calling AI Analysis... Local Doc={l_row.get('序号')} ({l_amt}) vs {len(candidates)} Candidates")
                if hasattr(self.summary_intelligence, 'ai_base_url'):
                     # self.log(f"AI Base URL: {self.summary_intelligence.ai_base_url}")
                     pass

                response = self.summary_intelligence.ai_client.chat.completions.create(
                    model=self.summary_intelligence.ai_model_name,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.1
                )
                ans = response.choices[0].message.content.strip()
                self.log(f"AI Response: {ans}") # <--- ADDED LOG
                
                match = re.search(r'-?\d+', ans)
                if match:
                    idx_val = int(match.group(0))
                    if idx_val >= 0 and idx_val < len(candidates):
                        # Found a match
                        best_y = candidates[idx_val]
                        suggestions.append({
                            'Local_Doc': l_row.get('序号'),
                            'Local_Desc': l_desc,
                            'Yikan_Doc': best_y.get('序号'),
                            'Yikan_Desc': best_y.get('摘要'),
                            'Reason': 'AI Detected Semantic Match'
                        })
            except Exception as e:
                self.log(f"AI Analysis Call Failed: {e}")
                pass # Ignore AI errors

        self.log(f"AI Analysis Summary: Skipped (No Candidates)={skipped_no_candidates}, Skipped (Amount Filter)={skipped_amt_filter}")
        return suggestions

    def reconcile(self, df_local_std: pd.DataFrame, df_yikan_std: pd.DataFrame, config: Dict[str, Any]) -> Dict[str, pd.DataFrame]:
        """
        Main Reconciliation Entry Point
        """
        self.log("Starting Reconciliation...")
        
        # 1. Parse Data
        df_local = self.parse_standard_df(df_local_std, 'Local')
        df_yikan = self.parse_standard_df(df_yikan_std, 'Yikan')
        
        # 2. Date Filter
        start_date = config.get('start_date') # datetime object
        end_date = config.get('end_date')     # datetime object

        self.log(
            "Parsed Rows (pre-filter): "
            f"Local={len(df_local)} (Dates: {int(df_local['Date'].notna().sum())}), "
            f"Yikan={len(df_yikan)} (Dates: {int(df_yikan['Date'].notna().sum())})"
        )
        if df_local['Date'].notna().any():
            l_min = df_local['Date'].min().date()
            l_max = df_local['Date'].max().date()
            self.log(f"Local Date Range: {l_min} to {l_max}")
        if df_yikan['Date'].notna().any():
            y_min = df_yikan['Date'].min().date()
            y_max = df_yikan['Date'].max().date()
            self.log(f"Yikan Date Range: {y_min} to {y_max}")
        
        if start_date and end_date:
            self.log(f"Filtering Date Range: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
            if df_local['Date'].notna().sum() == 0 or df_yikan['Date'].notna().sum() == 0:
                self.log("Warning: 日期列无法解析，日期过滤可能导致结果为空。")
            df_local = df_local[(df_local['Date'] >= start_date) & (df_local['Date'] <= end_date)].copy()
            df_yikan = df_yikan[(df_yikan['Date'] >= start_date) & (df_yikan['Date'] <= end_date)].copy()
            
        self.log(f"Local Rows: {len(df_local)}, Yikan Rows: {len(df_yikan)}")
        
        # 3. Code Mapping
        mapping = self._get_client_mapping()
        
        def map_code(row):
            raw = row['Code']
            # Direct Map
            if raw in mapping:
                return mapping[raw]
            # Fuzzy Map (e.g. L021 -> L0021)
            if config.get('fuzzy_code', True):
                 # Try adding '0's
                 if len(raw) > 1 and raw[0].isalpha():
                     c1 = raw[0] + '0' + raw[1:]
                     if c1 in mapping.values(): return c1 # Check against known Yikan codes?
                     # We need a set of valid Yikan codes
                     pass 
            return raw

        # Build valid Yikan codes set from Data + DB
        valid_yikan_codes = set(df_yikan['Code'].unique())
        if self.base_data_mgr:
            # Add all codes from DB
            partners = self.base_data_mgr.query("business_partner")
            for p in partners:
                if p.get('code'):
                    valid_yikan_codes.add(p['code'])
        
        def get_mapped_code(raw_code):
            norm = str(raw_code).strip().upper()
            if norm in valid_yikan_codes:
                return norm
            if norm in mapping:
                return mapping[norm]
            
            # Fuzzy match against valid_yikan_codes
            if config.get('fuzzy_code', True):
                if len(norm) > 1 and norm[0].isalpha():
                    c1 = norm[0] + '0' + norm[1:]
                    if c1 in valid_yikan_codes: return c1
                    c2 = norm[0] + '00' + norm[1:]
                    if c2 in valid_yikan_codes: return c2
            return norm

        df_local['Mapped_Code'] = df_local['Code'].apply(get_mapped_code)
        
        # 4. Initialize Match Status
        df_local['Match_Status'] = 'Unmatched'
        df_local['Match_Note'] = ''
        df_yikan['Match_Status'] = 'Unmatched'
        df_yikan['Match_Note'] = ''
        
        local_indices = set(df_local.index)
        yikan_indices = set(df_yikan.index)
        
        matched_results = []
        direction_mismatch_results = []
        direction_mismatch_keys = set()
        require_same_direction = bool(config.get("require_same_direction"))

        def record_direction_mismatch(l_idx, y_idxs, reason):
            key = (l_idx, tuple(sorted(y_idxs)))
            if key in direction_mismatch_keys:
                return
            direction_mismatch_keys.add(key)
            if l_idx in local_indices:
                local_indices.remove(l_idx)
            for y in y_idxs:
                if y in yikan_indices:
                    yikan_indices.remove(y)
            row_l = df_local.loc[l_idx]
            y_rows = df_yikan.loc[y_idxs]
            direction_mismatch_results.append(
                self._build_match_row(row_l, y_rows, reason, "方向不一致")
            )
        
        # --- Strategy 0: Global Unique Amount ---
        self.log("Running Strategy 0: Unique Amount...")
        # (Simplified logic from original)
        l_amounts = {}
        for idx in local_indices:
            amt = df_local.at[idx, 'Amount']
            if abs(amt) < 0.01: continue
            if amt not in l_amounts: l_amounts[amt] = []
            l_amounts[amt].append(idx)
            
        y_amounts = {}
        for idx in yikan_indices:
            amt = df_yikan.at[idx, 'Amount'] # Yikan Net is Debit - Credit
            if abs(amt) < 0.01: continue
            if amt not in y_amounts: y_amounts[amt] = []
            y_amounts[amt].append(idx)
            
        s0_count = 0
        for amt, l_idxs in l_amounts.items():
            if len(l_idxs) == 1:
                # Try exact match
                if amt in y_amounts and len(y_amounts[amt]) == 1:
                    l_idx = l_idxs[0]
                    y_idx = y_amounts[amt][0]
                    if require_same_direction:
                        if not self._direction_matches(df_local.loc[l_idx], df_yikan.loc[[y_idx]]):
                            record_direction_mismatch(l_idx, [y_idx], "Strategy 0: Unique Amount (方向不一致)")
                            continue
                    self._commit_match(df_local, df_yikan, l_idx, y_idx, "Strategy 0: Unique Amount", matched_results, local_indices, yikan_indices)
                    s0_count += 1
                elif require_same_direction and (-amt in y_amounts) and len(y_amounts[-amt]) == 1:
                    l_idx = l_idxs[0]
                    y_idx = y_amounts[-amt][0]
                    record_direction_mismatch(l_idx, [y_idx], "Strategy 0: Unique Amount (Opposite Direction)")
                    continue
        self.log(f"Strategy 0 matched {s0_count} items.")
        
        # --- Strategy 0.5: Date + Amount (Ignore Doc/Code, with Date Tolerance) ---
        self.log("Running Strategy 0.5: Date + Amount (Tolerance +/- 3 days)...")
        # Useful for Bank Reconciliation where Document Numbers don't match and Dates might be off by a few days
        s05_count = 0
        
        # Optimize: Sort Yikan by Date for window search or just iterate (dataset is small enough usually)
        # For simplicity and robustness with small-medium datasets (<10k rows), iteration is fine.
        # But we only want to match UNMATCHED items.
        
        # Convert Yikan Dates/Amounts to list for fast iteration
        y_candidates = []
        for idx in yikan_indices:
            date_val = df_yikan.at[idx, 'Date']
            if pd.isna(date_val): continue
            amt = df_yikan.at[idx, 'Amount']
            if abs(amt) < 0.01: continue
            y_candidates.append({
                'idx': idx,
                'date': date_val,
                'amt': amt
            })

        tolerance_days = timedelta(days=3)

        for l_idx in list(local_indices):
            if l_idx not in local_indices: continue
            
            l_date = df_local.at[l_idx, 'Date']
            if pd.isna(l_date): continue
            l_amt = df_local.at[l_idx, 'Amount']
            if abs(l_amt) < 0.01: continue
            
            # Search Yikan
            best_y = None
            min_day_diff = 100
            best_y_opposite = None
            min_day_diff_opposite = 100
            
            for cand in y_candidates:
                if cand['idx'] not in yikan_indices: continue
                
                # Check Amount First (Exact or Opposite)
                # Strict amount match required for loose date match
                y_amt = cand['amt']
                match_same = abs(l_amt - y_amt) < 0.05
                match_opp = abs(l_amt + y_amt) < 0.05
                if not (match_same or match_opp):
                    continue
                
                # Check Date Tolerance
                diff = abs((l_date - cand['date']).days)
                if diff <= 3:
                    # Found candidate
                    # Pick the one with closest date
                    if match_same:
                        if diff < min_day_diff:
                            min_day_diff = diff
                            best_y = cand['idx']
                    elif match_opp and require_same_direction:
                        if diff < min_day_diff_opposite:
                            min_day_diff_opposite = diff
                            best_y_opposite = cand['idx']
                    if diff == 0 and best_y is not None: # Exact date match is ideal
                        break
            
            if best_y is not None:
                if require_same_direction:
                    if not self._direction_matches(df_local.loc[l_idx], df_yikan.loc[[best_y]]):
                        record_direction_mismatch(l_idx, [best_y], f"Strategy 0.5: Amt Match, Date Diff {min_day_diff}d (方向不一致)")
                        continue
                reason = f"Strategy 0.5: Amt Match, Date Diff {min_day_diff}d"
                self._commit_match(df_local, df_yikan, l_idx, best_y, reason, matched_results, local_indices, yikan_indices)
                s05_count += 1
            elif require_same_direction and best_y_opposite is not None:
                record_direction_mismatch(
                    l_idx,
                    [best_y_opposite],
                    f"Strategy 0.5: Opposite Amt, Date Diff {min_day_diff_opposite}d",
                )
        
        self.log(f"Strategy 0.5 matched {s05_count} items.")

        # --- Strategy 1: Code + Doc ---
        self.log("Running Strategy 1: Code + Doc...")
        # Group Yikan by Code
        y_by_code = {}
        for idx in yikan_indices:
            code = df_yikan.at[idx, 'Code'] # Use Yikan raw code (which is standard code)
            if code not in y_by_code: y_by_code[code] = []
            y_by_code[code].append(idx)
            
        for l_idx in list(local_indices): # Copy list to iterate
            if l_idx not in local_indices: continue
            
            code = df_local.at[l_idx, 'Mapped_Code']
            doc = df_local.at[l_idx, 'Doc']
            if not doc: continue
            
            # Clean Doc (remove .0)
            if doc.endswith('.0'): doc = doc[:-2]
            
            candidates = y_by_code.get(code, [])
            potential_y = []
            for y_idx in candidates:
                if y_idx not in yikan_indices: continue
                y_doc = str(df_yikan.at[y_idx, 'Doc'])
                if y_doc.endswith('.0'): y_doc = y_doc[:-2]
                
                # Split Y Doc
                y_docs_split = re.split(r'[,\s/]+', y_doc)
                if doc in y_docs_split:
                    potential_y.append(y_idx)
            
            l_amt = df_local.at[l_idx, 'Amount']
            if potential_y:
                # Check Sum
                y_sum = sum(df_yikan.at[y, 'Amount'] for y in potential_y)
                
                # Standard Match (Same Sign)
                if abs(l_amt - y_sum) < 0.05:
                     if require_same_direction and not self._direction_matches(df_local.loc[l_idx], df_yikan.loc[potential_y]):
                         record_direction_mismatch(l_idx, potential_y, f"Strategy 1: Doc {doc} (方向不一致)")
                     else:
                         self._commit_match_multi(df_local, df_yikan, l_idx, potential_y, f"Strategy 1: Doc {doc}", matched_results, local_indices, yikan_indices)
                
                # Opposite Sign Match (Bank Recon often needs this)
                elif abs(l_amt + y_sum) < 0.05:
                     if require_same_direction:
                         record_direction_mismatch(l_idx, potential_y, f"Strategy 1: Doc {doc} (Opposite Direction)")
                     else:
                         self._commit_match_multi(df_local, df_yikan, l_idx, potential_y, f"Strategy 1: Doc {doc} (Opposite Sign)", matched_results, local_indices, yikan_indices)
                
                elif len(potential_y) > 1 and len(potential_y) <= 8:
                     # Subset sum
                     found = False
                     for r in range(1, len(potential_y)):
                         if found: break
                         for sub in combinations(potential_y, r):
                             s_sum = sum(df_yikan.at[y, 'Amount'] for y in sub)
                             if abs(l_amt - s_sum) < 0.05:
                                 if require_same_direction and not self._direction_matches(df_local.loc[l_idx], df_yikan.loc[list(sub)]):
                                     record_direction_mismatch(l_idx, list(sub), f"Strategy 1.2: Subset Doc {doc} (方向不一致)")
                                 else:
                                     self._commit_match_multi(df_local, df_yikan, l_idx, list(sub), f"Strategy 1.2: Subset Doc {doc}", matched_results, local_indices, yikan_indices)
                                     found = True
                                     break
                             # Subset Opposite
                             if abs(l_amt + s_sum) < 0.05:
                                 if require_same_direction:
                                     record_direction_mismatch(l_idx, list(sub), f"Strategy 1.2: Subset Doc {doc} (Opposite)")
                                 else:
                                     self._commit_match_multi(df_local, df_yikan, l_idx, list(sub), f"Strategy 1.2: Subset Doc {doc} (Opposite)", matched_results, local_indices, yikan_indices)
                                     found = True
                                     break
                     if not found:
                         self.log(f"Strategy 1 Failed for Doc={doc}, Code={code}: LocalAmt={l_amt}, CandidatesSum={y_sum} (Count={len(potential_y)})")
                else:
                    self.log(f"Strategy 1 Mismatch for Doc={doc}, Code={code}: LocalAmt={l_amt} != YikanSum={y_sum} (and not opposite)")
            else:
                # Log only if specific conditions met (e.g. large amount) to avoid spam
                # self.log(f"Strategy 1: No Yikan doc found for Local Doc={doc} (Code={code})")
                pass

        # --- Strategy 2: Code + Amount + Desc (Approx) ---
        self.log("Running Strategy 2: Code + Amount + Desc...")
        # Re-index remaining Yikan
        y_by_code = {}
        for idx in yikan_indices:
            code = df_yikan.at[idx, 'Code']
            if code not in y_by_code: y_by_code[code] = []
            y_by_code[code].append(idx)
            
        s2_candidates = []
        for l_idx in local_indices:
            code = df_local.at[l_idx, 'Mapped_Code']
            l_amt = df_local.at[l_idx, 'Amount']
            l_desc = df_local.at[l_idx, 'Desc'].lower()
            
            candidates = y_by_code.get(code, [])
            for y_idx in candidates:
                if y_idx not in yikan_indices: continue
                y_amt = df_yikan.at[y_idx, 'Amount']
                
                if abs(l_amt - y_amt) < 0.05:
                    y_desc = df_yikan.at[y_idx, 'Desc'].lower()
                    score = difflib.SequenceMatcher(None, l_desc, y_desc).ratio()
                    s2_candidates.append({
                        'l': l_idx, 'y': y_idx, 'score': score, 'y_desc': y_desc, 'l_desc': l_desc
                    })
        
        s2_candidates.sort(key=lambda x: -x['score'])
        s2_matched = 0
        for c in s2_candidates:
            if c['l'] in local_indices and c['y'] in yikan_indices:
                if require_same_direction and not self._direction_matches(df_local.loc[c['l']], df_yikan.loc[[c['y']]]):
                    record_direction_mismatch(c['l'], [c['y']], f"Strategy 2: Amount+Desc({c['score']:.2f}) (方向不一致)")
                    continue
                self.log(f"Strategy 2 Match: Score={c['score']:.2f}, L='{c['l_desc']}' vs Y='{c['y_desc']}'")
                self._commit_match(df_local, df_yikan, c['l'], c['y'], f"Strategy 2: Amount+Desc({c['score']:.2f})", matched_results, local_indices, yikan_indices)
                s2_matched += 1
        self.log(f"Strategy 2 matched {s2_matched} items.")

        # --- Strategy 1.5: Code + Subset Sum (Global) ---
        self.log("Running Strategy 1.5: Code + Subset Sum...")
        # (Simplified for brevity: Limit to size 4)
        s15_count = 0
        for l_idx in list(local_indices):
            if l_idx not in local_indices: continue
            code = df_local.at[l_idx, 'Mapped_Code']
            l_amt = df_local.at[l_idx, 'Amount']
            if abs(l_amt) < 0.01: continue
            
            candidates = [y for y in y_by_code.get(code, []) if y in yikan_indices]
            if not candidates or len(candidates) > 10: continue # Skip if too many
            
            found = False
            for r in range(1, 5):
                if found or r > len(candidates): break
                for sub in combinations(candidates, r):
                    s_sum = sum(df_yikan.at[y, 'Amount'] for y in sub)
                    if abs(l_amt - s_sum) < 0.05:
                         if require_same_direction and not self._direction_matches(df_local.loc[l_idx], df_yikan.loc[list(sub)]):
                             record_direction_mismatch(l_idx, list(sub), "Strategy 1.5: Global Sum (方向不一致)")
                         else:
                             self._commit_match_multi(df_local, df_yikan, l_idx, list(sub), "Strategy 1.5: Global Sum", matched_results, local_indices, yikan_indices)
                             found = True
                             s15_count += 1
                             break
                    if abs(l_amt + s_sum) < 0.05 and require_same_direction:
                         record_direction_mismatch(l_idx, list(sub), "Strategy 1.5: Global Sum (Opposite)")
                         found = True
                         break
        self.log(f"Strategy 1.5 matched {s15_count} items.")
                         
        # --- Strategy 3: Global Opt (Loose) ---
        # Skipping for now to save time/complexity, Strat 2 covers most exact amount matches.
        
        # 5. Build Result DataFrames (Standard Format)
        
        # Matched
        df_matched_out = pd.DataFrame(matched_results)
        
        # Unmatched Local
        df_local_unmatched = self._restore_standard_format(df_local.loc[list(local_indices)].copy(), df_local_std)
        
        # Inject AI_Mapped_Code PERMANENTLY into the result
        # This fixes the on-demand AI analysis in the GUI which uses this dataframe
        map_orig_to_code = dict(zip(df_local['Orig_Row_Idx'], df_local['Mapped_Code']))
        df_local_unmatched['AI_Mapped_Code'] = df_local_unmatched.index.map(map_orig_to_code)
        
        # Unmatched Yikan
        df_yikan_unmatched = self._restore_standard_format(df_yikan.loc[list(yikan_indices)].copy(), df_yikan_std)
        
        # 6. AI Analysis (Optional)
        ai_suggestions = []
        if config.get('use_ai_analysis', False):
            self.log("Running AI Error Analysis...")
            ai_suggestions = self.analyze_mismatches_with_ai(df_local_unmatched, df_yikan_unmatched)
            if ai_suggestions:
                self.log(f"AI found {len(ai_suggestions)} potential matches.")
        
        return {
            'matched': df_matched_out,
            'unmatched_local': df_local_unmatched,
            'unmatched_yikan': df_yikan_unmatched,
            'direction_mismatch': pd.DataFrame(direction_mismatch_results),
            'ai_suggestions': pd.DataFrame(ai_suggestions) if ai_suggestions else pd.DataFrame()
        }

    def _commit_match(self, df_l, df_y, l_idx, y_idx, reason, results, l_set, y_set):
        self._commit_match_multi(df_l, df_y, l_idx, [y_idx], reason, results, l_set, y_set)

    def _commit_match_multi(self, df_l, df_y, l_idx, y_idxs, reason, results, l_set, y_set):
        # Mark used
        if l_idx in l_set:
            l_set.remove(l_idx)
        for y in y_idxs:
            if y in y_set: y_set.remove(y)
            
        row_l = df_l.loc[l_idx]
        y_rows = df_y.loc[y_idxs]
        results.append(self._build_match_row(row_l, y_rows, reason, "已匹配"))

    def _restore_standard_format(self, internal_df, original_std_df):
        """
        Restore internal DF to Standard Template format using original data.
        """
        if internal_df.empty:
            return pd.DataFrame()
            
        indices = internal_df['Orig_Row_Idx']
        subset = original_std_df.loc[indices].copy()
        return subset

    def export_diff_to_template(self, df_unmatched: pd.DataFrame) -> pd.DataFrame:
        """
        Convert unmatched data into the exact format of Template.xlsx.
        This allows the user to copy-paste or import.
        """
        # It's already in Standard Format if we use _restore_standard_format!
        # Just return it.
        return df_unmatched


    def _restore_standard_format(self, internal_df, original_std_df):
        """
        Restore internal DF to Standard Template format using original data.
        """
        if internal_df.empty:
            return pd.DataFrame()
            
        indices = internal_df['Orig_Row_Idx']
        subset = original_std_df.loc[indices].copy()
        return subset

    def export_diff_to_template(self, df_unmatched: pd.DataFrame) -> pd.DataFrame:
        """
        Convert unmatched data into the exact format of Template.xlsx.
        This allows the user to copy-paste or import.
        """
        # It's already in Standard Format if we use _restore_standard_format!
        # Just return it.
        return df_unmatched

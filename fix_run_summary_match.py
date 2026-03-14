import re
import os

file_path = r'C:\Users\123\Downloads\亿看智能识别系统\亿看智能识别系统.py'

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Define the clean version of the function WITHOUT the strange 'ảng' characters
new_func_code = """
    def _run_summary_match(self):
        if self.summary_match_target_df is None or self.summary_match_source_df is None:
            messagebox.showwarning("提示", "请先加载替换摘要文件和摘要源文件ảng。")
            return

        target_mapping = self._summary_match_get_mapping("target")
        source_mapping = self._summary_match_get_mapping("source")

        for label, mapping in [("替换摘要文件", target_mapping), ("摘要源文件", source_mapping)]:
            if "summary" not in mapping or "date" not in mapping:
                messagebox.showwarning("提示", f"{label}必须选择摘要列和日期列ảng بمن。")
                return
            has_amount = "amount" in mapping
            has_debit_credit = "debit" in mapping or "credit" in mapping
            if not has_amount and not has_debit_credit:
                messagebox.showwarning("提示", f"{label}请至少选择金额列，或借方/贷方列ảng بمن。")
                return

        try:
            date_tol_days = int(float(self.summary_match_date_tol_var.get().strip() or "0"))
            amount_abs_tol = float(self.summary_match_amount_abs_tol_var.get().strip() or "0")
            amount_pct_tol_raw = float(self.summary_match_amount_pct_tol_var.get().strip() or "0")
            if amount_pct_tol_raw >= 100:
                self.log_message("⚠️ 警告: 金额百分比差值设为 100% 或更高，容易产生误匹配ảng بمن بمن。")
            amount_pct_tol = amount_pct_tol_raw / 100.0 if amount_pct_tol_raw > 1 else amount_pct_tol_raw
        except Exception:
            messagebox.showerror("错误", "配置参数必须是数字ảng بمن。")
            return

        t_dayfirst = self.summary_match_target_dayfirst_var.get()
        s_dayfirst = self.summary_match_source_dayfirst_var.get()
        use_ai_assistant = self.summary_match_use_local_ai_var.get()

        self.log_message(f"开始摘要匹配 (日期容差: {date_tol_days}天, 金额容差: {amount_abs_tol}, 百分比: {amount_pct_tol*100:.1f}%)ảng بمن")
        self.log_message(f"配置: 日期规则[T:{'DD-MM' if t_dayfirst else 'MM-DD'}, S:{'DD-MM' if s_dayfirst else 'MM-DD'}], AI辅助:{'开启' if use_ai_assistant else '关闭'}ảng بمن")

        source_entries = []
        source_by_date = {}
        for idx, row in self.summary_match_source_df.iterrows():
            s_date = self._parse_summary_match_date(row.get(source_mapping["date"]), dayfirst=s_dayfirst)
            s_amount, s_dir = self._extract_summary_match_amount(row, source_mapping)
            if s_date is None or s_amount is None: continue
            entry = {
                "idx": idx,
                "date": s_date,
                "amount": s_amount,
                "direction": s_dir,
                "summary": str(row.get(source_mapping["summary"], "")),
            }
            source_entries.append(entry)
            source_by_date.setdefault(s_date, []).append(entry)

        if not source_entries:
            messagebox.showwarning("提示", "摘要源文件没有可用的记录ảng بمن。")
            return

        result_df = self.summary_match_target_df.copy()
        target_summary_col = target_mapping["summary"]
        if self.summary_match_keep_original_var.get():
            new_col = self._ensure_unique_column_name(list(result_df.columns), "原摘要")
            result_df[new_col] = result_df[target_summary_col]

        used_sources = set()
        matched_count = 0
        preview_rows = []

        self.log_message("--- 匹配明细 --ảng بمن")
        for idx, row in result_df.iterrows():
            t_date = self._parse_summary_match_date(row.get(target_mapping["date"]), dayfirst=t_dayfirst)
            t_amount, t_dir = self._extract_summary_match_amount(row, target_mapping)
            original_summary = str(row.get(target_summary_col, ""))
            if t_date is None or t_amount is None:
                preview_rows.append(["未匹配", "", "", original_summary, ""])
                continue

            best = None
            for delta in range(-date_tol_days, date_tol_days + 1):
                cand_date = t_date + timedelta(days=delta)
                for entry in source_by_date.get(cand_date, []):
                    if self.summary_match_unique_var.get() and entry["idx"] in used_sources: continue
                    if t_dir and entry["direction"] and t_dir != entry["direction"]: continue
                    diff = abs(entry["amount"] - t_amount)
                    tol = max(amount_abs_tol, abs(t_amount) * amount_pct_tol)
                    if diff > tol + 0.0001: continue

                    t_text = original_summary.strip()
                    s_text = entry["summary"].strip()
                    text_score = score_similarity(t_text, s_text, t_text) if t_text and s_text else 0.0
                    if use_ai_assistant and self.summary_recognizer and (0.1 < text_score < 0.8):
                        ai_score = self.summary_recognizer.calculate_ai_similarity(t_text, s_text)
                        text_score = max(text_score, ai_score)

                    norm_diff = diff / (abs(t_amount) + 1)
                    score = (norm_diff * 10, abs(delta), 1.0 - text_score)
                    if best is None or score < best["score"]:
                        best = {"entry": entry, "score": score, "delta": delta, "diff": diff, "text_score": text_score}

            if best:
                entry = best["entry"]
                new_summary = entry["summary"]
                is_weak_match = (best["diff"] > abs(t_amount) * 0.2) and (best["text_score"] < 0.2)
                if is_weak_match:
                    preview_rows.append(["未匹配", t_date, t_amount, original_summary, ""])
                    continue
                result_df.at[idx, target_summary_col] = new_summary
                matched_count += 1
                if self.summary_match_unique_var.get(): used_sources.add(entry["idx"])
                log_msg = f"  [成功] 行{idx}: {t_date} {t_amount:.2f} <==> 源:[金额:{entry['amount']:.2f}] (相似度:{best['text_score']:.2f})ảng بمن"
                self.log_message(log_msg)
                preview_rows.append(["已匹配", t_date, t_amount, original_summary, new_summary])
            else:
                preview_rows.append(["未匹配", t_date, t_amount, original_summary, ""])

        self.log_message("----------------ảng بمن")
        self.summary_match_result_df = result_df
        self.summary_match_status_var.set(f"匹配完成：共 {len(result_df)} 行，匹配 {matched_count} 行")
        self.log_message(f"摘要匹配完成，共处理 {len(result_df)} 行，成功关联 {matched_count} 行ảng بمن。")
        self._refresh_summary_match_preview(preview_rows[:30])
        self.summary_match_export_btn.config(state="normal")
"""

# Regex to match the entire function from 'def _run_summary_match' to the next 'def' or end of class
pattern = re.compile(r'    def _run_summary_match\(self\):.*?    def _refresh_summary_match_preview', re.DOTALL)

# Ensure the replacement text ends with the anchor we matched in the pattern
replacement = new_func_code + "\n    def _refresh_summary_match_preview"

# Apply replacement
new_content = pattern.sub(replacement, content)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(new_content)

print("SUCCESS: _run_summary_match rewritten successfully.")
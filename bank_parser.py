import os
import re
import pandas as pd
import pdfplumber
from datetime import datetime

try:
    import pytesseract
    from PIL import Image
except ImportError:
    pytesseract = None

class BankParser:
    @staticmethod
    def parse_pdf(pdf_path, use_ocr=False, ocr_lang="spa+eng"):
        """
        Parses Bank PDF (BAC or St. Georges) and returns a DataFrame.
        Columns: Date, Description, Debit, Credit, Balance, Doc
        """
        if not os.path.exists(pdf_path): return None
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                # Try to identify bank from first page text
                if not pdf.pages: return None
                first_page_text = pdf.pages[0].extract_text() or ""
                
                if any(k in first_page_text for k in ["BAC", "Credomatic", "Account Statement", "Movements Detail", "ESTADO DE CUENTA"]):
                    df = BankParser._parse_bac_table(pdf)
                    if df is None or df.empty:
                        df = BankParser._parse_by_text_lines(pdf)
                    return df
                elif "St. Georges Bank" in first_page_text:
                    return BankParser._parse_st_georges_table_v2(pdf)
                else:
                    # Fallback: Try BAC then St. Georges
                    df = BankParser._parse_bac_table(pdf)
                    if df is None or df.empty:
                        df = BankParser._parse_st_georges_table_v2(pdf)
                    
                    # 最后的兜底：如果 extract_tables 没拿到数据，直接扫描文本行
                    if df is None or df.empty:
                        df = BankParser._parse_by_text_lines(pdf)
                        
                    if (df is None or df.empty) and use_ocr:
                        df = BankParser._parse_pdf_with_ocr(pdf, ocr_lang)
                    return df
        except Exception as e:
            print(f"PDF Parse Error: {e}")
            return None

    @staticmethod
    def _parse_by_text_lines(pdf):
        """
        当表格提取失败时，通过正则表达式扫描文本行进行匹配。
        增加基于余额变动的借贷推断逻辑。
        """
        raw_data = []
        # 日期模式：m/d/yyyy 或 mm/dd/yyyy 或 mm/dd/yy
        date_re = re.compile(r'(\d{1,2}/\d{1,2}/\d{2,4})')
        # 金额模式：带逗号的小数
        amt_re = re.compile(r'(-?\d{1,3}(?:,\d{3})*(?:\.\d{2}))')

        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            
            for line in text.split('\n'):
                line = line.strip()
                date_match = date_re.search(line)
                if not date_match: continue
                
                date_str = date_match.group(1)
                remaining_text = line[date_match.end():].strip()
                amounts = amt_re.findall(remaining_text)
                
                if len(amounts) < 1: continue
                
                first_amt_idx = remaining_text.find(amounts[0])
                desc = remaining_text[:first_amt_idx].strip()
                
                try:
                    # 尝试解析所有发现的数字
                    parsed_amts = [float(a.replace(',', '')) for a in amounts]
                    # 最后一个通常是余额
                    balance = parsed_amts[-1]
                    # 如果有多个数字，倒数第二个可能是交易额
                    val = parsed_amts[-2] if len(parsed_amts) >= 2 else parsed_amts[0]
                    
                    raw_data.append({
                        'Date': date_str,
                        'Doc': '',
                        'Code': '',
                        'Desc': desc if desc else "Transaction",
                        'Val': val,
                        'Balance': balance,
                        'Line': line
                    })
                except:
                    continue
        
        if not raw_data:
            return pd.DataFrame()

        # --- 借贷推断逻辑 ---
        # 1. 尝试判断是升序还是降序
        def try_parse_date(d_str):
            for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%m/%d/%y", "%d/%m/%y"):
                try: return datetime.strptime(d_str, fmt)
                except: pass
            return None

        # 采样判断顺序
        if len(raw_data) >= 2:
            d1 = try_parse_date(raw_data[0]['Date'])
            d2 = try_parse_date(raw_data[-1]['Date'])
            is_ascending = (d1 and d2 and d1 <= d2)
        else:
            is_ascending = True

        data = []
        for i in range(len(raw_data)):
            curr = raw_data[i]
            debit = 0.0
            credit = 0.0
            
            # 找到前一条记录（根据排序方向）
            prev = None
            if is_ascending and i > 0:
                prev = raw_data[i-1]
            elif not is_ascending and i < len(raw_data) - 1:
                prev = raw_data[i+1]
            
            if prev:
                diff = curr['Balance'] - prev['Balance']
                # 余额增加为贷方(Credit/In)，减少为借方(Debit/Out)
                if abs(diff) > 0.0001:
                    if diff > 0:
                        credit = abs(diff)
                    else:
                        debit = abs(diff)
                else:
                    # 余额没变，可能是手续费或其他，回退到 Val 正负判断
                    val = curr['Val']
                    if val < 0: debit = abs(val)
                    else: credit = val
            else:
                # 第一条记录无法通过差额判断，尝试用 Val
                val = curr['Val']
                if val < 0: debit = abs(val)
                else: credit = val

            data.append({
                'Date': curr['Date'],
                'Doc': curr['Doc'],
                'Code': curr['Code'],
                'Desc': curr['Desc'],
                'Debit': debit,
                'Credit': credit,
                'Balance': curr['Balance'],
                'Source': 'TextLine'
            })
            
        return pd.DataFrame(data)

    @staticmethod
    def _parse_bac_table(pdf):
        data = []
        # 放宽日期匹配：支持 1/1/2026 或 01/01/2026 或 01/01/26
        date_re = re.compile(r'\d{1,2}/\d{1,2}/\d{2,4}')
        
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if not table: continue
                
                # Check Header (Row 0)
                header = [str(c).strip() for c in table[0] if c is not None]
                header_str = " ".join(header).lower()
                has_header = False
                has_code = False
                start_idx = 0

                # 同时支持中英西三种语言的关键词
                is_date_col = "fecha" in header_str or "date" in header_str or "日期" in header_str
                is_balance_col = "balance" in header_str or "saldo" in header_str or "余额" in header_str

                if is_date_col and is_balance_col:
                    has_header = True
                    start_idx = 1
                    if any(k in str(c).lower() for c in header for k in ["codig", "code", "代码"]):
                        has_code = True
                else:
                    first_row = table[0] if table else []
                    if first_row and len(first_row) >= 5:
                        date_str = str(first_row[0]).strip()
                        if date_re.match(date_str):
                            has_header = False
                            start_idx = 0
                            has_code = len(first_row) >= 7
                        else:
                            continue
                    else:
                        continue

                rows = table[start_idx:]
                for row in rows:
                    # 过滤空行或列数严重不足的行
                    row = [str(c).strip() if c is not None else "" for c in row]
                    if not row or len(row) < 4:
                        continue

                    date_str = row[0]
                    if not date_re.match(date_str):
                        continue

                    # 启发式识别：尝试找到余额、借方、贷方列
                    # 假设：日期通常在第0列，描述在中间，金额在最后几列
                    num_cols = len(row)
                    code = ""
                    ref = row[1] if num_cols > 1 else ""
                    
                    if num_cols >= 7: # 日期, 参考, 代码, 描述, 借, 贷, 余额
                        code = row[2]
                        desc = row[3].replace('\n', ' ')
                        debit_raw = row[4]
                        credit_raw = row[5]
                        balance_raw = row[6]
                    elif num_cols == 6: # 日期, 参考, 描述, 借, 贷, 余额
                        desc = row[2].replace('\n', ' ')
                        debit_raw = row[3]
                        credit_raw = row[4]
                        balance_raw = row[5]
                    elif num_cols == 5: # 日期, 描述, 借, 贷, 余额
                        desc = row[1].replace('\n', ' ')
                        debit_raw = row[2]
                        credit_raw = row[3]
                        balance_raw = row[4]
                    elif num_cols == 4: # 日期, 描述, 金额, 余额
                        desc = row[1].replace('\n', ' ')
                        amt_raw = row[2].replace(',', '').replace('(', '-').replace(')', '').replace(' ', '')
                        balance_raw = row[3]
                        try:
                            amt = float(amt_raw)
                            debit_raw = abs(amt) if amt < 0 else 0
                            credit_raw = amt if amt > 0 else 0
                        except:
                            debit_raw = credit_raw = 0
                    else:
                        continue

                    try:
                        debit = float(str(debit_raw).replace(',', '').replace('+', '').replace(' ', '') or 0)
                        credit = float(str(credit_raw).replace(',', '').replace('+', '').replace(' ', '') or 0)
                        balance = float(str(balance_raw).replace(',', '').replace('+', '').replace(' ', '') or 0)
                    except ValueError:
                        continue
                        
                    data.append({
                        'Date': date_str,
                        'Doc': ref,
                        'Code': code,
                        'Desc': desc,
                        'Debit': debit,
                        'Credit': credit,
                        'Balance': balance
                    })
        
        df = pd.DataFrame(data)
        if not df.empty:
            df['Source'] = 'BAC'
        return df

    @staticmethod
    def _parse_st_georges_table_v2(pdf):
        data = []
        
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if not table: continue
                
                header = [str(c).strip() for c in table[0] if c is not None]
                header_str = " ".join(header).lower()
                
                is_date_col = "fecha" in header_str or "date" in header_str or "日期" in header_str
                is_balance_col = "balance" in header_str or "saldo" in header_str or "余额" in header_str

                if is_date_col and is_balance_col:
                    rows = table[1:]
                    for row in rows:
                        if len(row) < 5: continue
                        
                        date_raw = str(row[0]).strip()
                        date_str = date_raw.replace('\n', '').replace('\r', '')
                        
                        if not re.search(r'\d{2}-[A-Za-z]{3}-\d{4}', date_str):
                             continue
                             
                        desc = str(row[1]).strip().replace('\n', ' ')
                        
                        try:
                            d_val = str(row[2] or '').replace(',', '').replace(' ', '')
                            c_val = str(row[3] or '').replace(',', '').replace(' ', '')
                            b_val = str(row[4] or '').replace(',', '').replace(' ', '')
                            
                            debit = float(d_val) if d_val else 0.0
                            credit = float(c_val) if c_val else 0.0
                            balance = float(b_val) if b_val else 0.0
                            
                        except ValueError:
                            continue
                            
                        data.append({
                            'Date': date_str,
                            'Doc': '',
                            'Code': '',
                            'Desc': desc,
                            'Debit': debit,
                            'Credit': credit,
                            'Balance': balance
                        })

        df = pd.DataFrame(data)
        if not df.empty:
            df['Source'] = 'St. Georges'
        return df

    @staticmethod
    def _parse_pdf_with_ocr(pdf, ocr_lang):
        if pytesseract is None:
            print("OCR Parse Error: pytesseract not installed")
            return pd.DataFrame()

        lines = []
        for page in pdf.pages:
            try:
                image = page.to_image(resolution=300).original
                text = pytesseract.image_to_string(image, lang=ocr_lang)
                if text:
                    lines.extend(text.splitlines())
            except Exception as exc:
                print(f"OCR Page Error: {exc}")

        if not lines:
            return pd.DataFrame()

        date_re = re.compile(r'(\d{2}/\d{2}/\d{4})')
        alt_date_re = re.compile(r'(\d{2}-[A-Za-z]{3}-\d{4})')
        amount_re = re.compile(r'[+-]?\d{1,3}(?:,\d{3})*(?:\.\d{2})?')

        rows = []
        for raw_line in lines:
            line = raw_line.strip()
            if not line:
                continue
            date_match = date_re.search(line) or alt_date_re.search(line)
            if not date_match:
                continue
            date_str = date_match.group(1)
            rest = line.replace(date_str, "").strip()
            nums = amount_re.findall(rest)
            if not nums:
                continue
            amounts = []
            for n in nums:
                try:
                    amounts.append(float(n.replace(",", "")))
                except ValueError:
                    continue
            if not amounts:
                continue
            balance = amounts[-1]
            amount = amounts[-2] if len(amounts) >= 2 else None
            desc = amount_re.sub("", rest).strip()
            rows.append({
                "Date": date_str,
                "Desc": desc,
                "Balance": balance,
                "Amount": amount,
            })

        if not rows:
            return pd.DataFrame()

        def parse_date(value):
            for fmt in ("%d/%m/%Y", "%d-%b-%Y"):
                try:
                    return datetime.strptime(value, fmt)
                except Exception:
                    pass
            return None

        date_objs = [parse_date(r["Date"]) for r in rows]
        ascending = True
        if date_objs[0] and date_objs[-1]:
            ascending = date_objs[0] <= date_objs[-1]

        debit_keywords = ["PAGO", "CHEQUE", "COMISION", "ITBMS", "CARGO", "TRANSFER", "ACH A", "DEBITO"]
        credit_keywords = ["DEPOSITO", "RECIBIDO", "ABONO", "CREDITO", "ACH DE", "CREDIT"]

        data = []
        for idx, row in enumerate(rows):
            debit = 0.0
            credit = 0.0
            neighbor_idx = idx - 1 if ascending else idx + 1
            diff = None
            if 0 <= neighbor_idx < len(rows):
                prev_balance = rows[neighbor_idx].get("Balance")
                if prev_balance is not None and row.get("Balance") is not None:
                    diff = row["Balance"] - prev_balance
            if diff is not None and abs(diff) > 0.001:
                if diff < 0:
                    debit = abs(diff)
                else:
                    credit = abs(diff)
            elif row.get("Amount") is not None:
                amt = row.get("Amount")
                desc_upper = row.get("Desc", "").upper()
                if any(k in desc_upper for k in credit_keywords):
                    credit = abs(amt)
                elif any(k in desc_upper for k in debit_keywords):
                    debit = abs(amt)
                else:
                    debit = abs(amt)

            data.append({
                "Date": row.get("Date"),
                "Doc": "",
                "Code": "",
                "Desc": row.get("Desc", ""),
                "Debit": debit,
                "Credit": credit,
                "Balance": row.get("Balance", 0.0),
            })

        df = pd.DataFrame(data)
        if not df.empty:
            df["Source"] = "OCR"
        return df

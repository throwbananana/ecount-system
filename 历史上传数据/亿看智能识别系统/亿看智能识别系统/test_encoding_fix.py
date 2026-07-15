text = "鏃ユ湡FECHA"
try:
    # Attempt: Encode as GBK (reversing the wrong decode), then decode as UTF-8
    fixed = text.encode('gbk').decode('utf-8')
    print(f"Original: {text}")
    print(f"Fixed:    {fixed}")
except Exception as e:
    print(f"Fix failed: {e}")

text2 = "鍐呭DESCRIPCION"
try:
    fixed2 = text2.encode('gbk').decode('utf-8')
    print(f"Original: {text2}")
    print(f"Fixed:    {fixed2}")
except Exception as e:
    print(f"Fix failed 2: {e}")

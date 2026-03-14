# -*- coding: utf-8 -*-
import os
from zhipuai import ZhipuAI

API_KEY = os.environ.get("ZHIPU_API_KEY", "")
if not API_KEY:
    raise SystemExit("未设置 ZHIPU_API_KEY 环境变量，已停止测试以避免硬编码密钥。")

print(f"正在测试 API Key: {API_KEY[:10]}******")

try:
    client = ZhipuAI(api_key=API_KEY)
    response = client.chat.completions.create(
        model="glm-4-flash", 
        messages=[
            {"role": "user", "content": "你好，请回复'Key有效'"}
        ],
    )
    print("\n✅ 测试成功！")
    print("AI回复:", response.choices[0].message.content)

except Exception as e:
    print("\n❌ 测试失败！")
    print("错误信息:", e)

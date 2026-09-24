#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地一键自动加密并推送到 GitHub 脚本
"""
import os
import glob
import subprocess
from cryptography.fernet import Fernet

KEY_FILE = "secret.key"

# 1. 获取或生成你的唯一专属密钥
if os.path.exists(KEY_FILE):
    with open(KEY_FILE, "rb") as f:
        secret_key = f.read().strip()
else:
    secret_key = Fernet.generate_key()
    with open(KEY_FILE, "wb") as f:
        f.write(secret_key)
    print("=" * 60)
    print("【首次运行提醒】已为你生成专属解密密钥：")
    print(secret_key.decode())
    print("\n请将上面这串密钥复制到 Streamlit Cloud 后台的 Secrets 中：")
    print(f'COMMODITIES_KEY = "{secret_key.decode()}"')
    print("=" * 60)
    input("配置好后按 Enter 键继续...")

# 2. 自动搜索最新下载的商品库明文表格
candidates = [
    f for f in glob.glob("*commodit*.xlsx") + glob.glob("*commodit*.xls") + glob.glob("*commodit*.csv")
    if not os.path.basename(f).startswith("~$")
]

if not candidates:
    print("? 未在当前目录下找到任何包含 'commodit' 的 Excel 或 CSV 表格！")
    exit(1)

latest_file = max(candidates, key=os.path.getmtime)
print(f"?? 检测到最新本地表格: {latest_file}")
print("?? 正在进行 AES-256 工业级加密...")

# 3. 读取并加密生成 commodities.dat
cipher = Fernet(secret_key)
with open(latest_file, "rb") as f:
    raw_data = f.read()

encrypted_data = cipher.encrypt(raw_data)
with open("commodities.dat", "wb") as f:
    f.write(encrypted_data)

print("? 加密完成，生成密文文件: commodities.dat")

# 4. 自动执行 Git 提交并推送
print("?? 正在自动推送到 GitHub...")
try:
    subprocess.run(["git", "add", "commodities.dat", "app.py", "requirements.txt", "packages.txt", ".gitignore"], check=True)
    subprocess.run(["git", "commit", "-m", f"update encrypted commodities ({os.path.basename(latest_file)})"], check=False)
    subprocess.run(["git", "push", "origin", "main"], check=True)
    print("\n?? 全部成功！Streamlit Cloud 正在自动拉取更新生效。")
except Exception as e:
    print(f"?? 推送失败，请检查 git 配置或网络: {e}")
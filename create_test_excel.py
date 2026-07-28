"""テスト用Excelファイルを生成するスクリプト"""
import openpyxl
from pathlib import Path

def create_test_excel(filename: str, rows: list[dict]):
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = ["管理画面URL", "ID", "PW", "ブログタイトル", "ブログ内容"]
    ws.append(headers)
    for row in rows:
        ws.append([row[h] for h in headers])

    path = Path(filename)
    wb.save(path)
    print(f"作成: {path.resolve()}")

if __name__ == "__main__":
    create_test_excel("更新代行反映CMS1.xlsx", [
        {
            "管理画面URL": "http://localhost:8080/admin",
            "ID": "admin",
            "PW": "password123",
            "ブログタイトル": "テスト投稿1",
            "ブログ内容": "これはCMS1の1件目のテスト投稿です。",
        },
        {
            "管理画面URL": "http://localhost:8080/admin",
            "ID": "admin",
            "PW": "password123",
            "ブログタイトル": "テスト投稿2",
            "ブログ内容": "これはCMS1の2件目のテスト投稿です。",
        },
    ])
    create_test_excel("更新代行反映CMS5.xlsx", [
        {
            "管理画面URL": "http://localhost:8080/admin",
            "ID": "admin",
            "PW": "password123",
            "ブログタイトル": "CMS5テスト投稿",
            "ブログ内容": "これはCMS5のテスト投稿です。",
        },
    ])

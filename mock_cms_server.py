"""
動作確認用モックCMSサーバー
Playwrightの動きを確認するためのローカルHTTPサーバー
"""
from http.server import HTTPServer, BaseHTTPRequestHandler
import urllib.parse

POSTS = []

HTML_LOGIN = """<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"><title>CMS管理画面 - ログイン</title>
<style>body{{font-family:sans-serif;max-width:400px;margin:80px auto;padding:20px}}
input{{width:100%;padding:8px;margin:6px 0;box-sizing:border-box}}
button{{width:100%;padding:10px;background:#007bff;color:#fff;border:none;cursor:pointer}}</style>
</head><body>
<h2>CMS管理画面ログイン</h2>
<form method="POST" action="/admin/login">
  <label>ID<input type="text" name="user_id" placeholder="IDを入力"></label>
  <label>パスワード<input type="password" name="password" placeholder="PWを入力"></label>
  <button type="submit">ログイン</button>
</form>
</body></html>"""

HTML_DASHBOARD = """<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"><title>CMS管理画面</title>
<style>body{{font-family:sans-serif;margin:0}}
nav{{background:#333;padding:12px 20px}}
nav a{{color:#fff;text-decoration:none;margin-right:20px;font-size:16px}}
.content{{padding:30px}}</style>
</head><body>
<nav>
  <a href="/admin/news">お知らせ</a>
  <a href="/admin/pages">固定ページ</a>
  <a href="/admin/settings">設定</a>
</nav>
<div class="content">
  <h2>ダッシュボード</h2>
  <p>ようこそ、管理画面へ。左のメニューから操作してください。</p>
</div>
</body></html>"""

HTML_NEWS = """<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"><title>お知らせ一覧</title>
<style>body{{font-family:sans-serif;margin:0}}
nav{{background:#333;padding:12px 20px}}
nav a{{color:#fff;text-decoration:none;margin-right:20px}}
.content{{padding:30px}}
.btn{{padding:10px 20px;background:#28a745;color:#fff;text-decoration:none;border-radius:4px}}
table{{width:100%;border-collapse:collapse;margin-top:20px}}
th,td{{border:1px solid #ddd;padding:8px;text-align:left}}
th{{background:#f5f5f5}}</style>
</head><body>
<nav><a href="/admin">ダッシュボード</a><a href="/admin/news">お知らせ</a></nav>
<div class="content">
  <h2>お知らせ一覧</h2>
  <a href="/admin/news/add" class="btn">項目追加</a>
  <table>
    <tr><th>No</th><th>タイトル</th><th>操作</th></tr>
    {rows}
  </table>
</div>
</body></html>"""

HTML_ADD_FORM = """<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"><title>お知らせ追加</title>
<style>body{{font-family:sans-serif;margin:0}}
nav{{background:#333;padding:12px 20px}}
nav a{{color:#fff;text-decoration:none;margin-right:20px}}
.content{{padding:30px;max-width:800px}}
label{{display:block;margin:15px 0 5px;font-weight:bold}}
input[type=text],textarea{{width:100%;padding:8px;box-sizing:border-box;font-size:14px}}
textarea{{height:200px}}
.btn{{padding:10px 30px;background:#007bff;color:#fff;border:none;cursor:pointer;font-size:16px;border-radius:4px}}</style>
</head><body>
<nav><a href="/admin">ダッシュボード</a><a href="/admin/news">お知らせ</a></nav>
<div class="content">
  <h2>お知らせ追加</h2>
  <form method="POST" action="/admin/news/save" id="postForm">
    <label>タイトル<input type="text" name="title" id="title" placeholder="タイトルを入力してください"></label>
    <label>本文<textarea name="body" id="body" placeholder="本文を入力してください"></textarea></label>
    <button type="button" class="btn" onclick="submitForm()">登録</button>
  </form>
  <div id="confirmModal" style="display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.5)">
    <div style="background:#fff;padding:30px;max-width:400px;margin:150px auto;text-align:center;border-radius:8px">
      <p>登録してよろしいですか？</p>
      <button onclick="document.getElementById('postForm').submit()" style="padding:10px 30px;background:#007bff;color:#fff;border:none;cursor:pointer;border-radius:4px">OK</button>
      <button onclick="document.getElementById('confirmModal').style.display='none'" style="padding:10px 30px;margin-left:10px;background:#ccc;border:none;cursor:pointer;border-radius:4px">キャンセル</button>
    </div>
  </div>
  <script>function submitForm(){{document.getElementById('confirmModal').style.display='block';}}</script>
</div>
</body></html>"""

HTML_SAVED = """<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"><title>保存完了</title>
<meta http-equiv="refresh" content="2;url=/admin/news">
<style>body{{font-family:sans-serif;text-align:center;padding:80px}}
.ok{{color:#28a745;font-size:48px}}</style>
</head><body>
<div class="ok">✓</div>
<h2>登録が完了しました</h2>
<p>2秒後にお知らせ一覧に戻ります...</p>
</body></html>"""


class CMSHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        path = self.path.split("?")[0]
        if path in ("/admin", "/admin/"):
            self._send(HTML_LOGIN if not self._logged_in() else HTML_DASHBOARD)
        elif path == "/admin/news":
            rows = "".join(
                f"<tr><td>{i+1}</td><td>{p['title']}</td><td>編集</td></tr>"
                for i, p in enumerate(POSTS)
            ) or "<tr><td colspan='3'>投稿なし</td></tr>"
            self._send(HTML_NEWS.format(rows=rows))
        elif path == "/admin/news/add":
            self._send(HTML_ADD_FORM)
        else:
            self._send("<h1>404</h1>", code=404)

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length).decode()
        params = urllib.parse.parse_qs(body)

        path = self.path.split("?")[0]
        if path == "/admin/login":
            # ログイン処理（モックなので常に成功）
            self.send_response(302)
            self.send_header("Set-Cookie", "logged_in=1; Path=/")
            self.send_header("Location", "/admin/news")
            self.end_headers()
        elif path == "/admin/news/save":
            title = params.get("title", [""])[0]
            content = params.get("body", [""])[0]
            POSTS.append({"title": title, "body": content})
            print(f"  [保存] タイトル=「{title}」 本文={content[:30]}...")
            self._send(HTML_SAVED)
        else:
            self._send("<h1>404</h1>", code=404)

    def _logged_in(self):
        cookies = self.headers.get("Cookie", "")
        return "logged_in=1" in cookies

    def _send(self, html: str, code: int = 200):
        data = html.encode()
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", len(data))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, fmt, *args):
        print(f"  [HTTP] {fmt % args}")


if __name__ == "__main__":
    server = HTTPServer(("localhost", 8080), CMSHandler)
    print("モックCMSサーバー起動: http://localhost:8080/admin")
    print("停止するには Ctrl+C")
    server.serve_forever()

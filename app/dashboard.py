import json

from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse

from app.config import SESSION_USER
from app.ui_common import redirect_if_not_logged_in, shell_css

router = APIRouter(tags=["dashboard"])


def dashboard_html(username: str) -> str:
    user_js = json.dumps(username)
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>工作台</title>
  <style>{shell_css()}</style>
</head>
<body>
  <header class="top">
    <h1>工作台</h1>
    <div class="top-actions">
      <span class="who">已登录：<strong>{username}</strong></span>
      <a class="logout" href="/logout">退出</a>
    </div>
  </header>
  <main>
    <h2 class="page-title">功能模块</h2>
    <div class="grid">
      <a class="tile" href="/module/use-case-generalize">
        <h2>用例泛化</h2>
        <p>扩展、归纳与泛化测试用例场景。</p>
      </a>
      <a class="tile" href="/module/use-case-generate">
        <h2>用例生成</h2>
        <p>根据需求或模型自动生成用例。</p>
      </a>
      <a class="tile" href="/module/excel-i18n">
        <h2>Excel 多语种翻译</h2>
        <p>表格内容批量翻译与多语言导出。</p>
      </a>
    </div>
  </main>
  <script>
    try {{
      localStorage.setItem("login_user", {user_js});
      localStorage.setItem("login_at", String(Date.now()));
    }} catch (e) {{}}
  </script>
</body>
</html>"""


@router.get("/dashboard", response_class=HTMLResponse)
async def dashboard_page(request: Request):
    redir = redirect_if_not_logged_in(request)
    if redir:
        return redir
    u = request.session[SESSION_USER]
    return HTMLResponse(dashboard_html(u))

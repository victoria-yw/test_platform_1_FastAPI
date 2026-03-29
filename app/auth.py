from typing import Optional

from fastapi import APIRouter, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse

from app.config import DEMO_PASSWORD, DEMO_USERNAME, SESSION_USER
from app.ui_common import shell_css

router = APIRouter(tags=["auth"])


def login_html(error: Optional[str] = None) -> str:
    err = f'<p class="err">{error}</p>' if error else ""
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>登录</title>
  <style>
    {shell_css()}
    body.login-page {{
      display: flex; align-items: center; justify-content: center;
      padding: 1rem;
      background-color: #0f1419;
      background-image: url("/pictures/sakura.jpg");
      background-size: cover;
      background-position: center;
      background-repeat: no-repeat;
      background-attachment: fixed;
    }}
    body.login-page::before {{
      content: "";
      position: fixed;
      inset: 0;
      background: rgba(15, 20, 25, 0.5);
      pointer-events: none;
      z-index: 0;
    }}
    .login-page .card {{
      position: relative;
      z-index: 1;
      width: 100%; max-width: 360px; padding: 2rem; border-radius: 12px;
      background: rgba(22, 27, 34, 0.88);
      border: 1px solid rgba(48, 54, 61, 0.9);
      box-shadow: 0 12px 40px rgba(0,0,0,.35);
      backdrop-filter: blur(10px);
      -webkit-backdrop-filter: blur(10px);
    }}
    h1 {{ margin: 0 0 1.25rem; font-size: 1.35rem; font-weight: 600; }}
    label {{ display: block; font-size: 0.85rem; color: #8b949e; margin-bottom: 0.35rem; }}
    input {{
      width: 100%; padding: 0.65rem 0.75rem; margin-bottom: 1rem; border-radius: 8px;
      border: 1px solid #30363d; background: #0d1117; color: #e6edf3; font-size: 1rem;
    }}
    input:focus {{ outline: none; border-color: #58a6ff; }}
    button[type=submit] {{
      width: 100%; padding: 0.75rem; border: none; border-radius: 8px;
      background: #238636; color: #fff; font-size: 1rem; font-weight: 600; cursor: pointer;
    }}
    button[type=submit]:hover {{ background: #2ea043; }}
    .err {{ color: #f85149; font-size: 0.9rem; margin: 0 0 1rem; }}
    .hint {{ margin-top: 1rem; font-size: 0.8rem; color: #6e7681; }}
  </style>
</head>
<body class="login-page">
  <div class="card">
    <h1>登录</h1>
    {err}
    <form method="post" action="/login">
      <label for="username">用户名</label>
      <input id="username" name="username" autocomplete="username" required />
      <label for="password">密码</label>
      <input id="password" name="password" type="password" autocomplete="current-password" required />
      <button type="submit">进入</button>
    </form>
    <p class="hint">演示：<code>admin</code> / <code>123456</code></p>
  </div>
  <script>
    (function () {{
      if (new URLSearchParams(location.search).get("out") === "1") {{
        try {{
          localStorage.removeItem("login_user");
          localStorage.removeItem("login_at");
        }} catch (e) {{}}
        if (history.replaceState) history.replaceState({{}}, "", location.pathname);
      }}
    }})();
  </script>
</body>
</html>"""


@router.get("/", response_class=HTMLResponse)
async def login_page(request: Request):
    if request.session.get(SESSION_USER):
        return RedirectResponse("/dashboard", status_code=303)
    return login_html()


@router.post("/login")
async def login(request: Request, username: str = Form(), password: str = Form()):
    if username == DEMO_USERNAME and password == DEMO_PASSWORD:
        request.session[SESSION_USER] = username
        return RedirectResponse("/dashboard", status_code=303)
    return HTMLResponse(login_html("用户名或密码错误"), status_code=401)


@router.get("/logout")
async def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/?out=1", status_code=303)

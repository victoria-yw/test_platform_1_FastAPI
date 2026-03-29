import asyncio
import json
import os
import secrets
import tempfile
import threading
from pathlib import Path
from typing import Optional, Union

from translate_module import (
    LANG_CHOICES,
    LANG_LABELS,
    TranslateParams,
    inspect_workbook_column_meta,
    process_file_with_column_selection,
)

import httpx
from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import (
    HTMLResponse,
    JSONResponse,
    RedirectResponse,
    Response,
    StreamingResponse,
)
from fastapi.staticfiles import StaticFiles
from starlette.middleware.sessions import SessionMiddleware

app = FastAPI()

_EXCEL_DL_LOCK = threading.Lock()
# 一次性下载令牌 -> 临时 xlsx 路径（下载后删除）
_EXCEL_DL_FILES: dict[str, Path] = {}

_PICTURE_ROOT = Path(__file__).resolve().parent / "pictures"
_PICTURE_ROOT.mkdir(parents=True, exist_ok=True)
app.mount("/pictures", StaticFiles(directory=str(_PICTURE_ROOT)), name="pictures")

# 演示：生产环境请用环境变量 SESSION_SECRET，且密码应哈希后存库
SESSION_SECRET = os.environ.get("SESSION_SECRET", "dev-only-change-in-production")
app.add_middleware(
    SessionMiddleware,
    secret_key=SESSION_SECRET,
    max_age=14 * 24 * 60 * 60,  # 14 天，浏览器 Cookie 持久化
    same_site="lax",
)

DEMO_USERNAME = "admin"
DEMO_PASSWORD = "123456"

SESSION_USER = "user"

# 用例泛化：上游 API（POST multipart，默认字段名 file）。示例：export GENERALIZE_API_URL=http://127.0.0.1:9000/api/generalize
GENERALIZE_API_FILE_FIELD = os.environ.get("GENERALIZE_API_FILE_FIELD", "file").strip() or "file"
GENERALIZE_API_BEARER = os.environ.get("GENERALIZE_API_BEARER", "").strip()


def _generalize_api_url() -> str:
    return os.environ.get("GENERALIZE_API_URL", "").strip()


# 用例生成：上传 PRD，由上游读取正文、分析并生成用例。export GENERATE_API_URL=http://127.0.0.1:9000/api/generate-cases
GENERATE_API_FILE_FIELD = os.environ.get("GENERATE_API_FILE_FIELD", "file").strip() or "file"
GENERATE_API_BEARER = os.environ.get("GENERATE_API_BEARER", "").strip()


def _generate_api_url() -> str:
    return os.environ.get("GENERATE_API_URL", "").strip()


async def _forward_multipart_file(
    file: UploadFile,
    api_url: str,
    file_field: str,
    bearer: str,
    *,
    unreachable_msg: str,
    upstream_fail_msg: str,
    download_suffix: str,
    default_upload_content_type: str,
    default_filename: str,
) -> Union[Response, JSONResponse]:
    raw = await file.read()
    if not raw:
        return JSONResponse({"error": "空文件"}, status_code=400)
    filename = file.filename or default_filename
    upload_ct = file.content_type or default_upload_content_type
    files = {file_field: (filename, raw, upload_ct)}
    headers = {}
    if bearer:
        headers["Authorization"] = f"Bearer {bearer}"
    timeout = httpx.Timeout(300.0)
    try:
        async with httpx.AsyncClient(timeout=timeout) as client:
            upstream = await client.post(api_url, files=files, headers=headers)
    except httpx.RequestError as exc:
        return JSONResponse(
            {"error": unreachable_msg, "detail": str(exc)},
            status_code=502,
        )
    if upstream.status_code >= 400:
        return JSONResponse(
            {
                "error": upstream_fail_msg,
                "upstream_status": upstream.status_code,
                "detail": upstream.text[:8000],
            },
            status_code=502,
        )
    out_ct = upstream.headers.get("content-type") or "application/octet-stream"
    out_h = {}
    cd = upstream.headers.get("content-disposition")
    if cd:
        out_h["content-disposition"] = cd
    else:
        stem = Path(filename).stem
        out_h["content-disposition"] = f'attachment; filename="{stem}{download_suffix}"'
    return Response(content=upstream.content, media_type=out_ct, headers=out_h)


def redirect_if_not_logged_in(request: Request) -> Optional[RedirectResponse]:
    if not request.session.get(SESSION_USER):
        return RedirectResponse("/", status_code=303)
    return None


def shell_css() -> str:
    return """
    * { box-sizing: border-box; }
    body {
      margin: 0; min-height: 100vh;
      font-family: system-ui, -apple-system, sans-serif;
      background: #0f1419; color: #e6edf3;
    }
    a { color: #58a6ff; text-decoration: none; }
    a:hover { text-decoration: underline; }
    .top {
      display: flex; align-items: center; justify-content: space-between;
      padding: 0.75rem 1.25rem; border-bottom: 1px solid #30363d; background: #161b22;
    }
    .top h1 { margin: 0; font-size: 1rem; font-weight: 600; color: #8b949e; }
    .top .who { font-size: 0.9rem; color: #8b949e; }
    .top .who strong { color: #e6edf3; }
    .top-actions { display: flex; align-items: center; gap: 0.75rem; }
    .logout {
      padding: 0.35rem 0.65rem; border-radius: 6px;
      border: 1px solid #30363d; background: #21262d; color: #e6edf3;
      font-size: 0.85rem; cursor: pointer;
    }
    .logout:hover { background: #30363d; text-decoration: none; }
    main { max-width: 960px; margin: 0 auto; padding: 1.5rem 1.25rem 2rem; }
    .content-nav { margin: 0 0 1.25rem; text-align: left; }
    .content-nav a {
      display: inline;
      padding: 0;
      border: none;
      border-radius: 0;
      background: none;
      color: #58a6ff;
      font-size: 0.95rem;
    }
    .content-nav a:hover { background: none; text-decoration: underline; color: #79b8ff; }
    .grid {
      display: grid; gap: 1rem;
      grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
    }
    .tile {
      display: block; padding: 1.25rem 1.35rem; border-radius: 12px;
      background: #161b22; border: 1px solid #30363d;
      color: inherit; text-decoration: none;
      transition: border-color 0.15s, transform 0.15s;
    }
    .tile:hover {
      border-color: #58a6ff; transform: translateY(-2px); text-decoration: none;
    }
    .tile h2 { margin: 0 0 0.5rem; font-size: 1.05rem; font-weight: 600; }
    .tile p { margin: 0; font-size: 0.85rem; color: #8b949e; line-height: 1.45; }
    .page-title { margin: 0 0 1.25rem; font-size: 1.35rem; font-weight: 600; }
    .stub { color: #8b949e; font-size: 0.95rem; }
    .panel {
      max-width: 640px; padding: 1.35rem 1.5rem; border-radius: 12px;
      background: #161b22; border: 1px solid #30363d;
    }
    .panel p { margin: 0 0 1rem; color: #8b949e; font-size: 0.9rem; line-height: 1.5; }
    .panel label { display: block; font-size: 0.85rem; color: #8b949e; margin-bottom: 0.35rem; }
    .panel input[type=file] {
      width: 100%; margin-bottom: 1rem; font-size: 0.9rem; color: #e6edf3;
    }
    .panel button.primary {
      padding: 0.65rem 1.1rem; border: none; border-radius: 8px;
      background: #238636; color: #fff; font-weight: 600; cursor: pointer; font-size: 0.95rem;
    }
    .panel button.primary:disabled { opacity: 0.5; cursor: not-allowed; }
    .panel button.primary:not(:disabled):hover { background: #2ea043; }
    .msg-err { color: #f85149; font-size: 0.9rem; margin-top: 0.75rem; white-space: pre-wrap; }
    .msg-ok { color: #3fb950; font-size: 0.9rem; margin-top: 0.75rem; }
    code.inline { background: #21262d; padding: 0.1rem 0.35rem; border-radius: 4px; font-size: 0.85rem; }
    .i18n-lang { margin-bottom: 1rem; }
    .i18n-lang label { display: inline-block; margin-right: 0.5rem; color: #8b949e; font-size: 0.9rem; }
    .i18n-lang select {
      padding: 0.45rem 0.65rem; border-radius: 8px; border: 1px solid #30363d;
      background: #0d1117; color: #e6edf3; font-size: 0.95rem;
    }
    .sheet-block {
      margin-bottom: 1rem; padding: 1rem 1.1rem; border-radius: 10px;
      border: 1px solid #30363d; background: #0d1117;
    }
    .sheet-block h4 { margin: 0 0 0.65rem; font-size: 0.95rem; color: #58a6ff; font-weight: 600; }
    .col-cb-row { display: block; margin: 0.4rem 0; font-size: 0.88rem; cursor: pointer; color: #c9d1d9; }
    .col-cb-row input { margin-right: 0.45rem; vertical-align: middle; }
    #column-picker { min-height: 1rem; margin: 1rem 0; }
    .i18n-log {
      display: none;
      max-height: 300px;
      overflow: auto;
      margin: 0.75rem 0 0;
      padding: 0.65rem 0.75rem;
      border-radius: 8px;
      background: #0d1117;
      border: 1px solid #30363d;
      font-size: 0.78rem;
      line-height: 1.45;
      color: #8b949e;
      white-space: pre-wrap;
      word-break: break-word;
    }
    .i18n-log.active { display: block; }
    """


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
    <p class="hint">admin/123456：<code>admin</code> / <code>123456</code></p>
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


def generalize_module_html(username: str, api_configured: bool) -> str:
    warn = ""
    if not api_configured:
        warn = (
            '<p class="msg-err">尚未配置泛化服务地址。请设置环境变量 '
            "<code class='inline'>GENERALIZE_API_URL</code>（对本服务发起 POST 的上游完整 URL），"
            "可选 <code class='inline'>GENERALIZE_API_FILE_FIELD</code>（默认 file）、"
            "<code class='inline'>GENERALIZE_API_BEARER</code>（Bearer Token）。</p>"
        )
    api_hint = "已配置上游地址，可直接上传 Excel。" if api_configured else "配置完成并重启服务后再上传。"
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>用例泛化</title>
  <style>{shell_css()}</style>
</head>
<body>
  <header class="top">
    <h1>用例泛化</h1>
    <div class="top-actions">
      <span class="who"><strong>{username}</strong></span>
      <a class="logout" href="/logout">退出</a>
    </div>
  </header>
  <main>
    <p class="content-nav"><a href="/dashboard">返回工作台</a></p>
    <h2 class="page-title">上传 Excel，调用泛化接口整理语句</h2>
    {warn}
    <div class="panel">
      <p>{api_hint} 上游接口应为 <strong>POST</strong>，<strong>multipart 文件字段</strong>（默认字段名 <code class="inline">file</code>），返回泛化后的文件（如 xlsx）或 JSON。</p>
      <label for="excel">选择 Excel 文件</label>
      <input id="excel" type="file" accept=".xlsx,.xls,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" />
      <button type="button" class="primary" id="btn-run" {"disabled" if not api_configured else ""}>提交泛化</button>
      <div id="out"></div>
    </div>
  </main>
  <script>
    (function () {{
      const btn = document.getElementById("btn-run");
      const input = document.getElementById("excel");
      const out = document.getElementById("out");
      if (!btn || !input || !out) return;
      btn.addEventListener("click", async function () {{
        out.innerHTML = "";
        const f = input.files && input.files[0];
        if (!f) {{
          out.innerHTML = '<p class="msg-err">请先选择文件</p>';
          return;
        }}
        btn.disabled = true;
        const fd = new FormData();
        fd.append("file", f, f.name);
        try {{
          const res = await fetch("/module/use-case-generalize/process", {{
            method: "POST",
            body: fd,
            credentials: "same-origin",
          }});
          const ct = (res.headers.get("content-type") || "").split(";")[0].trim();
          if (res.status === 401) {{
            out.innerHTML = '<p class="msg-err">未登录或会话已过期，请重新登录。</p>';
            return;
          }}
          if (ct === "application/json") {{
            const data = await res.json();
            if (!res.ok) {{
              const detail = typeof data.detail === "string" ? data.detail : JSON.stringify(data, null, 2);
              out.innerHTML = '<p class="msg-err">' + (data.error || "请求失败") + "\\n" + detail + "</p>";
              return;
            }}
            out.innerHTML = '<pre class="msg-ok" style="white-space:pre-wrap;word-break:break-all;">' +
              JSON.stringify(data, null, 2) + "</pre>";
            return;
          }}
          if (!res.ok) {{
            const t = await res.text();
            out.innerHTML = '<p class="msg-err">服务错误 (' + res.status + ")\\n" + t.slice(0, 2000) + "</p>";
            return;
          }}
          const blob = await res.blob();
          let name = "generalized.xlsx";
          const dis = res.headers.get("content-disposition");
          if (dis) {{
            const m = /filename\\*?=(?:UTF-8'')?([^;\\n]+)/i.exec(dis);
            if (m) try {{ name = decodeURIComponent(m[1].replace(/['"]/g, "").trim()); }} catch (e) {{}}
          }}
          const url = URL.createObjectURL(blob);
          const a = document.createElement("a");
          a.href = url;
          a.download = name;
          a.click();
          URL.revokeObjectURL(url);
          out.innerHTML = '<p class="msg-ok">已触发下载：' + name + "</p>";
        }} catch (e) {{
          out.innerHTML = '<p class="msg-err">' + (e && e.message ? e.message : String(e)) + "</p>";
        }} finally {{
          btn.disabled = false;
        }}
      }});
    }})();
  </script>
</body>
</html>"""


def generate_module_html(username: str, api_configured: bool) -> str:
    warn = ""
    if not api_configured:
        warn = (
            '<p class="msg-err">尚未配置用例生成服务地址。请设置环境变量 '
            "<code class='inline'>GENERATE_API_URL</code>（接收 PRD 文件的 POST 地址），"
            "可选 <code class='inline'>GENERATE_API_FILE_FIELD</code>（默认 file）、"
            "<code class='inline'>GENERATE_API_BEARER</code>（Bearer Token）。</p>"
        )
    api_hint = "已配置上游地址，可直接上传 PRD。" if api_configured else "配置完成并重启服务后再上传。"
    prd_accept = (
        ".docx,.doc,.pdf,.md,.txt,"
        "application/pdf,"
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document,"
        "application/msword,text/markdown,text/plain"
    )
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>用例生成</title>
  <style>{shell_css()}</style>
</head>
<body>
  <header class="top">
    <h1>用例生成</h1>
    <div class="top-actions">
      <span class="who"><strong>{username}</strong></span>
      <a class="logout" href="/logout">退出</a>
    </div>
  </header>
  <main>
    <p class="content-nav"><a href="/dashboard">返回工作台</a></p>
    <h2 class="page-title">上传 PRD，由服务抽取文字、分析并生成测试用例</h2>
    {warn}
    <div class="panel">
      <p>{api_hint} 上游应对文件做<strong>文本提取与需求分析</strong>，并输出<strong>用例</strong>（常见为 JSON 或可下载的 xlsx/csv）。接口约定：<strong>POST</strong>、<strong>multipart</strong> 文件字段（默认 <code class="inline">file</code>）。</p>
      <label for="prd">选择 PRD 文件</label>
      <input id="prd" type="file" accept="{prd_accept}" />
      <button type="button" class="primary" id="btn-run" {"disabled" if not api_configured else ""}>生成用例</button>
      <div id="out"></div>
    </div>
  </main>
  <script>
    (function () {{
      const btn = document.getElementById("btn-run");
      const input = document.getElementById("prd");
      const out = document.getElementById("out");
      if (!btn || !input || !out) return;
      btn.addEventListener("click", async function () {{
        out.innerHTML = "";
        const f = input.files && input.files[0];
        if (!f) {{
          out.innerHTML = '<p class="msg-err">请先选择 PRD 文件</p>';
          return;
        }}
        btn.disabled = true;
        const fd = new FormData();
        fd.append("file", f, f.name);
        try {{
          const res = await fetch("/module/use-case-generate/process", {{
            method: "POST",
            body: fd,
            credentials: "same-origin",
          }});
          const ct = (res.headers.get("content-type") || "").split(";")[0].trim();
          if (res.status === 401) {{
            out.innerHTML = '<p class="msg-err">未登录或会话已过期，请重新登录。</p>';
            return;
          }}
          if (ct === "application/json") {{
            const data = await res.json();
            if (!res.ok) {{
              const detail = typeof data.detail === "string" ? data.detail : JSON.stringify(data, null, 2);
              out.innerHTML = '<p class="msg-err">' + (data.error || "请求失败") + "\\n" + detail + "</p>";
              return;
            }}
            out.innerHTML = '<pre class="msg-ok" style="white-space:pre-wrap;word-break:break-all;">' +
              JSON.stringify(data, null, 2) + "</pre>";
            return;
          }}
          if (!res.ok) {{
            const t = await res.text();
            out.innerHTML = '<p class="msg-err">服务错误 (' + res.status + ")\\n" + t.slice(0, 2000) + "</p>";
            return;
          }}
          const blob = await res.blob();
          let name = "test_cases.xlsx";
          const dis = res.headers.get("content-disposition");
          if (dis) {{
            const m = /filename\\*?=(?:UTF-8'')?([^;\\n]+)/i.exec(dis);
            if (m) try {{ name = decodeURIComponent(m[1].replace(/['"]/g, "").trim()); }} catch (e) {{}}
          }}
          const url = URL.createObjectURL(blob);
          const a = document.createElement("a");
          a.href = url;
          a.download = name;
          a.click();
          URL.revokeObjectURL(url);
          out.innerHTML = '<p class="msg-ok">已触发下载：' + name + "</p>";
        }} catch (e) {{
          out.innerHTML = '<p class="msg-err">' + (e && e.message ? e.message : String(e)) + "</p>";
        }} finally {{
          btn.disabled = false;
        }}
      }});
    }})();
  </script>
</body>
</html>"""


def excel_i18n_module_html(username: str) -> str:
    lang_opts = "".join(
        f'<option value="{code}">{LANG_LABELS.get(code, code)}</option>'
        for code in LANG_CHOICES
    )
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Excel 多语种翻译</title>
  <style>{shell_css()}</style>
</head>
<body>
  <header class="top">
    <h1>Excel 多语种翻译</h1>
    <div class="top-actions">
      <span class="who"><strong>{username}</strong></span>
      <a class="logout" href="/logout">退出</a>
    </div>
  </header>
  <main>
    <p class="content-nav"><a href="/dashboard">返回工作台</a></p>
    <h2 class="page-title">上传 Excel，选择要翻译的列，合并输出</h2>
    <div class="panel">
      <p>支持 .xlsx / .xls。先上传并点击「识别列」，再<strong>多选</strong>需要翻译的列（字母 + 表头）。列名为 <code class="inline">messages</code> 时按 JSON 内 <code class="inline">content</code> 翻译，其它列按纯文本翻译。未选中的列保持原样写入结果表。</p>
      <label for="excel-i18n-file">选择文件</label>
      <input id="excel-i18n-file" type="file" accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" />
      <p style="margin:1rem 0 0.5rem">
        <button type="button" class="primary" id="btn-inspect">识别列</button>
      </p>
      <div class="i18n-lang">
        <label for="target-lang">目标语言</label>
        <select id="target-lang">{lang_opts}</select>
      </div>
      <div id="column-picker"></div>
      <p style="margin-top:1rem">
        <button type="button" class="primary" id="btn-translate" disabled>开始翻译并下载</button>
      </p>
      <pre id="i18n-log" class="i18n-log" aria-live="polite"></pre>
      <div id="i18n-out"></div>
    </div>
  </main>
  <script>
    (function () {{
      const fileInput = document.getElementById("excel-i18n-file");
      const btnInspect = document.getElementById("btn-inspect");
      const btnTranslate = document.getElementById("btn-translate");
      const picker = document.getElementById("column-picker");
      const langSel = document.getElementById("target-lang");
      const out = document.getElementById("i18n-out");
      const logEl = document.getElementById("i18n-log");
      let lastSheets = null;

      function renderPicker(sheets) {{
        picker.innerHTML = "";
        if (!sheets || !sheets.length) {{
          picker.innerHTML = '<p class="msg-err">未读到工作表</p>';
          btnTranslate.disabled = true;
          return;
        }}
        sheets.forEach(function (sh) {{
          const block = document.createElement("div");
          block.className = "sheet-block";
          const h = document.createElement("h4");
          h.textContent = "工作表：" + sh.name;
          block.appendChild(h);
          (sh.columns || []).forEach(function (c) {{
            const lab = document.createElement("label");
            lab.className = "col-cb-row";
            const inp = document.createElement("input");
            inp.type = "checkbox";
            inp.className = "col-cb";
            inp.setAttribute("data-sheet", sh.name);
            inp.setAttribute("data-col", String(c.index));
            lab.appendChild(inp);
            lab.appendChild(document.createTextNode(" " + c.letter + " — " + c.header));
            block.appendChild(lab);
          }});
          picker.appendChild(block);
        }});
        btnTranslate.disabled = false;
      }}

      function buildColumnsPayload() {{
        const map = {{}};
        document.querySelectorAll(".col-cb:checked").forEach(function (cb) {{
          const sh = cb.getAttribute("data-sheet");
          const idx = parseInt(cb.getAttribute("data-col"), 10);
          if (!map[sh]) map[sh] = [];
          map[sh].push(idx);
        }});
        return map;
      }}

      btnInspect.addEventListener("click", async function () {{
        out.innerHTML = "";
        if (logEl) {{ logEl.textContent = ""; logEl.classList.remove("active"); }}
        const f = fileInput.files && fileInput.files[0];
        if (!f) {{
          out.innerHTML = '<p class="msg-err">请先选择 Excel 文件</p>';
          return;
        }}
        btnInspect.disabled = true;
        const fd = new FormData();
        fd.append("file", f, f.name);
        try {{
          const res = await fetch("/module/excel-i18n/inspect", {{
            method: "POST",
            body: fd,
            credentials: "same-origin",
          }});
          if (res.status === 401) {{
            out.innerHTML = '<p class="msg-err">未登录或会话已过期。</p>';
            return;
          }}
          const data = await res.json();
          if (!res.ok) {{
            out.innerHTML = '<p class="msg-err">' + (data.error || "识别失败") + "</p>";
            return;
          }}
          lastSheets = data.sheets;
          renderPicker(lastSheets);
          out.innerHTML = '<p class="msg-ok">已识别列，请勾选要翻译的列后点击「开始翻译」。</p>';
        }} catch (e) {{
          out.innerHTML = '<p class="msg-err">' + (e && e.message ? e.message : String(e)) + "</p>";
        }} finally {{
          btnInspect.disabled = false;
        }}
      }});

      btnTranslate.addEventListener("click", async function () {{
        out.innerHTML = "";
        if (logEl) {{
          logEl.textContent = "";
          logEl.classList.add("active");
        }}
        const f = fileInput.files && fileInput.files[0];
        if (!f) {{
          out.innerHTML = '<p class="msg-err">请先选择文件并识别列</p>';
          if (logEl) logEl.classList.remove("active");
          return;
        }}
        const cols = buildColumnsPayload();
        const any = Object.keys(cols).some(function (k) {{ return cols[k].length > 0; }});
        if (!any) {{
          out.innerHTML = '<p class="msg-err">请至少勾选一列</p>';
          if (logEl) logEl.classList.remove("active");
          return;
        }}
        btnTranslate.disabled = true;
        const fd = new FormData();
        fd.append("file", f, f.name);
        fd.append("language", langSel.value);
        fd.append("columns_json", JSON.stringify(cols));
        let streamFailed = false;
        let downloadToken = null;
        let doneStem = "export";
        let doneLang = langSel.value || "de";
        try {{
          const res = await fetch("/module/excel-i18n/translate", {{
            method: "POST",
            body: fd,
            credentials: "same-origin",
          }});
          if (res.status === 401) {{
            out.innerHTML = '<p class="msg-err">未登录或会话已过期。</p>';
            streamFailed = true;
            return;
          }}
          const ct = (res.headers.get("content-type") || "").split(";")[0].trim();
          if (ct === "application/json") {{
            const data = await res.json();
            out.innerHTML = '<p class="msg-err">' + (data.error || "失败") + "\\n" + (data.detail || "") + "</p>";
            streamFailed = true;
            return;
          }}
          if (!res.ok || !res.body) {{
            const t = await res.text();
            out.innerHTML = '<p class="msg-err">错误 ' + res.status + "\\n" + t.slice(0, 1500) + "</p>";
            streamFailed = true;
            return;
          }}
          const reader = res.body.getReader();
          const dec = new TextDecoder();
          let buf = "";
          while (true) {{
            const {{ done, value }} = await reader.read();
            if (done) break;
            buf += dec.decode(value, {{ stream: true }});
            for (;;) {{
              const sep = buf.indexOf("\\n\\n");
              if (sep === -1) break;
              const chunk = buf.slice(0, sep);
              buf = buf.slice(sep + 2);
              chunk.split("\\n").forEach(function (line) {{
                if (line.indexOf("data: ") !== 0) return;
                try {{
                  const data = JSON.parse(line.slice(6));
                  if (data.type === "log" && logEl) {{
                    logEl.textContent += data.line + "\\n";
                    logEl.scrollTop = logEl.scrollHeight;
                  }} else if (data.type === "error") {{
                    streamFailed = true;
                    out.innerHTML = '<p class="msg-err">' + (data.message || "翻译失败") + "</p>";
                  }} else if (data.type === "done") {{
                    downloadToken = data.token;
                    if (data.stem) doneStem = data.stem;
                    if (data.lang) doneLang = data.lang;
                  }}
                }} catch (e) {{}}
              }});
            }}
          }}
          if (!streamFailed && downloadToken) {{
            const q = "?stem=" + encodeURIComponent(doneStem) + "&lang=" + encodeURIComponent(doneLang);
            const dlUrl = "/module/excel-i18n/download/" + encodeURIComponent(downloadToken) + q;
            const name = doneStem + "_" + doneLang + "_i18n.xlsx";
            out.innerHTML =
              '<p class="msg-ok">翻译完成。请点击下方按钮下载（需在本页登录状态下点击）。</p>' +
              '<p style="margin-top:0.85rem">' +
              '<a class="logout" href="' + dlUrl + '" download="' + name.replace(/"/g, "") + '">下载 ' + name + "</a>" +
              "</p>" +
              '<p class="stub" style="margin-top:0.6rem">若未弹出保存窗口，可右键链接选择「链接另存为…」。</p>';
          }} else if (!streamFailed && !downloadToken) {{
            out.innerHTML = '<p class="msg-err">未收到完成信号，请重试。</p>';
          }}
        }} catch (e) {{
          out.innerHTML = '<p class="msg-err">' + (e && e.message ? e.message : String(e)) + "</p>";
        }} finally {{
          btnTranslate.disabled = false;
        }}
      }});
    }})();
  </script>
</body>
</html>"""


def stub_module_html(title: str, username: str) -> str:
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{title}</title>
  <style>{shell_css()}</style>
</head>
<body>
  <header class="top">
    <h1>{title}</h1>
    <div class="top-actions">
      <span class="who"><strong>{username}</strong></span>
      <a class="logout" href="/logout">退出</a>
    </div>
  </header>
  <main>
    <p class="content-nav"><a href="/dashboard">返回工作台</a></p>
    <p class="stub">该模块页面可在此接入具体功能（当前为占位）。</p>
  </main>
</body>
</html>"""


@app.get("/", response_class=HTMLResponse)
async def login_page(request: Request):
    if request.session.get(SESSION_USER):
        return RedirectResponse("/dashboard", status_code=303)
    return login_html()


@app.post("/login")
async def login(request: Request, username: str = Form(), password: str = Form()):
    if username == DEMO_USERNAME and password == DEMO_PASSWORD:
        request.session[SESSION_USER] = username
        return RedirectResponse("/dashboard", status_code=303)
    return HTMLResponse(login_html("用户名或密码错误"), status_code=401)


@app.get("/logout")
async def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/?out=1", status_code=303)


@app.get("/dashboard", response_class=HTMLResponse)
async def dashboard(request: Request):
    redir = redirect_if_not_logged_in(request)
    if redir:
        return redir
    u = request.session[SESSION_USER]
    return HTMLResponse(dashboard_html(u))


@app.get("/module/use-case-generalize", response_class=HTMLResponse)
async def module_generalize(request: Request):
    redir = redirect_if_not_logged_in(request)
    if redir:
        return redir
    return HTMLResponse(
        generalize_module_html(request.session[SESSION_USER], bool(_generalize_api_url()))
    )


@app.post("/module/use-case-generalize/process")
async def module_generalize_process(request: Request, file: UploadFile = File(...)):
    if not request.session.get(SESSION_USER):
        return JSONResponse({"error": "未登录"}, status_code=401)
    api_url = _generalize_api_url()
    if not api_url:
        return JSONResponse(
            {"error": "未配置 GENERALIZE_API_URL"},
            status_code=503,
        )
    return await _forward_multipart_file(
        file,
        api_url,
        GENERALIZE_API_FILE_FIELD,
        GENERALIZE_API_BEARER,
        unreachable_msg="无法连接泛化服务",
        upstream_fail_msg="泛化服务返回错误",
        download_suffix="_generalized.xlsx",
        default_upload_content_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        default_filename="upload.xlsx",
    )


@app.get("/module/use-case-generate", response_class=HTMLResponse)
async def module_generate(request: Request):
    redir = redirect_if_not_logged_in(request)
    if redir:
        return redir
    return HTMLResponse(
        generate_module_html(request.session[SESSION_USER], bool(_generate_api_url()))
    )


@app.post("/module/use-case-generate/process")
async def module_generate_process(request: Request, file: UploadFile = File(...)):
    if not request.session.get(SESSION_USER):
        return JSONResponse({"error": "未登录"}, status_code=401)
    api_url = _generate_api_url()
    if not api_url:
        return JSONResponse(
            {"error": "未配置 GENERATE_API_URL"},
            status_code=503,
        )
    return await _forward_multipart_file(
        file,
        api_url,
        GENERATE_API_FILE_FIELD,
        GENERATE_API_BEARER,
        unreachable_msg="无法连接用例生成服务",
        upstream_fail_msg="用例生成服务返回错误",
        download_suffix="_test_cases.xlsx",
        default_upload_content_type=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        default_filename="prd.docx",
    )


@app.get("/module/excel-i18n", response_class=HTMLResponse)
async def module_excel_i18n(request: Request):
    redir = redirect_if_not_logged_in(request)
    if redir:
        return redir
    return HTMLResponse(excel_i18n_module_html(request.session[SESSION_USER]))


@app.post("/module/excel-i18n/inspect")
async def excel_i18n_inspect(request: Request, file: UploadFile = File(...)):
    if not request.session.get(SESSION_USER):
        return JSONResponse({"error": "未登录"}, status_code=401)
    suf = Path(file.filename or "").suffix.lower()
    if suf not in (".xlsx", ".xls"):
        return JSONResponse({"error": "请上传 .xlsx 或 .xls"}, status_code=400)
    raw = await file.read()
    if not raw:
        return JSONResponse({"error": "空文件"}, status_code=400)
    fd: Optional[int] = None
    tmp_path: Optional[str] = None
    try:
        fd, tmp_path = tempfile.mkstemp(suffix=suf)
        os.write(fd, raw)
        os.close(fd)
        fd = None
        sheets = inspect_workbook_column_meta(Path(tmp_path))
        return JSONResponse({"sheets": sheets})
    except Exception as exc:
        return JSONResponse({"error": "读取表头失败", "detail": str(exc)}, status_code=400)
    finally:
        if fd is not None:
            try:
                os.close(fd)
            except OSError:
                pass
        if tmp_path:
            Path(tmp_path).unlink(missing_ok=True)


def _excel_i18n_sse_chunk(payload: dict) -> bytes:
    return f"data: {json.dumps(payload, ensure_ascii=False)}\n\n".encode("utf-8")


def _parse_columns_by_sheet(raw: str) -> dict[str, set[int]]:
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError("columns_json 应为 JSON 对象：工作表名 -> 列索引数组")
    out: dict[str, set[int]] = {}
    for k, v in data.items():
        sk = str(k)
        if not isinstance(v, list):
            raise ValueError(f"工作表 {sk!r} 的值应为数组")
        try:
            out[sk] = {int(x) for x in v}
        except (TypeError, ValueError) as e:
            raise ValueError(f"工作表 {sk!r} 的列须为整数") from e
    return out


@app.post("/module/excel-i18n/translate")
async def excel_i18n_translate(
    request: Request,
    file: UploadFile = File(...),
    language: str = Form(...),
    columns_json: str = Form(...),
):
    if not request.session.get(SESSION_USER):
        return JSONResponse({"error": "未登录"}, status_code=401)
    lang = language.strip().lower()
    if lang not in LANG_CHOICES:
        return JSONResponse(
            {"error": "无效目标语言", "choices": list(LANG_CHOICES)},
            status_code=400,
        )
    try:
        columns_by_sheet = _parse_columns_by_sheet(columns_json)
    except json.JSONDecodeError as exc:
        return JSONResponse(
            {"error": "columns_json 不是合法 JSON", "detail": str(exc)},
            status_code=400,
        )
    except ValueError as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)
    if not columns_by_sheet or not any(columns_by_sheet.values()):
        return JSONResponse({"error": "请至少选择一列"}, status_code=400)

    suf = Path(file.filename or "").suffix.lower()
    if suf not in (".xlsx", ".xls"):
        return JSONResponse({"error": "请上传 .xlsx 或 .xls"}, status_code=400)
    raw = await file.read()
    if not raw:
        return JSONResponse({"error": "空文件"}, status_code=400)

    stem = Path(file.filename or "export").stem
    safe_stem = stem.replace('"', "_").replace("/", "_").replace("\\", "_")[:80]

    fd_in: Optional[int] = None
    fd_out: Optional[int] = None
    tmp_in: Optional[str] = None
    tmp_out: Optional[str] = None
    try:
        fd_in, tmp_in = tempfile.mkstemp(suffix=suf)
        os.write(fd_in, raw)
        os.close(fd_in)
        fd_in = None
        fd_out, tmp_out = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd_out)
        fd_out = None
    except OSError as exc:
        return JSONResponse({"error": "临时文件创建失败", "detail": str(exc)}, status_code=500)

    tmp_in_path = tmp_in
    tmp_out_path = tmp_out
    columns_ref = columns_by_sheet
    lang_ref = lang
    stem_ref = safe_stem

    async def event_stream():
        loop = asyncio.get_running_loop()
        q: asyncio.Queue = asyncio.Queue()

        def worker() -> None:
            try:

                def log_cb(msg: str) -> None:
                    loop.call_soon_threadsafe(q.put_nowait, ("log", msg))

                process_file_with_column_selection(
                    Path(tmp_in_path),
                    lang_ref,
                    Path(tmp_out_path),
                    columns_ref,
                    verbose=False,
                    params=TranslateParams(log_sink=log_cb),
                )
                tok = secrets.token_urlsafe(24)
                with _EXCEL_DL_LOCK:
                    _EXCEL_DL_FILES[tok] = Path(tmp_out_path)
                loop.call_soon_threadsafe(q.put_nowait, ("token", tok))
            except Exception as e:
                loop.call_soon_threadsafe(q.put_nowait, ("err", str(e)))
                try:
                    Path(tmp_out_path).unlink(missing_ok=True)
                except OSError:
                    pass
            finally:
                try:
                    Path(tmp_in_path).unlink(missing_ok=True)
                except OSError:
                    pass

        threading.Thread(target=worker, daemon=True).start()
        while True:
            kind, *rest = await q.get()
            if kind == "log":
                yield _excel_i18n_sse_chunk({"type": "log", "line": rest[0]})
            elif kind == "err":
                yield _excel_i18n_sse_chunk({"type": "error", "message": rest[0]})
                return
            elif kind == "token":
                yield _excel_i18n_sse_chunk(
                    {
                        "type": "done",
                        "token": rest[0],
                        "lang": lang_ref,
                        "stem": stem_ref,
                    }
                )
                return

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


@app.get("/module/excel-i18n/download/{token}")
async def excel_i18n_download(
    request: Request,
    token: str,
    stem: str = "export",
    lang: str = "de",
):
    if not request.session.get(SESSION_USER):
        return JSONResponse({"error": "未登录"}, status_code=401)
    if not token or len(token) > 128:
        return JSONResponse({"error": "无效令牌"}, status_code=400)
    lang_l = lang.strip().lower()
    if lang_l not in LANG_CHOICES:
        lang_l = "de"
    safe_stem = stem.replace('"', "_").replace("/", "_").replace("\\", "_")[:80]
    out_name = f"{safe_stem}_{lang_l}_i18n.xlsx"
    with _EXCEL_DL_LOCK:
        path = _EXCEL_DL_FILES.pop(token, None)
    if path is None or not path.exists():
        return JSONResponse({"error": "文件不存在或已下载"}, status_code=404)
    try:
        out_bytes = path.read_bytes()
    finally:
        path.unlink(missing_ok=True)
    return Response(
        content=out_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
    )


@app.get("/items/{item_id}")
async def read_item(item_id: str):
    return {"item_id": item_id}

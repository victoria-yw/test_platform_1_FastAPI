from fastapi import APIRouter, File, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse

from app.config import (
    GENERATE_API_BEARER,
    GENERATE_API_FILE_FIELD,
    SESSION_USER,
    generate_api_url,
)
from app.ui_common import redirect_if_not_logged_in, shell_css
from app.upstream_multipart import forward_multipart_file

router = APIRouter(tags=["module-generate"])


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


@router.get("/module/use-case-generate", response_class=HTMLResponse)
async def module_generate_page(request: Request):
    redir = redirect_if_not_logged_in(request)
    if redir:
        return redir
    return HTMLResponse(
        generate_module_html(request.session[SESSION_USER], bool(generate_api_url()))
    )


@router.post("/module/use-case-generate/process")
async def module_generate_process(request: Request, file: UploadFile = File(...)):
    if not request.session.get(SESSION_USER):
        return JSONResponse({"error": "未登录"}, status_code=401)
    api_url = generate_api_url()
    if not api_url:
        return JSONResponse(
            {"error": "未配置 GENERATE_API_URL"},
            status_code=503,
        )
    return await forward_multipart_file(
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

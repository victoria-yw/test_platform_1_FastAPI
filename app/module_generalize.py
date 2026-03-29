from fastapi import APIRouter, File, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse

from app.config import (
    GENERALIZE_API_BEARER,
    GENERALIZE_API_FILE_FIELD,
    SESSION_USER,
    generalize_api_url,
)
from app.ui_common import redirect_if_not_logged_in, shell_css
from app.upstream_multipart import forward_multipart_file

router = APIRouter(tags=["module-generalize"])


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


@router.get("/module/use-case-generalize", response_class=HTMLResponse)
async def module_generalize_page(request: Request):
    redir = redirect_if_not_logged_in(request)
    if redir:
        return redir
    return HTMLResponse(
        generalize_module_html(request.session[SESSION_USER], bool(generalize_api_url()))
    )


@router.post("/module/use-case-generalize/process")
async def module_generalize_process(request: Request, file: UploadFile = File(...)):
    if not request.session.get(SESSION_USER):
        return JSONResponse({"error": "未登录"}, status_code=401)
    api_url = generalize_api_url()
    if not api_url:
        return JSONResponse(
            {"error": "未配置 GENERALIZE_API_URL"},
            status_code=503,
        )
    return await forward_multipart_file(
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

import asyncio
import json
import os
import secrets
import tempfile
import threading
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, File, Form, Request, UploadFile
from fastapi.responses import (
    HTMLResponse,
    JSONResponse,
    Response,
    StreamingResponse,
)
from translate_module import (
    LANG_CHOICES,
    LANG_LABELS,
    TranslateParams,
    inspect_workbook_column_meta,
    process_file_with_column_selection,
)

from app.config import SESSION_USER
from app.ui_common import redirect_if_not_logged_in, shell_css

router = APIRouter(tags=["module-excel-i18n"])

_EXCEL_DL_LOCK = threading.Lock()
_EXCEL_DL_FILES: dict[str, Path] = {}


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


@router.get("/module/excel-i18n", response_class=HTMLResponse)
async def module_excel_i18n_page(request: Request):
    redir = redirect_if_not_logged_in(request)
    if redir:
        return redir
    return HTMLResponse(excel_i18n_module_html(request.session[SESSION_USER]))


@router.post("/module/excel-i18n/inspect")
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


@router.post("/module/excel-i18n/translate")
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


@router.get("/module/excel-i18n/download/{token}")
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

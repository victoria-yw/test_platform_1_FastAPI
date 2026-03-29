from typing import Optional

from fastapi import Request
from fastapi.responses import RedirectResponse

from app.config import SESSION_USER


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

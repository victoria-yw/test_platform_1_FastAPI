from pathlib import Path
from typing import Union

import httpx
from fastapi import UploadFile
from fastapi.responses import JSONResponse, Response


async def forward_multipart_file(
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

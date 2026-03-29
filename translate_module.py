#!/usr/bin/env python3
"""
Excel 消息翻译工具（5000 字截断版）：多工作表、input / messages 列，目标语 de / fr / nl / en。

与 translate.py 行为一致，区别：任一段待译文本若超过 5000 个字符，只取前 5000 字送翻，
不拆成多段拼接。列名与 JSON 键名、结构不变。
"""

from __future__ import annotations

import argparse
import json
import random
import sys
import time
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Optional

warnings.filterwarnings(
    "ignore",
    message=r"urllib3 v2 only supports OpenSSL",
)

import pandas as pd
from deep_translator import GoogleTranslator

LANG_CHOICES = ("de", "fr", "nl", "en")

LANG_LABELS = {
    "de": "德语 (de)",
    "fr": "法语 (fr)",
    "nl": "荷兰语 (nl)",
    "en": "英语 (en)",
}
INPUT_SUFFIXES = (".xlsx", ".xls")
COL_INPUT = "input"
COL_MESSAGES = "messages"
JSON_INDENT = 4

MAX_TRANSLATE_CHARS = 5000
DEFAULT_MAX_RETRIES = 8
DEFAULT_RETRY_BASE_DELAY = 1.5


@dataclass(frozen=True)
class TranslateParams:
    max_retries: int = DEFAULT_MAX_RETRIES
    retry_base_delay: float = DEFAULT_RETRY_BASE_DELAY
    # 若设置，每条进度文案会回调（例如 Web 端回显）；可与 verbose 同时使用
    log_sink: Optional[Callable[[str], None]] = None


def _notify(params: TranslateParams, msg: str, *, verbose: bool) -> None:
    if params.log_sink is not None:
        params.log_sink(msg)
    if verbose:
        print(msg, flush=True)


def _is_empty_cell(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    return False


def _preview_text(s: str, max_len: int = 72) -> str:
    t = s.replace("\n", " ").strip()
    if len(t) > max_len:
        return t[: max_len - 1] + "…"
    return t


def _translate_once(
    translator: GoogleTranslator,
    text: str,
    params: TranslateParams,
    *,
    verbose: bool = False,
) -> str:
    last_exc: Exception | None = None
    for attempt in range(params.max_retries):
        try:
            return translator.translate(text)
        except Exception as e:
            last_exc = e
            if attempt == params.max_retries - 1:
                break
            delay = params.retry_base_delay * (2**attempt) + random.uniform(0.0, 0.5)
            _notify(
                params,
                f"      请求失败（{type(e).__name__}），{delay:.1f}s 后重试 "
                f"{attempt + 2}/{params.max_retries}…",
                verbose=verbose,
            )
            time.sleep(delay)
    assert last_exc is not None
    raise last_exc


def _truncate_for_translate(
    s: str,
    params: TranslateParams,
    *,
    verbose: bool = False,
    label: str = "",
) -> str:
    """超过 MAX_TRANSLATE_CHARS 时只保留前缀。"""
    s = str(s).strip()
    if not s:
        return s
    if len(s) <= MAX_TRANSLATE_CHARS:
        return s
    _notify(
        params,
        f"      原文 {len(s)} 字，超过 {MAX_TRANSLATE_CHARS}，仅翻译前 {MAX_TRANSLATE_CHARS} 字"
        f"{(' — ' + label) if label else ''}",
        verbose=verbose,
    )
    return s[:MAX_TRANSLATE_CHARS]


def _translate_capped(
    s: str,
    translator: GoogleTranslator,
    params: TranslateParams,
    *,
    verbose: bool = False,
    cap_label: str = "",
) -> str:
    s = str(s).strip()
    if not s:
        return s
    to_send = _truncate_for_translate(s, params, verbose=verbose, label=cap_label)
    return _translate_once(translator, to_send, params, verbose=verbose)


def _translate_plain(
    text,
    translator: GoogleTranslator,
    params: TranslateParams,
    *,
    verbose: bool = False,
):
    if _is_empty_cell(text):
        return text
    s = str(text).strip()
    if not s:
        return text
    return _translate_capped(s, translator, params, verbose=verbose, cap_label="input")


def _parse_messages_json(cell):
    if _is_empty_cell(cell):
        return None
    s = str(cell).strip()
    if not s:
        return None
    data = json.loads(s)
    if not isinstance(data, list):
        raise ValueError("JSON must be a list")
    return data


def _dump_messages_json(messages: list) -> str:
    return json.dumps(messages, ensure_ascii=False, indent=JSON_INDENT)


def _translate_messages_cell(
    cell,
    translator: GoogleTranslator,
    *,
    sheet_name: str = "",
    row_num: int = 0,
    row_total: int = 0,
    verbose: bool = True,
    params: TranslateParams | None = None,
):
    if params is None:
        params = TranslateParams()
    if _is_empty_cell(cell):
        return cell
    try:
        messages = _parse_messages_json(cell)
        if messages is None:
            return cell
        nmsg = len(messages)
        new_list = []
        for mi, m in enumerate(messages):
            if not isinstance(m, dict):
                new_list.append(m)
                continue
            d = {}
            role_guess = m.get("role", "?")
            for k, v in m.items():
                if k == "content":
                    if _is_empty_cell(v):
                        d[k] = v
                    else:
                        c = str(v).strip()
                        if c:
                            _notify(
                                params,
                                f"    [{sheet_name}] 行 {row_num}/{row_total} · "
                                f"messages 第 {mi + 1}/{nmsg} 条 ({role_guess}) → {_preview_text(c)}",
                                verbose=verbose,
                            )
                            d[k] = _translate_capped(
                                c,
                                translator,
                                params,
                                verbose=verbose,
                                cap_label=f"messages 第{mi + 1}条",
                            )
                        else:
                            d[k] = v
                else:
                    d[k] = v
            new_list.append(d)
        return _dump_messages_json(new_list)
    except (json.JSONDecodeError, ValueError) as e:
        _notify(
            params,
            f"警告: 跳过无效的 messages 单元格: {e}",
            verbose=verbose,
        )
        if not params.log_sink:
            print(f"警告: 跳过无效的 messages 单元格: {e}", file=sys.stderr)
        return cell


def _excel_engine_for_path(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".xls":
        return "xlrd"
    return "openpyxl"


def _read_excel_all(path: Path) -> dict[str, pd.DataFrame]:
    return pd.read_excel(path, sheet_name=None, engine=_excel_engine_for_path(path))


def _read_excel_headers_only(path: Path) -> dict[str, pd.DataFrame]:
    """仅表头，用于快速列出列字母与表头名。"""
    return pd.read_excel(
        path, sheet_name=None, nrows=0, engine=_excel_engine_for_path(path)
    )


def column_index_to_letter(index: int) -> str:
    """0 -> A, 25 -> Z, 26 -> AA（Excel 列字母）。"""
    if index < 0:
        raise ValueError("column index must be non-negative")
    n = index + 1
    letters: list[str] = []
    while n:
        n, r = divmod(n - 1, 26)
        letters.append(chr(65 + r))
    return "".join(reversed(letters))


def inspect_workbook_column_meta(path: Path) -> list[dict[str, Any]]:
    """供 Web 上传预览：各工作表的列 index / 字母 / 表头字符串。"""
    sheets = _read_excel_headers_only(path)
    out: list[dict[str, Any]] = []
    for sheet_name, df in sheets.items():
        cols = []
        for i, c in enumerate(df.columns):
            cols.append(
                {
                    "index": i,
                    "letter": column_index_to_letter(i),
                    "header": str(c),
                }
            )
        out.append({"name": str(sheet_name), "columns": cols})
    return out


def _first_col_index(df: pd.DataFrame, logical_name: str) -> int:
    for j, c in enumerate(df.columns):
        if str(c) == logical_name:
            return j
    raise KeyError(logical_name)


def _translate_dataframe(
    df: pd.DataFrame,
    translator: GoogleTranslator,
    sheet_name: str,
    verbose: bool,
    params: TranslateParams,
) -> pd.DataFrame:
    out = df.copy()
    cols = set(out.columns.astype(str))
    has_input = COL_INPUT in cols
    has_msgs = COL_MESSAGES in cols
    if not has_input and not has_msgs:
        return out

    ic_in = _first_col_index(out, COL_INPUT) if has_input else None
    ic_msg = _first_col_index(out, COL_MESSAGES) if has_msgs else None
    n_rows = len(out)

    for i in range(n_rows):
        if has_input:
            val = out.iat[i, ic_in]
            if not _is_empty_cell(val):
                s = str(val).strip()
                if s:
                    _notify(
                        params,
                        f"  [{sheet_name}] 行 {i + 1}/{n_rows} · input → {_preview_text(s)}",
                        verbose=verbose,
                    )
            out.iat[i, ic_in] = _translate_plain(
                val,
                translator,
                params,
                verbose=verbose,
            )

        if has_msgs:
            val = out.iat[i, ic_msg]
            out.iat[i, ic_msg] = _translate_messages_cell(
                val,
                translator,
                sheet_name=sheet_name,
                row_num=i + 1,
                row_total=n_rows,
                verbose=verbose,
                params=params,
            )

    return out


def _translate_dataframe_selected_columns(
    df: pd.DataFrame,
    col_indices: set[int],
    translator: GoogleTranslator,
    sheet_name: str,
    verbose: bool,
    params: TranslateParams,
) -> pd.DataFrame:
    """只翻译指定列索引；列名为 messages 时仍走 JSON 内 content 翻译逻辑，其余列按纯文本翻译。"""
    out = df.copy()
    n_rows = len(out)
    for j in sorted(col_indices):
        if j < 0 or j >= len(out.columns):
            continue
        col_name = str(out.columns[j])
        for i in range(n_rows):
            val = out.iat[i, j]
            if col_name == COL_MESSAGES:
                out.iat[i, j] = _translate_messages_cell(
                    val,
                    translator,
                    sheet_name=sheet_name,
                    row_num=i + 1,
                    row_total=n_rows,
                    verbose=verbose,
                    params=params,
                )
            else:
                if not _is_empty_cell(val):
                    s = str(val).strip()
                    if s:
                        _notify(
                            params,
                            f"  [{sheet_name}] 行 {i + 1}/{n_rows} · "
                            f"列 {column_index_to_letter(j)} ({col_name}) → {_preview_text(s)}",
                            verbose=verbose,
                        )
                out.iat[i, j] = _translate_plain(
                    val,
                    translator,
                    params,
                    verbose=verbose,
                )
    return out


def translate_workbook_with_selected_columns(
    path: Path,
    target_lang: str,
    columns_by_sheet: dict[str, set[int]],
    *,
    verbose: bool = True,
    params: TranslateParams | None = None,
) -> dict[str, pd.DataFrame]:
    """
    按工作表、按列多选翻译：未选中的列保持原样，与译后列合并为完整表输出。
    columns_by_sheet 的 key 为工作表名（字符串），value 为要翻译的 0-based 列索引集合。
    """
    if params is None:
        params = TranslateParams()
    _notify(params, f"读取: {path.name}", verbose=verbose)
    desc = {k: sorted(v) for k, v in sorted(columns_by_sheet.items())}
    _notify(params, f"待译列（按工作表）: {desc}", verbose=verbose)
    translator = GoogleTranslator(source="auto", target=target_lang)
    sheets = _read_excel_all(path)
    result: dict[str, pd.DataFrame] = {}
    total = len(sheets)
    for si, (name, df) in enumerate(sheets.items(), 1):
        name_str = str(name)
        sel = columns_by_sheet.get(name_str, set())
        if not sel:
            _notify(
                params,
                f"工作表 [{si}/{total}] {name_str!r} — 未选择列，原样保留",
                verbose=verbose,
            )
            result[name] = df.copy()
            continue
        valid = {j for j in sel if 0 <= j < len(df.columns)}
        if not valid:
            _notify(
                params,
                f"工作表 [{si}/{total}] {name_str!r} — 所选列无效，原样保留",
                verbose=verbose,
            )
            result[name] = df.copy()
            continue
        _notify(
            params,
            f"工作表 [{si}/{total}] {name_str!r} — {len(df)} 行，翻译列 {sorted(valid)}",
            verbose=verbose,
        )
        result[name] = _translate_dataframe_selected_columns(
            df,
            valid,
            translator,
            name_str,
            verbose,
            params,
        )
    return result


def process_file_with_column_selection(
    inp: Path,
    lang: str,
    out_path: Path,
    columns_by_sheet: dict[str, set[int]],
    *,
    verbose: bool = False,
    params: TranslateParams | None = None,
) -> None:
    """与 process_file 类似，但只翻译 columns_by_sheet 指定的列，其余列原样写入同一工作簿。"""
    if params is None:
        params = TranslateParams()
    _notify(
        params,
        "开始翻译（调用在线翻译服务，可能需要较长时间）…",
        verbose=verbose,
    )
    sheets = translate_workbook_with_selected_columns(
        inp,
        lang,
        columns_by_sheet,
        verbose=verbose,
        params=params,
    )
    _notify(params, "正在写入 Excel…", verbose=verbose)
    _write_xlsx(sheets, out_path)
    _notify(params, "写入完成。请在页面上点击「下载」保存文件。", verbose=verbose)
    if verbose:
        print(f"输出路径: {out_path}", flush=True)


def _should_process_sheet(df: pd.DataFrame) -> bool:
    cols = set(df.columns.astype(str))
    return COL_INPUT in cols or COL_MESSAGES in cols


def translate_workbook(
    path: Path,
    target_lang: str,
    *,
    verbose: bool = True,
    params: TranslateParams | None = None,
) -> dict[str, pd.DataFrame]:
    if params is None:
        params = TranslateParams()
    _notify(params, f"读取: {path.name}", verbose=verbose)
    translator = GoogleTranslator(source="auto", target=target_lang)
    sheets = _read_excel_all(path)
    result: dict[str, pd.DataFrame] = {}
    total = len(sheets)
    for si, (name, df) in enumerate(sheets.items(), 1):
        if _should_process_sheet(df):
            _notify(
                params,
                f"工作表 [{si}/{total}] {name!r} — {len(df)} 行",
                verbose=verbose,
            )
            result[name] = _translate_dataframe(
                df,
                translator,
                str(name),
                verbose,
                params,
            )
        else:
            _notify(
                params,
                f"工作表 [{si}/{total}] {name!r} — 无 input/messages，跳过",
                verbose=verbose,
            )
            result[name] = df.copy()
    return result


def _excel_safe_sheet_name(name: str, used: set[str]) -> str:
    base = (str(name)[:31] if name else "Sheet") or "Sheet"
    candidate = base
    i = 1
    while candidate in used:
        suffix = f"_{i}"
        candidate = (base[: 31 - len(suffix)] + suffix) if len(base) + len(suffix) > 31 else base + suffix
        i += 1
    used.add(candidate)
    return candidate


def _write_xlsx(sheets: dict[str, pd.DataFrame], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    used_names: set[str] = set()
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe = _excel_safe_sheet_name(sheet_name, used_names)
            df.to_excel(writer, sheet_name=safe, index=False)


def _list_excel_files(root: Path) -> list[Path]:
    files: list[Path] = []
    for suf in INPUT_SUFFIXES:
        files.extend(sorted(root.glob(f"*{suf}")))
    return files


def _default_stem_suffix(lang: str) -> str:
    """与 translate.py 区分，避免覆盖默认输出。"""
    return f"{lang}_cap5000"


def _resolve_single_file_output(inp: Path, lang: str, output: Path | None) -> Path:
    default_name = f"{inp.stem}_{_default_stem_suffix(lang)}.xlsx"
    if output is None:
        return inp.with_name(default_name)
    out = output.expanduser()
    if out.exists():
        if out.is_dir():
            return (out / default_name).resolve()
        return out.resolve()
    if out.suffix.lower() in INPUT_SUFFIXES:
        return out.resolve()
    out.mkdir(parents=True, exist_ok=True)
    return (out / default_name).resolve()


def process_file(
    inp: Path,
    lang: str,
    out_path: Path,
    *,
    verbose: bool = True,
    params: TranslateParams | None = None,
) -> None:
    if params is None:
        params = TranslateParams()
    if verbose:
        print(f"目标语言: {lang}  输出: {out_path}", flush=True)
        print(
            f"长度策略: 超过 {MAX_TRANSLATE_CHARS} 字仅翻译前 {MAX_TRANSLATE_CHARS} 字（其余丢弃，不送翻）",
            flush=True,
        )
        print(
            f"请求重试: 最多 {params.max_retries} 次，退避基数 {params.retry_base_delay}s",
            flush=True,
        )
    sheets = translate_workbook(inp, lang, verbose=verbose, params=params)
    if verbose:
        print("正在写入 Excel…", flush=True)
    _write_xlsx(sheets, out_path)
    print(f"已完成: {out_path}", flush=True)


def main() -> int:
    p = argparse.ArgumentParser(
        description=(
            "翻译 Excel（5000 字截断版）：input / messages 列译成 de/fr/nl；"
            f"单段文本超过 {MAX_TRANSLATE_CHARS} 字则只译前 {MAX_TRANSLATE_CHARS} 字。"
        )
    )
    p.add_argument("input_path", type=Path, help="输入 .xlsx / .xls 文件或所在文件夹")
    p.add_argument(
        "--language",
        required=True,
        choices=LANG_CHOICES,
        help="目标语言: de=德语, fr=法语, nl=荷兰语, en=英语",
    )
    p.add_argument(
        "--output",
        type=Path,
        default=None,
        help="输出文件路径或文件夹（可选）",
    )
    p.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        help="不打印翻译进度（仅保留完成行）",
    )
    p.add_argument(
        "--retries",
        type=int,
        default=DEFAULT_MAX_RETRIES,
        metavar="N",
        help="单次翻译请求失败时的最大重试次数（默认 %(default)s）",
    )
    p.add_argument(
        "--retry-delay",
        type=float,
        default=DEFAULT_RETRY_BASE_DELAY,
        metavar="SEC",
        help="重试退避时间基数（秒，默认 %(default)s）",
    )
    args = p.parse_args()
    verbose = not args.quiet
    params = TranslateParams(
        max_retries=max(1, args.retries),
        retry_base_delay=max(0.1, args.retry_delay),
    )
    inp = args.input_path.resolve()
    lang = args.language
    suf = _default_stem_suffix(lang)

    if not inp.exists():
        print(f"路径不存在: {inp}", file=sys.stderr)
        return 1

    if inp.is_file():
        if inp.suffix.lower() not in INPUT_SUFFIXES:
            print(f"不支持的文件类型: {inp.suffix}", file=sys.stderr)
            return 1
        out_path = _resolve_single_file_output(inp, lang, args.output)
        try:
            process_file(inp, lang, out_path, verbose=verbose, params=params)
        except Exception as e:
            print(f"处理失败: {e}", file=sys.stderr)
            return 1
        return 0

    if inp.is_dir():
        files = _list_excel_files(inp)
        if not files:
            print(f"文件夹中未找到 .xlsx / .xls: {inp}", file=sys.stderr)
            return 1
        out_root = args.output
        if out_root is None:
            out_root = inp.parent / f"{inp.name}_translated_{suf}"
        else:
            out_root = out_root.expanduser()
            if out_root.exists() and out_root.is_file():
                print(f"--output 应为目录: {out_root}", file=sys.stderr)
                return 1
        out_root = out_root.resolve()
        out_root.mkdir(parents=True, exist_ok=True)
        for f in files:
            out_path = out_root / f"{f.stem}_{suf}.xlsx"
            try:
                if verbose:
                    print(f"\n==== 文件: {f.name} ====", flush=True)
                process_file(f, lang, out_path, verbose=verbose, params=params)
            except Exception as e:
                print(f"跳过 {f}: {e}", file=sys.stderr)
        return 0

    print(f"无效路径: {inp}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())

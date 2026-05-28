"""Docs tools — paragraph-anchored editing on the user's open Google Doc.

Mirrors the Word MCP's model: read paragraphs to get 1-based indices, then edit
by index. Edits via the Docs v1 API show up live in the open browser tab through
Google's realtime sync.
"""

from typing import Any

from google_mcp import auth

INSTRUCTIONS = """\
You are working live on the user's open Google Doc, via the Docs API. Edits
show up in their open browser tab through Google's realtime sync.

First, point the server at the right doc: call docs_set_active(url) with the URL
from the browser address bar. After that all other tools default to that doc.
Confirm with docs_status.

Editing is paragraph-anchored: read structure with docs_get_paragraphs to get
the 1-based index of the paragraph you want, then edit by index —
docs_insert_paragraph(after=N or before=N, text=..., style=...) for mid-document
inserts, docs_replace_paragraph(N, text), docs_delete_paragraph(N).
docs_append_paragraph adds at the end. docs_get_outline returns only headings.

For small intra-paragraph edits prefer docs_find_replace(paragraph=N, find, replace)
over a full paragraph rewrite — it only touches the bytes that change.

Tables: docs_insert_table(rows, cols, after=N or before=N, data=[[...]]) inserts
and optionally fills in one call; docs_fill_table(table_index, data) overwrites
cells of an existing table; docs_get_tables() returns each table's contents.

Inline formatting: docs_format_range(paragraph=N, start, end, bold/italic/font/
size/color/link/...) sets character-level style on a range within a paragraph.
color is [r, g, b] 0-255; size is points.

Paragraph styles: NORMAL_TEXT, TITLE, SUBTITLE, HEADING_1..HEADING_6.
"""

_STYLES = {
    "NORMAL_TEXT",
    "TITLE",
    "SUBTITLE",
    "HEADING_1",
    "HEADING_2",
    "HEADING_3",
    "HEADING_4",
    "HEADING_5",
    "HEADING_6",
}


def _list_paragraphs(doc: dict) -> list[dict]:
    """Body paragraphs as {index, style, text, startIndex, endIndex}, 1-based index."""
    out: list[dict] = []
    for el in doc.get("body", {}).get("content", []):
        if "paragraph" not in el:
            continue
        text = "".join(
            run.get("textRun", {}).get("content", "")
            for run in el["paragraph"].get("elements", [])
        )
        out.append(
            {
                "index": len(out) + 1,
                "style": el["paragraph"]
                .get("paragraphStyle", {})
                .get("namedStyleType", "NORMAL_TEXT"),
                "text": text.rstrip("\n"),
                "startIndex": el.get("startIndex", 1),
                "endIndex": el["endIndex"],
            }
        )
    return out


def _paragraph_at(svc, doc_id: str, index: int) -> dict:
    paragraphs = _list_paragraphs(svc.documents().get(documentId=doc_id).execute())
    if not 1 <= index <= len(paragraphs):
        raise ValueError(f"paragraph index {index} out of range (1..{len(paragraphs)})")
    return paragraphs[index - 1]


def _check_style(style: str | None) -> None:
    if style is not None and style not in _STYLES:
        raise ValueError(f"unknown style {style!r}; choices: {sorted(_STYLES)}")


def _text_style_payload(
    *,
    bold=None,
    italic=None,
    underline=None,
    strikethrough=None,
    font=None,
    size=None,
    color=None,
    link=None,
) -> tuple[dict, str]:
    """Build (textStyle, fields-mask) for an updateTextStyle request from the
    attributes actually passed in. Returns ({}, "") if nothing was specified."""
    style: dict = {}
    fields: list[str] = []
    if bold is not None:
        style["bold"] = bold
        fields.append("bold")
    if italic is not None:
        style["italic"] = italic
        fields.append("italic")
    if underline is not None:
        style["underline"] = underline
        fields.append("underline")
    if strikethrough is not None:
        style["strikethrough"] = strikethrough
        fields.append("strikethrough")
    if font is not None:
        style["weightedFontFamily"] = {"fontFamily": font}
        fields.append("weightedFontFamily")
    if size is not None:
        style["fontSize"] = {"magnitude": float(size), "unit": "PT"}
        fields.append("fontSize")
    if color is not None:
        if len(color) != 3:
            raise ValueError("color must be [r, g, b] with 0-255 values")
        r, g, b = color
        style["foregroundColor"] = {
            "color": {"rgbColor": {"red": r / 255, "green": g / 255, "blue": b / 255}}
        }
        fields.append("foregroundColor")
    if link is not None:
        style["link"] = {"url": link}
        fields.append("link")
    return style, ",".join(fields)


def _walk_tables(doc: dict):
    """Iterate (1-based table index, table element) over body tables."""
    ti = 0
    for el in doc.get("body", {}).get("content", []):
        if "table" in el:
            ti += 1
            yield ti, el


def _read_cell_text(cell: dict) -> str:
    text = ""
    for sub in cell.get("content", []):
        if "paragraph" in sub:
            text += "".join(
                run.get("textRun", {}).get("content", "")
                for run in sub["paragraph"].get("elements", [])
            )
    return text.rstrip("\n")


def _table_contents(table_el: dict) -> list[list[str]]:
    return [
        [_read_cell_text(c) for c in row["tableCells"]]
        for row in table_el["table"]["tableRows"]
    ]


def _fill_cells(svc, sid: str, table_el: dict, data: list[list]) -> None:
    """Overwrite table cells with `data[r][c]` values. Indices are taken from
    the fresh `table_el`; requests are issued in descending order of location so
    higher-offset edits don't shift lower-offset ones within the batch."""
    requests: list[dict] = []
    for r_i, row in enumerate(table_el["table"]["tableRows"]):
        if r_i >= len(data):
            break
        for c_i, cell in enumerate(row["tableCells"]):
            if c_i >= len(data[r_i]):
                continue
            text = str(data[r_i][c_i])
            if not text:
                continue
            first_para = cell["content"][0]
            start = first_para["startIndex"]
            end = first_para["endIndex"]
            if end - start > 1:
                requests.append(
                    {"deleteContentRange": {"range": {"startIndex": start, "endIndex": end - 1}}}
                )
            requests.append({"insertText": {"location": {"index": start}, "text": text}})

    def key(req):
        if "deleteContentRange" in req:
            return (-req["deleteContentRange"]["range"]["startIndex"], 0)
        return (-req["insertText"]["location"]["index"], 1)

    requests.sort(key=key)
    if requests:
        svc.documents().batchUpdate(documentId=sid, body={"requests": requests}).execute()


def register(mcp):
    @mcp.tool
    def docs_set_active(url_or_id: str) -> str:
        """Point the server at a Google Doc by URL (or ID). Persists across sessions."""
        return f"active document set: {auth.set_active('docs', url_or_id)}"

    @mcp.tool
    def docs_status() -> dict:
        """The active doc's title, ID, URL, and paragraph count."""
        sid = auth.require_active("docs")
        doc = auth.docs_service().documents().get(documentId=sid).execute()
        return {
            "id": sid,
            "title": doc.get("title"),
            "url": f"https://docs.google.com/document/d/{sid}/edit",
            "paragraphs": len(_list_paragraphs(doc)),
        }

    @mcp.tool
    def docs_get_paragraphs() -> list:
        """All paragraphs as {index, style, text}. Indices are 1-based."""
        sid = auth.require_active("docs")
        doc = auth.docs_service().documents().get(documentId=sid).execute()
        return [
            {"index": p["index"], "style": p["style"], "text": p["text"]}
            for p in _list_paragraphs(doc)
        ]

    @mcp.tool
    def docs_get_outline() -> list:
        """Headings only as {index, level, text}. TITLE is level 0, HEADING_N is level N."""
        sid = auth.require_active("docs")
        doc = auth.docs_service().documents().get(documentId=sid).execute()
        out = []
        for p in _list_paragraphs(doc):
            style = p["style"]
            if style == "TITLE":
                level = 0
            elif style.startswith("HEADING_"):
                level = int(style.split("_")[1])
            else:
                continue
            out.append({"index": p["index"], "level": level, "text": p["text"]})
        return out

    @mcp.tool
    def docs_replace_paragraph(index: int, text: str) -> dict:
        """Replace paragraph N's text. Preserves the paragraph's existing style."""
        sid = auth.require_active("docs")
        svc = auth.docs_service()
        p = _paragraph_at(svc, sid, index)
        requests: list[dict] = []
        if p["endIndex"] - p["startIndex"] > 1:
            # keep the trailing newline so the paragraph break survives
            requests.append(
                {
                    "deleteContentRange": {
                        "range": {"startIndex": p["startIndex"], "endIndex": p["endIndex"] - 1}
                    }
                }
            )
        if text:
            requests.append(
                {"insertText": {"location": {"index": p["startIndex"]}, "text": text}}
            )
        if requests:
            svc.documents().batchUpdate(
                documentId=sid, body={"requests": requests}
            ).execute()
        return {"replaced": index}

    @mcp.tool
    def docs_insert_paragraph(
        after: int | None = None,
        before: int | None = None,
        text: str = "",
        style: str = "NORMAL_TEXT",
    ) -> dict:
        """Insert a new paragraph after or before paragraph N. Style defaults to
        NORMAL_TEXT (a plain body paragraph); other choices: TITLE, SUBTITLE,
        HEADING_1..HEADING_6."""
        if (after is None) == (before is None):
            raise ValueError("pass exactly one of after= or before=")
        _check_style(style)
        sid = auth.require_active("docs")
        svc = auth.docs_service()
        anchor = _paragraph_at(svc, sid, after if after is not None else before)
        if after is not None:
            insert_at = anchor["endIndex"] - 1  # just before anchor's trailing \n
            body_text = "\n" + text
            style_start = insert_at + 1
        else:
            insert_at = anchor["startIndex"]
            body_text = text + "\n"
            style_start = insert_at
        requests: list[dict] = [
            {"insertText": {"location": {"index": insert_at}, "text": body_text}},
            {
                "updateParagraphStyle": {
                    "range": {
                        "startIndex": style_start,
                        "endIndex": style_start + max(len(text), 1),
                    },
                    "paragraphStyle": {"namedStyleType": style},
                    "fields": "namedStyleType",
                }
            },
        ]
        svc.documents().batchUpdate(documentId=sid, body={"requests": requests}).execute()
        return {"inserted_at": insert_at}

    @mcp.tool
    def docs_append_paragraph(text: str = "", style: str = "NORMAL_TEXT") -> dict:
        """Add a new paragraph at the very end of the document. Style defaults to
        NORMAL_TEXT; other choices: TITLE, SUBTITLE, HEADING_1..HEADING_6."""
        _check_style(style)
        sid = auth.require_active("docs")
        svc = auth.docs_service()
        paragraphs = _list_paragraphs(svc.documents().get(documentId=sid).execute())
        last = paragraphs[-1]
        if not last["text"]:
            # Last paragraph is empty — fill it instead of creating an empty
            # paragraph between the previous text and our new content.
            insert_at = last["startIndex"]
            body_text = text
            new_range = (insert_at, insert_at + max(len(text), 1))
        else:
            insert_at = last["endIndex"] - 1  # just before the final \n
            body_text = "\n" + text
            new_range = (insert_at + 1, insert_at + 1 + max(len(text), 1))
        start, end = new_range
        requests: list[dict] = [
            {"insertText": {"location": {"index": insert_at}, "text": body_text}},
            {
                "updateParagraphStyle": {
                    "range": {"startIndex": start, "endIndex": end},
                    "paragraphStyle": {"namedStyleType": style},
                    "fields": "namedStyleType",
                }
            },
        ]
        svc.documents().batchUpdate(documentId=sid, body={"requests": requests}).execute()
        return {"appended_at": insert_at}

    @mcp.tool
    def docs_delete_paragraph(index: int) -> dict:
        """Delete paragraph N (and its trailing paragraph break)."""
        sid = auth.require_active("docs")
        svc = auth.docs_service()
        p = _paragraph_at(svc, sid, index)
        # Can't delete the very last paragraph break in a doc; clear text instead.
        paragraphs = _list_paragraphs(svc.documents().get(documentId=sid).execute())
        is_last = index == len(paragraphs)
        if is_last and p["endIndex"] - p["startIndex"] <= 1:
            return {"deleted": index, "note": "already empty"}
        if is_last:
            end = p["endIndex"] - 1
        else:
            end = p["endIndex"]
        svc.documents().batchUpdate(
            documentId=sid,
            body={
                "requests": [
                    {
                        "deleteContentRange": {
                            "range": {"startIndex": p["startIndex"], "endIndex": end}
                        }
                    }
                ]
            },
        ).execute()
        return {"deleted": index}

    @mcp.tool
    def docs_get_tables() -> list:
        """All tables in the doc as {index, rows, cols, contents}. `contents` is a
        2-D list of cell texts. `index` is 1-based among tables."""
        sid = auth.require_active("docs")
        doc = auth.docs_service().documents().get(documentId=sid).execute()
        out = []
        for ti, el in _walk_tables(doc):
            tbl = el["table"]
            out.append(
                {
                    "index": ti,
                    "rows": tbl["rows"],
                    "cols": tbl["columns"],
                    "contents": _table_contents(el),
                }
            )
        return out

    @mcp.tool
    def docs_insert_table(
        rows: int,
        cols: int,
        after: int | None = None,
        before: int | None = None,
        data: list[list] | None = None,
    ) -> dict:
        """Insert a rows x cols table after or before paragraph N. If `data` is
        provided (2-D list), the cells are populated in the same call."""
        if (after is None) == (before is None):
            raise ValueError("pass exactly one of after= or before=")
        if rows < 1 or cols < 1:
            raise ValueError("rows and cols must be >= 1")
        sid = auth.require_active("docs")
        svc = auth.docs_service()
        anchor = _paragraph_at(svc, sid, after if after is not None else before)
        # insertTable splits at the given index; using endIndex-1 for "after"
        # places it just before the anchor's trailing \n (so the table sits
        # between the anchor and whatever followed). before=1 special-cases to
        # index 1.
        location = anchor["endIndex"] - 1 if after is not None else max(anchor["startIndex"] - 1, 1)
        before_count = sum(1 for _ in _walk_tables(svc.documents().get(documentId=sid).execute()))
        svc.documents().batchUpdate(
            documentId=sid,
            body={"requests": [
                {"insertTable": {"rows": rows, "columns": cols, "location": {"index": location}}}
            ]},
        ).execute()
        doc = svc.documents().get(documentId=sid).execute()
        tables = list(_walk_tables(doc))
        if len(tables) != before_count + 1:
            raise RuntimeError(
                f"expected 1 new table, found {len(tables) - before_count}"
            )
        # The new table is the one whose startIndex is closest to `location`
        # (ties broken by earlier-in-doc). Identify it by scanning for the
        # table at or just past `location`.
        new_ti, new_el = next(
            (ti, el) for ti, el in tables if el.get("startIndex", 0) >= location
        )
        if data:
            _fill_cells(svc, sid, new_el, data)
        return {"table_index": new_ti, "rows": rows, "cols": cols}

    @mcp.tool
    def docs_fill_table(table_index: int, data: list[list]) -> dict:
        """Overwrite cells of an existing table (1-based among tables). `data` is
        a 2-D list; only cells covered by `data` are touched."""
        sid = auth.require_active("docs")
        svc = auth.docs_service()
        doc = svc.documents().get(documentId=sid).execute()
        target = next((el for ti, el in _walk_tables(doc) if ti == table_index), None)
        if target is None:
            raise ValueError(f"table_index {table_index} not found")
        _fill_cells(svc, sid, target, data)
        return {"filled_table": table_index}

    @mcp.tool
    def docs_format_range(
        paragraph: int,
        start: int | None = None,
        end: int | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        strikethrough: bool | None = None,
        font: str | None = None,
        size: float | None = None,
        color: list | None = None,
        link: str | None = None,
    ) -> dict:
        """Apply inline character formatting to a range within paragraph N.

        start/end are 0-based character offsets within the paragraph (end exclusive).
        Both default to spanning the whole paragraph. color is [r, g, b] 0-255.
        Only attributes passed in are changed; others are left as they were.
        """
        style, fields = _text_style_payload(
            bold=bold,
            italic=italic,
            underline=underline,
            strikethrough=strikethrough,
            font=font,
            size=size,
            color=color,
            link=link,
        )
        if not fields:
            raise ValueError(
                "no formatting changes requested; pass at least one of "
                "bold/italic/underline/strikethrough/font/size/color/link"
            )
        sid = auth.require_active("docs")
        svc = auth.docs_service()
        p = _paragraph_at(svc, sid, paragraph)
        s = 0 if start is None else start
        e = len(p["text"]) if end is None else end
        if not 0 <= s < e <= len(p["text"]):
            raise ValueError(
                f"start={s}, end={e} invalid for paragraph length {len(p['text'])}"
            )
        abs_start = p["startIndex"] + s
        abs_end = p["startIndex"] + e
        svc.documents().batchUpdate(
            documentId=sid,
            body={"requests": [
                {"updateTextStyle": {
                    "range": {"startIndex": abs_start, "endIndex": abs_end},
                    "textStyle": style,
                    "fields": fields,
                }}
            ]},
        ).execute()
        return {"paragraph": paragraph, "range": [s, e], "fields": fields}

    @mcp.tool
    def docs_find_replace(paragraph: int, find: str, replace: str) -> dict:
        """Replace all occurrences of `find` with `replace` within paragraph N.

        Scoped to a single paragraph by design (a global find/replace at MCP
        level is a foot-gun). Returns the count of replacements."""
        if not find:
            raise ValueError("find cannot be empty")
        sid = auth.require_active("docs")
        svc = auth.docs_service()
        p = _paragraph_at(svc, sid, paragraph)
        text = p["text"]
        positions: list[int] = []
        i = 0
        while True:
            j = text.find(find, i)
            if j < 0:
                break
            positions.append(j)
            i = j + len(find)
        if not positions:
            return {"paragraph": paragraph, "count": 0}
        requests: list[dict] = []
        for pos in reversed(positions):
            s = p["startIndex"] + pos
            e = s + len(find)
            requests.append(
                {"deleteContentRange": {"range": {"startIndex": s, "endIndex": e}}}
            )
            if replace:
                requests.append(
                    {"insertText": {"location": {"index": s}, "text": replace}}
                )
        svc.documents().batchUpdate(documentId=sid, body={"requests": requests}).execute()
        return {"paragraph": paragraph, "count": len(positions)}

    @mcp.tool
    def docs_set_style(index: int, style: str) -> dict:
        """Set paragraph N's style (NORMAL_TEXT, TITLE, SUBTITLE, HEADING_1..HEADING_6)."""
        _check_style(style)
        sid = auth.require_active("docs")
        svc = auth.docs_service()
        p = _paragraph_at(svc, sid, index)
        svc.documents().batchUpdate(
            documentId=sid,
            body={
                "requests": [
                    {
                        "updateParagraphStyle": {
                            "range": {"startIndex": p["startIndex"], "endIndex": p["endIndex"]},
                            "paragraphStyle": {"namedStyleType": style},
                            "fields": "namedStyleType",
                        }
                    }
                ]
            },
        ).execute()
        return {"styled": index, "style": style}

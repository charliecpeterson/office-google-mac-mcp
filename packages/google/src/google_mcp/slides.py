"""Slides tools — index-and-objectId-anchored editing on the user's open deck.

Slides are addressed by 1-based slide index; the page elements within a slide
(text boxes, placeholders, images, etc.) are addressed by their `objectId`,
returned from `slides_read`. Edits via the Slides v1 API show up live in the
open browser tab through Google's realtime sync.
"""

from google_mcp import auth

INSTRUCTIONS = """\
You are working live on the user's open Google Slides deck, via the Slides API.
Edits show up in their open browser tab through Google's realtime sync.

First, point the server at the right deck: call slides_set_active(url) with the
URL from the browser address bar. After that all other tools default to that
deck. Confirm with slides_status.

Slides are addressed by 1-based slide index; the page elements within a slide
(text boxes, placeholders, images) are addressed by their `objectId`. Read
structure first: slides_list gives a per-slide summary, slides_read(N) returns
that slide's element objectIds and text.

Edit by index/objectId: slides_add(after=N or before=N, layout=...) creates a
new slide, slides_duplicate(N) clones, slides_delete(N), slides_move(from, to).
For text in a shape: slides_set_text(slide_index=N, object_id=..., text=...).

Inline formatting: slides_format_text(slide=N, object_id, bold/italic/font/
size/color/link) — use size= to shrink overflowing body text (the most common
fix when text exceeds its placeholder).

Images: slides_insert_image(slide=N, url, x, y, width, height) inserts from a
public URL. Position in points (72pt = 1in); default slide is 720x405pt (16:9).

Tables: slides_insert_table(slide=N, rows, cols, x, y, width, height, data=)
inserts and optionally fills in one call. slides_fill_table overwrites cells
of an existing table; slides_get_tables reads all tables back.

Layouts: BLANK, CAPTION_ONLY, TITLE, TITLE_AND_BODY, TITLE_AND_TWO_COLUMNS,
TITLE_ONLY, SECTION_HEADER, SECTION_TITLE_AND_DESCRIPTION, ONE_COLUMN_TEXT,
MAIN_POINT, BIG_NUMBER.
"""

_LAYOUTS = {
    "BLANK",
    "CAPTION_ONLY",
    "TITLE",
    "TITLE_AND_BODY",
    "TITLE_AND_TWO_COLUMNS",
    "TITLE_ONLY",
    "SECTION_HEADER",
    "SECTION_TITLE_AND_DESCRIPTION",
    "ONE_COLUMN_TEXT",
    "MAIN_POINT",
    "BIG_NUMBER",
}


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
    """Build (textStyle, fields-mask) for an updateTextStyle request. Slides uses
    the same textStyle schema as Docs, but the request field is named `style`."""
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
        style["fontFamily"] = font
        fields.append("fontFamily")
    if size is not None:
        style["fontSize"] = {"magnitude": float(size), "unit": "PT"}
        fields.append("fontSize")
    if color is not None:
        if len(color) != 3:
            raise ValueError("color must be [r, g, b] with 0-255 values")
        r, g, b = color
        style["foregroundColor"] = {
            "opaqueColor": {"rgbColor": {"red": r / 255, "green": g / 255, "blue": b / 255}}
        }
        fields.append("foregroundColor")
    if link is not None:
        style["link"] = {"url": link}
        fields.append("link")
    return style, ",".join(fields)


_PT_PER_EMU = 1.0 / 9525.0


def _to_pt(dim: dict | None) -> float | None:
    if not dim:
        return None
    mag = dim.get("magnitude", 0)
    return mag * _PT_PER_EMU if dim.get("unit") == "EMU" else mag


def _element_properties(slide_object_id: str, x: float, y: float, w: float, h: float) -> dict:
    return {
        "pageObjectId": slide_object_id,
        "size": {
            "width": {"magnitude": w, "unit": "PT"},
            "height": {"magnitude": h, "unit": "PT"},
        },
        "transform": {
            "scaleX": 1,
            "scaleY": 1,
            "translateX": x,
            "translateY": y,
            "unit": "PT",
        },
    }


def _cell_text(cell: dict) -> str:
    text = ""
    for te in cell.get("text", {}).get("textElements", []):
        if "textRun" in te:
            text += te["textRun"].get("content", "")
    return text.rstrip("\n")


def _slides(pres: dict) -> list[dict]:
    return pres.get("slides", [])


def _slide_at(svc, pres_id: str, index: int) -> dict:
    slides = _slides(svc.presentations().get(presentationId=pres_id).execute())
    if not 1 <= index <= len(slides):
        raise ValueError(f"slide index {index} out of range (1..{len(slides)})")
    return slides[index - 1]


def _shape_text(shape: dict) -> str:
    text = ""
    for te in shape.get("text", {}).get("textElements", []):
        if "textRun" in te:
            text += te["textRun"].get("content", "")
    return text


def _read_element(el: dict) -> dict:
    out = {"objectId": el["objectId"]}
    if "shape" in el:
        shape = el["shape"]
        out["type"] = "shape"
        out["shapeType"] = shape.get("shapeType")
        out["text"] = _shape_text(shape).rstrip("\n")
        ph = shape.get("placeholder")
        if ph:
            out["placeholder"] = ph.get("type")
    elif "image" in el:
        out["type"] = "image"
        out["url"] = el["image"].get("contentUrl")
    elif "table" in el:
        out["type"] = "table"
        out["rows"] = el["table"]["rows"]
        out["cols"] = el["table"]["columns"]
    elif "line" in el:
        out["type"] = "line"
    else:
        out["type"] = "other"
    return out


def _slide_title(slide: dict) -> str | None:
    for el in slide.get("pageElements", []):
        if "shape" not in el:
            continue
        ph = el["shape"].get("placeholder", {})
        if ph.get("type") in ("TITLE", "CENTERED_TITLE"):
            return _shape_text(el["shape"]).rstrip("\n") or None
    return None


def register(mcp):
    @mcp.tool
    def slides_set_active(url_or_id: str) -> str:
        """Point the server at a Slides deck by URL (or ID). Persists across sessions."""
        return f"active presentation set: {auth.set_active('slides', url_or_id)}"

    @mcp.tool
    def slides_status() -> dict:
        """The active deck's title, ID, URL, and slide count."""
        sid = auth.require_active("slides")
        pres = auth.slides_service().presentations().get(presentationId=sid).execute()
        return {
            "id": sid,
            "title": pres.get("title"),
            "url": f"https://docs.google.com/presentation/d/{sid}/edit",
            "slides": len(_slides(pres)),
        }

    @mcp.tool
    def slides_list() -> list:
        """Per-slide summary: {index, objectId, title, elements}."""
        sid = auth.require_active("slides")
        pres = auth.slides_service().presentations().get(presentationId=sid).execute()
        return [
            {
                "index": i,
                "objectId": s["objectId"],
                "title": _slide_title(s),
                "elements": len(s.get("pageElements", [])),
            }
            for i, s in enumerate(_slides(pres), 1)
        ]

    @mcp.tool
    def slides_read(index: int) -> dict:
        """One slide's elements: {index, objectId, elements: [{objectId, type, ...}]}."""
        sid = auth.require_active("slides")
        slide = _slide_at(auth.slides_service(), sid, index)
        return {
            "index": index,
            "objectId": slide["objectId"],
            "elements": [_read_element(el) for el in slide.get("pageElements", [])],
        }

    @mcp.tool
    def slides_add(
        after: int | None = None,
        before: int | None = None,
        layout: str = "TITLE_AND_BODY",
    ) -> dict:
        """Insert a new slide after or before slide N, using a predefined layout."""
        if (after is None) == (before is None):
            raise ValueError("pass exactly one of after= or before=")
        if layout not in _LAYOUTS:
            raise ValueError(f"unknown layout {layout!r}; choices: {sorted(_LAYOUTS)}")
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        anchor = after if after is not None else before
        # API insertionIndex is 0-based position where the new slide ends up.
        insertion = anchor if after is not None else anchor - 1
        resp = svc.presentations().batchUpdate(
            presentationId=sid,
            body={"requests": [
                {"createSlide": {
                    "insertionIndex": insertion,
                    "slideLayoutReference": {"predefinedLayout": layout},
                }}
            ]},
        ).execute()
        return {
            "new_slide_index": insertion + 1,
            "objectId": resp["replies"][0]["createSlide"]["objectId"],
        }

    @mcp.tool
    def slides_duplicate(index: int) -> dict:
        """Duplicate slide N. The clone lands right after the original."""
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        slide = _slide_at(svc, sid, index)
        resp = svc.presentations().batchUpdate(
            presentationId=sid,
            body={"requests": [{"duplicateObject": {"objectId": slide["objectId"]}}]},
        ).execute()
        return {
            "duplicated_from": index,
            "new_objectId": resp["replies"][0]["duplicateObject"]["objectId"],
        }

    @mcp.tool
    def slides_delete(index: int) -> dict:
        """Delete slide N."""
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        slide = _slide_at(svc, sid, index)
        svc.presentations().batchUpdate(
            presentationId=sid,
            body={"requests": [{"deleteObject": {"objectId": slide["objectId"]}}]},
        ).execute()
        return {"deleted": index}

    @mcp.tool
    def slides_move(from_index: int, to_index: int) -> dict:
        """Move slide `from_index` so it ends up at 1-based position `to_index`."""
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        slide = _slide_at(svc, sid, from_index)
        svc.presentations().batchUpdate(
            presentationId=sid,
            body={"requests": [
                {"updateSlidesPosition": {
                    "slideObjectIds": [slide["objectId"]],
                    "insertionIndex": to_index - 1,
                }}
            ]},
        ).execute()
        return {"moved": from_index, "to": to_index}

    @mcp.tool
    def slides_format_text(
        slide_index: int,
        object_id: str,
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
        """Inline text formatting on a shape on a slide.

        start/end are 0-based character offsets; both None = the whole shape.
        Pass only the attributes you want to change. color is [r, g, b] 0-255;
        size is points. Use this to shrink overflowing body text (e.g. size=14).
        """
        style, fields = _text_style_payload(
            bold=bold, italic=italic, underline=underline, strikethrough=strikethrough,
            font=font, size=size, color=color, link=link,
        )
        if not fields:
            raise ValueError(
                "no formatting changes requested; pass at least one of "
                "bold/italic/underline/strikethrough/font/size/color/link"
            )
        if (start is None) != (end is None):
            raise ValueError("pass both start and end, or neither")
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        slide = _slide_at(svc, sid, slide_index)
        target = next(
            (el for el in slide.get("pageElements", []) if el["objectId"] == object_id),
            None,
        )
        if target is None:
            raise ValueError(f"object {object_id!r} not found on slide {slide_index}")
        if "shape" not in target:
            raise ValueError(f"object {object_id!r} is not a shape (cannot hold text style)")
        text_range = (
            {"type": "FIXED_RANGE", "startIndex": start, "endIndex": end}
            if start is not None
            else {"type": "ALL"}
        )
        svc.presentations().batchUpdate(
            presentationId=sid,
            body={"requests": [
                {"updateTextStyle": {
                    "objectId": object_id,
                    "textRange": text_range,
                    "style": style,
                    "fields": fields,
                }}
            ]},
        ).execute()
        return {"slide": slide_index, "objectId": object_id, "fields": fields}

    @mcp.tool
    def slides_insert_image(
        slide_index: int,
        url: str,
        x: float = 72.0,
        y: float = 100.0,
        width: float = 400.0,
        height: float = 300.0,
    ) -> dict:
        """Insert an image on a slide from a publicly accessible URL. Position
        and size in points (1 inch = 72pt). Returns the new image's objectId.

        Note: Google's backend fetches the URL — some hosts (Wikimedia
        thumbnails, hot-linked images) get rejected even when the URL works in
        a browser. If you get a "provided image was not found" error, try a
        different host (direct PNG/JPG, not a redirect, not behind anti-bot)."""
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        slide = _slide_at(svc, sid, slide_index)
        resp = svc.presentations().batchUpdate(
            presentationId=sid,
            body={"requests": [
                {"createImage": {
                    "url": url,
                    "elementProperties": _element_properties(
                        slide["objectId"], x, y, width, height
                    ),
                }}
            ]},
        ).execute()
        return {
            "slide": slide_index,
            "objectId": resp["replies"][0]["createImage"]["objectId"],
        }

    @mcp.tool
    def slides_insert_table(
        slide_index: int,
        rows: int,
        cols: int,
        x: float = 72.0,
        y: float = 100.0,
        width: float = 600.0,
        height: float = 200.0,
        data: list[list] | None = None,
    ) -> dict:
        """Insert a rows x cols table on a slide. If `data` is provided (2-D
        list), cells are populated in the same call. Returns the table's objectId."""
        if rows < 1 or cols < 1:
            raise ValueError("rows and cols must be >= 1")
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        slide = _slide_at(svc, sid, slide_index)
        resp = svc.presentations().batchUpdate(
            presentationId=sid,
            body={"requests": [
                {"createTable": {
                    "rows": rows,
                    "columns": cols,
                    "elementProperties": _element_properties(
                        slide["objectId"], x, y, width, height
                    ),
                }}
            ]},
        ).execute()
        table_oid = resp["replies"][0]["createTable"]["objectId"]
        if data:
            requests = []
            for r_i, row in enumerate(data):
                for c_i, val in enumerate(row):
                    text = str(val)
                    if not text:
                        continue
                    requests.append(
                        {"insertText": {
                            "objectId": table_oid,
                            "cellLocation": {"rowIndex": r_i, "columnIndex": c_i},
                            "text": text,
                            "insertionIndex": 0,
                        }}
                    )
            if requests:
                svc.presentations().batchUpdate(
                    presentationId=sid, body={"requests": requests}
                ).execute()
        return {"slide": slide_index, "objectId": table_oid, "rows": rows, "cols": cols}

    @mcp.tool
    def slides_fill_table(
        slide_index: int, table_object_id: str, data: list[list]
    ) -> dict:
        """Overwrite cells of an existing table. Only cells covered by `data` are touched."""
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        slide = _slide_at(svc, sid, slide_index)
        target = next(
            (el for el in slide.get("pageElements", []) if el["objectId"] == table_object_id),
            None,
        )
        if target is None or "table" not in target:
            raise ValueError(
                f"table {table_object_id!r} not found on slide {slide_index}"
            )
        requests: list[dict] = []
        for r_i, row in enumerate(target["table"].get("tableRows", [])):
            if r_i >= len(data):
                break
            for c_i, cell in enumerate(row.get("tableCells", [])):
                if c_i >= len(data[r_i]):
                    continue
                text = str(data[r_i][c_i])
                if not text:
                    continue
                if _cell_text(cell):
                    requests.append(
                        {"deleteText": {
                            "objectId": table_object_id,
                            "cellLocation": {"rowIndex": r_i, "columnIndex": c_i},
                            "textRange": {"type": "ALL"},
                        }}
                    )
                requests.append(
                    {"insertText": {
                        "objectId": table_object_id,
                        "cellLocation": {"rowIndex": r_i, "columnIndex": c_i},
                        "text": text,
                        "insertionIndex": 0,
                    }}
                )
        if requests:
            svc.presentations().batchUpdate(
                presentationId=sid, body={"requests": requests}
            ).execute()
        return {"slide": slide_index, "objectId": table_object_id}

    @mcp.tool
    def slides_get_tables() -> list:
        """All tables across all slides as {slide, objectId, rows, cols, contents}."""
        sid = auth.require_active("slides")
        pres = auth.slides_service().presentations().get(presentationId=sid).execute()
        out = []
        for i, slide in enumerate(_slides(pres), 1):
            for el in slide.get("pageElements", []):
                if "table" not in el:
                    continue
                tbl = el["table"]
                contents = [
                    [_cell_text(c) for c in row.get("tableCells", [])]
                    for row in tbl.get("tableRows", [])
                ]
                out.append(
                    {
                        "slide": i,
                        "objectId": el["objectId"],
                        "rows": tbl["rows"],
                        "cols": tbl["columns"],
                        "contents": contents,
                    }
                )
        return out

    @mcp.tool
    def slides_set_text(slide_index: int, object_id: str, text: str) -> dict:
        """Replace all text in a shape on a slide (object_id from slides_read)."""
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        slide = _slide_at(svc, sid, slide_index)
        target = next(
            (el for el in slide.get("pageElements", []) if el["objectId"] == object_id),
            None,
        )
        if target is None:
            raise ValueError(f"object {object_id!r} not found on slide {slide_index}")
        if "shape" not in target:
            raise ValueError(f"object {object_id!r} is not a shape (cannot hold text)")
        current = _shape_text(target["shape"])
        requests: list[dict] = []
        if current:
            requests.append(
                {"deleteText": {"objectId": object_id, "textRange": {"type": "ALL"}}}
            )
        if text:
            requests.append(
                {"insertText": {"objectId": object_id, "insertionIndex": 0, "text": text}}
            )
        if requests:
            svc.presentations().batchUpdate(
                presentationId=sid, body={"requests": requests}
            ).execute()
        return {"slide": slide_index, "objectId": object_id, "text": text}

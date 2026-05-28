"""Slides tools — index-and-objectId-anchored editing on the user's open deck.

Slides are addressed by 1-based slide index; the page elements within a slide
(text boxes, placeholders, images, etc.) are addressed by their `objectId`,
returned from `slides_read`. Edits via the Slides v1 API show up live in the
open browser tab through Google's realtime sync.
"""

import urllib.request

from fastmcp.utilities.types import Image

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
slides_format_cell(slide, table_object_id, row, col, ...) bolds/styles a
single table cell (e.g. header row).

Custom layout: slides_insert_text_box(slide, x, y, w, h, text=) drops a sized
text box for captions / side-by-side compositions the built-in layouts
don't cover. slides_set_element_bounds(slide, object_id, x=, y=, width=,
height=) moves/resizes any element; slides_delete_element removes one.
slides_read includes each element's x/y/width/height in points so positions
are visible without a screenshot.

Speaker notes: slides_set_notes(slide=N, text) sets the notes, slides_get_notes
reads them.

Composite: slides_add_content_slide(title, bullets, after=N or before=N)
inserts a TITLE_AND_BODY slide and fills title + body with real bullets in
one call.

Visual verification: slides_thumbnail(slide=N) returns a PNG of the rendered
slide (cheaper than a full export). Use after a layout edit to confirm bounds
look right.

Escape hatch: slides_batch_update(requests) accepts a raw Slides API request
list.

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


_PT_PER_EMU = 1.0 / 12700.0  # Google Slides uses 914400 EMU/inch, 72 PT/inch


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


def _element_bounds_pt(el: dict) -> dict | None:
    """Displayed (x, y, width, height) in points for a page element, computed
    from its size + transform. Returns None if size is missing."""
    size = el.get("size", {})
    inherent_w = _to_pt(size.get("width"))
    inherent_h = _to_pt(size.get("height"))
    if inherent_w is None or inherent_h is None:
        return None
    t = el.get("transform", {})
    tx = t.get("translateX", 0)
    ty = t.get("translateY", 0)
    if t.get("unit") == "EMU":
        tx *= _PT_PER_EMU
        ty *= _PT_PER_EMU
    sx = t.get("scaleX", 1)
    sy = t.get("scaleY", 1)
    return {
        "x": round(tx, 2),
        "y": round(ty, 2),
        "width": round(inherent_w * sx, 2),
        "height": round(inherent_h * sy, 2),
    }


def _read_element(el: dict) -> dict:
    out = {"objectId": el["objectId"]}
    bounds = _element_bounds_pt(el)
    if bounds is not None:
        out.update(bounds)
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
        list), cells are populated in the same call. Returns the table's objectId.

        Note: the Slides API picks default cell sizes for createTable; the
        width/height passed here set the table's bounding box position but the
        rendered table size is auto-computed from row/column count and content.
        Use slides_set_element_bounds afterwards if you need to enforce a size."""
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
    def slides_set_notes(slide_index: int, text: str) -> dict:
        """Set the speaker notes for slide N (replaces existing notes)."""
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        slide = _slide_at(svc, sid, slide_index)
        notes_page = slide.get("slideProperties", {}).get("notesPage", {})
        notes_oid = (
            notes_page.get("notesProperties", {}).get("speakerNotesObjectId")
        )
        if not notes_oid:
            raise RuntimeError(
                f"no speaker-notes shape found on slide {slide_index}"
            )
        current = ""
        for el in notes_page.get("pageElements", []):
            if el.get("objectId") == notes_oid and "shape" in el:
                current = _shape_text(el["shape"])
                break
        requests: list[dict] = []
        if current:
            requests.append(
                {"deleteText": {"objectId": notes_oid, "textRange": {"type": "ALL"}}}
            )
        if text:
            requests.append(
                {"insertText": {
                    "objectId": notes_oid,
                    "insertionIndex": 0,
                    "text": text,
                }}
            )
        if requests:
            svc.presentations().batchUpdate(
                presentationId=sid, body={"requests": requests}
            ).execute()
        return {"slide": slide_index, "objectId": notes_oid}

    @mcp.tool
    def slides_set_element_bounds(
        slide_index: int,
        object_id: str,
        x: float | None = None,
        y: float | None = None,
        width: float | None = None,
        height: float | None = None,
    ) -> dict:
        """Move and/or resize an element on a slide. Any parameter omitted keeps
        its current value. Coordinates and sizes in points (1in = 72pt). Existing
        rotation/shear is preserved."""
        if x is None and y is None and width is None and height is None:
            raise ValueError("pass at least one of x, y, width, height")
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        slide = _slide_at(svc, sid, slide_index)
        target = next(
            (el for el in slide.get("pageElements", []) if el["objectId"] == object_id),
            None,
        )
        if target is None:
            raise ValueError(f"object {object_id!r} not found on slide {slide_index}")
        size = target.get("size", {})
        inherent_w = _to_pt(size.get("width")) or 1.0
        inherent_h = _to_pt(size.get("height")) or 1.0
        cur = target.get("transform", {})
        cur_tx = cur.get("translateX", 0)
        cur_ty = cur.get("translateY", 0)
        if cur.get("unit") == "EMU":
            cur_tx *= _PT_PER_EMU
            cur_ty *= _PT_PER_EMU
        new_tx = cur_tx if x is None else x
        new_ty = cur_ty if y is None else y
        new_sx = cur.get("scaleX", 1) if width is None else (width / inherent_w)
        new_sy = cur.get("scaleY", 1) if height is None else (height / inherent_h)
        svc.presentations().batchUpdate(
            presentationId=sid,
            body={"requests": [
                {"updatePageElementTransform": {
                    "objectId": object_id,
                    "applyMode": "ABSOLUTE",
                    "transform": {
                        "scaleX": new_sx,
                        "scaleY": new_sy,
                        "translateX": new_tx,
                        "translateY": new_ty,
                        "shearX": cur.get("shearX", 0),
                        "shearY": cur.get("shearY", 0),
                        "unit": "PT",
                    },
                }}
            ]},
        ).execute()
        return {
            "slide": slide_index,
            "objectId": object_id,
            "x": new_tx,
            "y": new_ty,
            "width": new_sx * inherent_w,
            "height": new_sy * inherent_h,
        }

    @mcp.tool
    def slides_delete_element(slide_index: int, object_id: str) -> dict:
        """Delete a single page element (shape/image/table/line) from a slide."""
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        slide = _slide_at(svc, sid, slide_index)
        target = next(
            (el for el in slide.get("pageElements", []) if el["objectId"] == object_id),
            None,
        )
        if target is None:
            raise ValueError(f"object {object_id!r} not found on slide {slide_index}")
        svc.presentations().batchUpdate(
            presentationId=sid,
            body={"requests": [{"deleteObject": {"objectId": object_id}}]},
        ).execute()
        return {"slide": slide_index, "deleted": object_id}

    @mcp.tool
    def slides_insert_text_box(
        slide_index: int,
        x: float,
        y: float,
        width: float,
        height: float,
        text: str = "",
    ) -> dict:
        """Insert a TEXT_BOX shape on a slide at (x, y) with the given size, and
        optionally pre-fill it with text. Use this for custom layouts (captions,
        image-plus-text side-by-side, etc.) where the built-in layouts don't fit."""
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        slide = _slide_at(svc, sid, slide_index)
        resp = svc.presentations().batchUpdate(
            presentationId=sid,
            body={"requests": [
                {"createShape": {
                    "shapeType": "TEXT_BOX",
                    "elementProperties": _element_properties(
                        slide["objectId"], x, y, width, height
                    ),
                }}
            ]},
        ).execute()
        shape_oid = resp["replies"][0]["createShape"]["objectId"]
        if text:
            svc.presentations().batchUpdate(
                presentationId=sid,
                body={"requests": [
                    {"insertText": {
                        "objectId": shape_oid,
                        "insertionIndex": 0,
                        "text": text,
                    }}
                ]},
            ).execute()
        return {"slide": slide_index, "objectId": shape_oid}

    @mcp.tool
    def slides_format_cell(
        slide_index: int,
        table_object_id: str,
        row: int,
        col: int,
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
        """Inline text formatting on a single table cell (row, col are 0-based).
        Common use: bold the header row by calling once per cell with bold=True."""
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
            (el for el in slide.get("pageElements", []) if el["objectId"] == table_object_id),
            None,
        )
        if target is None or "table" not in target:
            raise ValueError(
                f"table {table_object_id!r} not found on slide {slide_index}"
            )
        text_range = (
            {"type": "FIXED_RANGE", "startIndex": start, "endIndex": end}
            if start is not None
            else {"type": "ALL"}
        )
        svc.presentations().batchUpdate(
            presentationId=sid,
            body={"requests": [
                {"updateTextStyle": {
                    "objectId": table_object_id,
                    "cellLocation": {"rowIndex": row, "columnIndex": col},
                    "textRange": text_range,
                    "style": style,
                    "fields": fields,
                }}
            ]},
        ).execute()
        return {
            "slide": slide_index,
            "table": table_object_id,
            "row": row,
            "col": col,
            "fields": fields,
        }

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

    @mcp.tool
    def slides_add_content_slide(
        title: str,
        bullets: list,
        after: int | None = None,
        before: int | None = None,
    ) -> dict:
        """Composite: insert a TITLE_AND_BODY slide, set the title, and fill the
        body with one bullet per item in `bullets` (real bullet rendering via the
        Slides bullet API, not just lines with dots). Mirrors ppt_add_content_slide
        on the Office side."""
        if (after is None) == (before is None):
            raise ValueError("pass exactly one of after= or before=")
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        anchor = after if after is not None else before
        insertion = anchor if after is not None else anchor - 1
        resp = svc.presentations().batchUpdate(
            presentationId=sid,
            body={"requests": [
                {"createSlide": {
                    "insertionIndex": insertion,
                    "slideLayoutReference": {"predefinedLayout": "TITLE_AND_BODY"},
                }}
            ]},
        ).execute()
        new_oid = resp["replies"][0]["createSlide"]["objectId"]
        new_idx = insertion + 1
        pres = svc.presentations().get(presentationId=sid).execute()
        slide = _slides(pres)[new_idx - 1]
        title_oid = next(
            (el["objectId"] for el in slide.get("pageElements", [])
             if "shape" in el and el["shape"].get("placeholder", {}).get("type") == "TITLE"),
            None,
        )
        body_oid = next(
            (el["objectId"] for el in slide.get("pageElements", [])
             if "shape" in el and el["shape"].get("placeholder", {}).get("type") == "BODY"),
            None,
        )
        if not title_oid or not body_oid:
            raise RuntimeError(
                f"created slide {new_idx} missing TITLE or BODY placeholder"
            )
        body_text = "\n".join(str(b) for b in bullets)
        requests: list[dict] = [
            {"insertText": {"objectId": title_oid, "insertionIndex": 0, "text": title}},
        ]
        if body_text:
            requests.append(
                {"insertText": {"objectId": body_oid, "insertionIndex": 0, "text": body_text}}
            )
            requests.append(
                {"createParagraphBullets": {
                    "objectId": body_oid,
                    "textRange": {"type": "ALL"},
                    "bulletPreset": "BULLET_DISC_CIRCLE_SQUARE",
                }}
            )
        svc.presentations().batchUpdate(
            presentationId=sid, body={"requests": requests}
        ).execute()
        return {
            "new_slide_index": new_idx,
            "objectId": new_oid,
            "title_objectId": title_oid,
            "body_objectId": body_oid,
        }

    @mcp.tool
    def slides_get_notes(slide_index: int) -> str | None:
        """Read the speaker notes for slide N (None if empty)."""
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        slide = _slide_at(svc, sid, slide_index)
        notes_page = slide.get("slideProperties", {}).get("notesPage", {})
        notes_oid = (
            notes_page.get("notesProperties", {}).get("speakerNotesObjectId")
        )
        if not notes_oid:
            return None
        for el in notes_page.get("pageElements", []):
            if el.get("objectId") == notes_oid and "shape" in el:
                text = _shape_text(el["shape"]).rstrip("\n")
                return text or None
        return None

    @mcp.tool
    def slides_thumbnail(slide_index: int) -> Image:
        """Render slide N as a PNG thumbnail (the visual the agent would need a
        screenshot for in Office). Uses the Slides API's pages.getThumbnail —
        works on the doc's existing OAuth scope, no Drive elevation needed."""
        sid = auth.require_active("slides")
        svc = auth.slides_service()
        slide = _slide_at(svc, sid, slide_index)
        meta = (
            svc.presentations()
            .pages()
            .getThumbnail(presentationId=sid, pageObjectId=slide["objectId"])
            .execute()
        )
        with urllib.request.urlopen(meta["contentUrl"]) as resp:
            return Image(data=resp.read(), format="png")

    @mcp.tool
    def slides_batch_update(requests: list) -> dict:
        """Raw escape hatch: send a list of Slides API batchUpdate requests."""
        if not isinstance(requests, list) or not requests:
            raise ValueError("requests must be a non-empty list of API request dicts")
        sid = auth.require_active("slides")
        resp = (
            auth.slides_service()
            .presentations()
            .batchUpdate(presentationId=sid, body={"requests": requests})
            .execute()
        )
        return {"replies": resp.get("replies", [])}

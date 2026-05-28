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

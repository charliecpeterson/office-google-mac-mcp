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

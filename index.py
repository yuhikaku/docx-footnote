import json
import os
from typing import Annotated

from fastapi import Depends, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import Response
from fastapi.staticfiles import StaticFiles

from footnotes import apply_footnotes, find_unused_numbers
from security import require_allowed_ip

DOCX_MEDIA_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)

app = FastAPI()


@app.post("/api/footnotes", dependencies=[Depends(require_allowed_ip)])
async def create_footnotes(
    file: Annotated[UploadFile, File()],
    footnotes: Annotated[str, Form()],
) -> Response:
    """docx 内の注マーカーを脚注に置換して返す"""
    try:
        pairs = json.loads(footnotes)
        mapping = {str(k).strip(): str(v).strip() for k, v in pairs}
    except json.JSONDecodeError, TypeError, ValueError:
        raise HTTPException(
            status_code=400,
            detail="footnotes must be a JSON array of [number, text] pairs",
        )

    docx_bytes = await file.read()
    result, used = apply_footnotes(docx_bytes, mapping)

    footnote_nums = [int(k) for k in mapping if k.isdecimal()]
    missing = []
    if footnote_nums:
        max_num = max(footnote_nums)
        missing = find_unused_numbers(max_num, set(used))

    return Response(
        content=result,
        media_type=DOCX_MEDIA_TYPE,
        headers={
            "X-Footnote-Count": str(len(used)),
            "X-Missing-Numbers": ",".join(map(str, missing)),
        },
    )


if os.environ.get("SERVE_STATIC"):
    app.mount("/", StaticFiles(directory="public", html=True), name="public")

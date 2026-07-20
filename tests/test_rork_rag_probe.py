import base64
import io
import json
import sys
import time
import urllib.error
import urllib.request
import uuid
import zipfile

BASE_URL = "https://toolkit.rork.com"
AGENT_URL = f"{BASE_URL}/agent/chat"
TIMEOUT_SECONDS = 60


def log(event: str, **data):
    payload = {
        "event": event,
        "python": f"{sys.version_info.major}.{sys.version_info.minor}",
        "timestamp": time.time(),
        **data,
    }
    print("RORK_RAG_TEST " + json.dumps(payload, ensure_ascii=False, sort_keys=True), flush=True)


def request(url: str, payload=None, method="POST", timeout=TIMEOUT_SECONDS):
    body = None if payload is None else json.dumps(payload).encode("utf-8")
    headers = {
        "User-Agent": "rork-toolkit-rag-capability-probe/1.0",
        "Accept": "*/*",
    }
    if body is not None:
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    started = time.perf_counter()
    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            raw = response.read()
            return {
                "status": response.status,
                "headers": {k.lower(): v for k, v in response.headers.items()},
                "body": raw,
                "latency_ms": round((time.perf_counter() - started) * 1000, 2),
                "error": None,
            }
    except urllib.error.HTTPError as exc:
        return {
            "status": exc.code,
            "headers": {k.lower(): v for k, v in exc.headers.items()} if exc.headers else {},
            "body": exc.read(),
            "latency_ms": round((time.perf_counter() - started) * 1000, 2),
            "error": f"HTTPError: {exc}",
        }
    except Exception as exc:  # noqa: BLE001
        return {
            "status": 0,
            "headers": {},
            "body": b"",
            "latency_ms": round((time.perf_counter() - started) * 1000, 2),
            "error": f"{type(exc).__name__}: {exc}",
        }


def parse_stream(raw: bytes):
    frames = []
    for line in raw.decode("utf-8", errors="replace").splitlines():
        line = line.strip()
        if not line.startswith("data:"):
            continue
        data = line[5:].strip()
        if data == "[DONE]":
            frames.append({"type": "[DONE]"})
            continue
        try:
            frames.append(json.loads(data))
        except json.JSONDecodeError:
            frames.append({"type": "unparsed", "raw": data})
    return frames


def extract_text(frames):
    return "".join(
        str(frame.get("delta", ""))
        for frame in frames
        if frame.get("type") == "text-delta"
    )


def agent_request(messages, tools=None):
    payload = {"messages": messages}
    if tools is not None:
        payload["tools"] = tools
    result = request(AGENT_URL, payload)
    result["frames"] = parse_stream(result["body"])
    result["text"] = extract_text(result["frames"])
    return result


def data_url(media_type: str, raw: bytes):
    return f"data:{media_type};base64,{base64.b64encode(raw).decode('ascii')}"


def file_message(media_type: str, filename: str, raw: bytes, instruction: str):
    return {
        "id": "u-" + uuid.uuid4().hex[:12],
        "role": "user",
        "parts": [
            {
                "type": "file",
                "mediaType": media_type,
                "filename": filename,
                "url": data_url(media_type, raw),
            },
            {"type": "text", "text": instruction},
        ],
    }


def make_pdf(text: str):
    escaped = text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    stream = f"BT /F1 18 Tf 72 720 Td ({escaped}) Tj ET\n".encode("latin-1")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Length " + str(len(stream)).encode("ascii") + b" >>\nstream\n" + stream + b"endstream",
    ]
    output = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for index, obj in enumerate(objects, start=1):
        offsets.append(len(output))
        output.extend(f"{index} 0 obj\n".encode("ascii"))
        output.extend(obj)
        output.extend(b"\nendobj\n")
    xref = len(output)
    output.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    output.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        output.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    output.extend(
        (
            f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\n"
            f"startxref\n{xref}\n%%EOF\n"
        ).encode("ascii")
    )
    return bytes(output)


def make_docx(text: str):
    buffer = io.BytesIO()
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body><w:p><w:r><w:t>{text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        '</Relationships>'
    )
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", rels)
        archive.writestr("word/document.xml", document_xml)
    return buffer.getvalue()


def body_preview(result, limit=350):
    return result["body"].decode("utf-8", errors="replace")[:limit]


def test_00_candidate_embedding_and_rag_routes():
    if sys.version_info[:2] != (3, 11):
        return
    paths = [
        "/embeddings",
        "/embedding",
        "/embed",
        "/llm/embeddings",
        "/llm/embedding",
        "/llm/embed",
        "/llm/embed-many",
        "/llm/embedMany",
        "/text/embeddings",
        "/ai/embeddings",
        "/vector/embeddings",
        "/rag",
        "/rag/query",
        "/knowledge",
        "/knowledge/search",
        "/documents",
        "/documents/upload",
        "/files",
        "/files/upload",
        "/upload",
    ]
    probe_payload = {
        "input": ["RORK_EMBEDDING_PROBE"],
        "text": "RORK_EMBEDDING_PROBE",
        "texts": ["RORK_EMBEDDING_PROBE"],
    }
    results = []
    for path in paths:
        post = request(BASE_URL + path, probe_payload, method="POST", timeout=30)
        options = request(BASE_URL + path, None, method="OPTIONS", timeout=30)
        results.append(
            {
                "path": path,
                "post_status": post["status"],
                "post_type": post["headers"].get("content-type"),
                "post_body": body_preview(post, 180),
                "options_status": options["status"],
                "allow": options["headers"].get("allow"),
            }
        )
    log("candidate_routes", results=results)
    assert all(item["post_status"] >= 0 for item in results)


def test_01_agent_reads_plain_text_file():
    if sys.version_info[:2] != (3, 9):
        return
    marker = "TXT_FILE_MARKER_7319"
    message = file_message(
        "text/plain",
        "facts.txt",
        f"The hidden answer is {marker}.\n",
        "Read the attached file and reply with only the hidden answer.",
    )
    result = agent_request([message])
    log(
        "agent_text_file",
        status=result["status"],
        latency_ms=result["latency_ms"],
        text=result["text"],
        frame_types=[frame.get("type") for frame in result["frames"]],
        body=body_preview(result),
    )
    assert result["status"] == 200
    assert marker in result["text"]


def test_02_agent_reads_pdf_file():
    if sys.version_info[:2] != (3, 10):
        return
    marker = "PDF_FILE_MARKER_8426"
    message = file_message(
        "application/pdf",
        "facts.pdf",
        make_pdf(f"The hidden answer is {marker}"),
        "Read the attached PDF and reply with only the hidden answer.",
    )
    result = agent_request([message])
    log(
        "agent_pdf_file",
        status=result["status"],
        latency_ms=result["latency_ms"],
        text=result["text"],
        frame_types=[frame.get("type") for frame in result["frames"]],
        body=body_preview(result),
    )
    assert result["status"] == 200
    assert marker in result["text"]


def test_03_agent_reads_docx_file():
    if sys.version_info[:2] != (3, 11):
        return
    marker = "DOCX_FILE_MARKER_9537"
    message = file_message(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "facts.docx",
        make_docx(f"The hidden answer is {marker}"),
        "Read the attached Word document and reply with only the hidden answer.",
    )
    result = agent_request([message])
    log(
        "agent_docx_file",
        status=result["status"],
        latency_ms=result["latency_ms"],
        text=result["text"],
        frame_types=[frame.get("type") for frame in result["frames"]],
        body=body_preview(result),
    )
    assert result["status"] == 200
    assert marker in result["text"]


def test_04_rag_via_search_tool_roundtrip():
    if sys.version_info[:2] != (3, 11):
        return
    tool_name = "search_documents"
    user = {
        "id": "u-" + uuid.uuid4().hex[:12],
        "role": "user",
        "parts": [
            {
                "type": "text",
                "text": "Search the company documents before answering: What is the Acme warranty period?",
            }
        ],
    }
    tools = {
        tool_name: {
            "description": "Search the private company document index for relevant passages.",
            "jsonSchema": {
                "type": "object",
                "properties": {"query": {"type": "string"}},
                "required": ["query"],
                "additionalProperties": False,
            },
        }
    }
    first = agent_request([user], tools)
    call = next(
        (frame for frame in first["frames"] if frame.get("type") == "tool-input-available"),
        None,
    )
    assert first["status"] == 200
    assert call is not None
    assert call.get("toolName") == tool_name
    assistant = {
        "id": "a-" + uuid.uuid4().hex[:12],
        "role": "assistant",
        "parts": [
            {
                "type": f"tool-{tool_name}",
                "toolCallId": call["toolCallId"],
                "state": "output-available",
                "input": call["input"],
                "output": {
                    "matches": [
                        {
                            "source": "warranty-policy.txt",
                            "score": 0.992,
                            "text": "Acme products have a warranty period of exactly 37 months from purchase date.",
                        }
                    ]
                },
            }
        ],
    }
    second = agent_request([user, assistant], tools)
    log(
        "rag_tool_roundtrip",
        first_status=first["status"],
        tool_call=call,
        second_status=second["status"],
        text=second["text"],
        frame_types=[frame.get("type") for frame in second["frames"]],
    )
    assert second["status"] == 200
    assert "37" in second["text"]


def test_05_llm_text_rejects_or_accepts_file_content_part():
    if sys.version_info[:2] != (3, 11):
        return
    marker = "LLM_FILE_MARKER_1648"
    payload = {
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "file",
                        "file": data_url("text/plain", f"Answer: {marker}".encode("utf-8")),
                        "mediaType": "text/plain",
                        "filename": "facts.txt",
                    },
                    {"type": "text", "text": "Reply with only the answer inside the file."},
                ],
            }
        ]
    }
    result = request(f"{BASE_URL}/llm/text", payload)
    log(
        "llm_text_file_part",
        status=result["status"],
        latency_ms=result["latency_ms"],
        body=body_preview(result, 500),
    )
    assert result["status"] in (200, 400, 422, 500)

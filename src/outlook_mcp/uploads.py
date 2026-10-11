"""HTTP upload storage and attachment resolution helpers."""

from __future__ import annotations

import base64
import hashlib
import html
import os
import re
import time
from contextvars import ContextVar, Token
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path, PureWindowsPath
from uuid import UUID, uuid4

from .tools.common import ToolValidationError

UPLOAD_RELATIVE_PREFIX = "uploads"
DEFAULT_UPLOAD_MAX_BYTES = 3 * 1024 * 1024
DEFAULT_UPLOAD_TTL_SECONDS = 24 * 60 * 60
DEFAULT_UPLOAD_QUOTA_BYTES = 30 * 1024 * 1024
DEFAULT_UPLOAD_QUOTA_FILES = 100
DEFAULT_INLINE_IMAGE_MAX_WIDTH_PX = 600
ALLOWED_IMAGE_MIME_TYPES = {"image/png", "image/jpeg", "image/gif", "image/webp"}

_IMG_TAG_RE = re.compile(r"<img\b(?P<attrs>[^>]*)>", re.IGNORECASE)
_HTML_ATTR_RE = re.compile(
    r"""(?P<name>[A-Za-z_:][-A-Za-z0-9_:.]*)(?:\s*=\s*(?P<value>"[^"]*"|'[^']*'|[^\s"'=<>`]+))?"""
)
_STYLE_SIZE_RE = re.compile(r"(?:^|;)\s*(?:max-)?(?:width|height)\s*:", re.IGNORECASE)
_current_credential_id: ContextVar[int | None] = ContextVar("outlook_mcp_upload_credential_id", default=None)


@dataclass(frozen=True)
class UploadedFile:
    relative_path: str
    path: Path
    name: str
    mime_type: str
    size: int
    sha256: str
    width: int | None = None
    height: int | None = None
    content: bytes | None = None


@dataclass(frozen=True)
class ResolvedAttachment:
    name: str
    content: bytes
    content_base64: str
    content_type: str
    disposition: str
    content_id: str | None = None
    width: int | None = None
    height: int | None = None


def set_upload_credential_id(credential_id: int | None) -> Token[int | None]:
    return _current_credential_id.set(credential_id)


def reset_upload_credential_id(token: Token[int | None]) -> None:
    _current_credential_id.reset(token)


def current_upload_credential_id() -> int | None:
    return _current_credential_id.get()


def upload_root() -> Path:
    configured = os.environ.get("OUTLOOK_MCP_UPLOAD_DIR")
    return Path(configured).expanduser() if configured else Path.home() / ".outlook-mcp" / "uploads"


def upload_max_bytes() -> int:
    return _positive_int_env("OUTLOOK_MCP_UPLOAD_MAX_BYTES", DEFAULT_UPLOAD_MAX_BYTES)


def upload_ttl_seconds() -> int:
    return _positive_int_env("OUTLOOK_MCP_UPLOAD_TTL_SECONDS", DEFAULT_UPLOAD_TTL_SECONDS)


def upload_quota_bytes() -> int:
    return _positive_int_env("OUTLOOK_MCP_UPLOAD_QUOTA_BYTES", DEFAULT_UPLOAD_QUOTA_BYTES)


def upload_quota_files() -> int:
    return _positive_int_env("OUTLOOK_MCP_UPLOAD_QUOTA_FILES", DEFAULT_UPLOAD_QUOTA_FILES)


def cleanup_expired_uploads(*, now: float | None = None) -> int:
    """Delete expired upload files and the empty directories they leave behind."""
    root = upload_root()
    if not root.exists():
        return 0
    root_resolved = root.resolve(strict=True)
    cutoff = (time.time() if now is None else now) - upload_ttl_seconds()
    removed = 0
    for path in root.rglob("*"):
        try:
            stat = path.lstat()
        except FileNotFoundError:
            continue
        if path.is_dir() and not path.is_symlink():
            continue
        resolved_parent = path.parent.resolve(strict=True)
        _ensure_inside(resolved_parent, root_resolved)
        if stat.st_mtime < cutoff:
            try:
                path.unlink()
            except FileNotFoundError:
                continue
            removed += 1

    directories = sorted(
        (path for path in root.rglob("*") if path.is_dir() and not path.is_symlink()),
        key=lambda path: len(path.parts),
        reverse=True,
    )
    for directory in directories:
        try:
            directory.rmdir()
        except FileNotFoundError:
            continue
        except OSError as exc:
            if any(directory.iterdir()):
                continue
            raise exc
    return removed


def save_uploaded_file(
    *,
    credential_id: int,
    filename: str,
    mime_type: str,
    content: bytes,
    now: datetime | None = None,
) -> UploadedFile:
    if not isinstance(credential_id, int) or credential_id <= 0:
        raise ToolValidationError("credential_id is required for uploads")
    safe_name = sanitize_filename(filename)
    detected_mime, width, height = _validated_image_info(content, mime_type)
    date_text = (now or datetime.now(timezone.utc)).strftime("%Y%m%d")
    upload_id = uuid4().hex
    root = upload_root()
    root.mkdir(parents=True, exist_ok=True)
    try:
        os.chmod(root, 0o700)
    except OSError:
        pass
    _check_upload_quota(root / str(credential_id), incoming_bytes=len(content))
    root_resolved = root.resolve(strict=True)
    directory = root / str(credential_id) / date_text / upload_id
    try:
        directory.mkdir(parents=True, exist_ok=False)
    except FileExistsError:
        raise ToolValidationError("upload directory already exists — try again")
    directory_resolved = directory.resolve(strict=True)
    _ensure_inside(directory_resolved, root_resolved)

    path = directory / safe_name
    path.write_bytes(content)
    resolved = path.resolve(strict=True)
    _ensure_inside(resolved, root_resolved)

    relative_path = f"{UPLOAD_RELATIVE_PREFIX}/{credential_id}/{date_text}/{upload_id}/{safe_name}"
    return UploadedFile(
        relative_path=relative_path,
        path=resolved,
        name=safe_name,
        mime_type=detected_mime,
        size=len(content),
        sha256=hashlib.sha256(content).hexdigest(),
        width=width,
        height=height,
    )


def resolve_uploaded_file(relative_path: str, *, credential_id: int | None = None) -> UploadedFile:
    effective_credential_id = credential_id if credential_id is not None else current_upload_credential_id()
    if effective_credential_id is None:
        raise ToolValidationError("relative_path attachments require an HTTP credential context")
    if not isinstance(relative_path, str) or not relative_path.strip():
        raise ToolValidationError("relative_path is required")

    parts = _validate_relative_path(relative_path.strip(), effective_credential_id)
    root = upload_root()
    try:
        root_resolved = root.resolve(strict=True)
    except FileNotFoundError as exc:
        raise ToolValidationError("uploaded file was not found") from exc
    path = root.joinpath(*parts[1:])
    try:
        resolved = path.resolve(strict=True)
    except FileNotFoundError as exc:
        raise ToolValidationError("uploaded file was not found") from exc
    _ensure_inside(resolved, root_resolved)
    if not resolved.is_file():
        raise ToolValidationError("uploaded path is not a file")

    stat = resolved.stat()
    if stat.st_size <= 0:
        raise ToolValidationError("uploaded file is empty")
    if stat.st_size > upload_max_bytes():
        raise ToolValidationError(f"uploaded file must be less than or equal to {upload_max_bytes()} bytes")
    if time.time() - stat.st_mtime > upload_ttl_seconds():
        _remove_upload_and_empty_parents(resolved, root_resolved)
        raise ToolValidationError("uploaded file has expired")

    content = resolved.read_bytes()
    detected_mime, width, height = _validated_image_info(content, None)
    return UploadedFile(
        relative_path="/".join(parts),
        path=resolved,
        name=parts[-1],
        mime_type=detected_mime,
        size=len(content),
        sha256=hashlib.sha256(content).hexdigest(),
        width=width,
        height=height,
        content=content,
    )


def attachment_from_upload(raw: dict[str, object], *, html: bool) -> ResolvedAttachment:
    relative_path = str(raw.get("relative_path") or "").strip()
    disposition = _attachment_disposition(raw.get("disposition"))
    content_id = _content_id(raw.get("content_id"))
    if disposition == "inline":
        if not html:
            raise ToolValidationError("inline attachments require html=true")
        if not content_id:
            raise ToolValidationError("inline attachments require content_id")
    elif content_id:
        raise ToolValidationError("content_id is only valid for inline attachments")
    uploaded = resolve_uploaded_file(relative_path)
    content = uploaded.content
    assert content is not None
    return ResolvedAttachment(
        name=uploaded.name,
        content=content,
        content_base64=base64.b64encode(content).decode("ascii"),
        content_type=uploaded.mime_type,
        disposition=disposition,
        content_id=content_id,
        width=uploaded.width,
        height=uploaded.height,
    )


def normalize_inline_image_html(
    body: str,
    attachments: list[dict[str, object]] | None,
    max_inline_width_px: int = DEFAULT_INLINE_IMAGE_MAX_WIDTH_PX,
) -> str:
    if not body or not attachments:
        return body
    if max_inline_width_px < 1:
        raise ToolValidationError("max_inline_width_px must be greater than 0")
    inline_images = _inline_image_dimensions(attachments)
    if not inline_images:
        return body

    def replace_img(match: re.Match[str]) -> str:
        tag = match.group(0)
        attrs = _html_attributes(match.group("attrs"))
        src = html.unescape(attrs.get("src", "")).strip()
        if not src.lower().startswith("cid:"):
            return tag
        dimensions = inline_images.get(src[4:])
        if dimensions is None or _img_has_size_constraint(attrs):
            return tag
        width, height = dimensions
        if width <= max_inline_width_px:
            additions = f' width="{width}" height="{height}"'
        else:
            additions = f' width="{max_inline_width_px}" style="max-width:{max_inline_width_px}px;height:auto;"'
        prefix = tag[:-1].rstrip()
        self_closing = prefix.endswith("/")
        if self_closing:
            prefix = prefix[:-1].rstrip()
        return prefix + additions + (" />" if self_closing else ">")

    return _IMG_TAG_RE.sub(replace_img, body)


def sanitize_filename(filename: str) -> str:
    basename = Path(str(filename or "").replace("\\", "/")).name.strip()
    cleaned = "".join(char if char.isalnum() or char in "._-" else "_" for char in basename)
    cleaned = cleaned.strip("._")
    if not cleaned:
        raise ToolValidationError("uploaded filename is required")
    return cleaned[:180]


def validate_image_bytes(content: bytes, declared_mime_type: str | None) -> str:
    return _validated_image_info(content, declared_mime_type)[0]


def detect_image_dimensions(content: bytes, mime_type: str) -> tuple[int, int] | None:
    if mime_type == "image/png":
        if len(content) >= 24 and content.startswith(b"\x89PNG\r\n\x1a\n") and content[12:16] == b"IHDR":
            return int.from_bytes(content[16:20], "big"), int.from_bytes(content[20:24], "big")
        return None
    if mime_type == "image/jpeg":
        return _jpeg_dimensions(content)
    if mime_type == "image/gif":
        if len(content) >= 10 and content.startswith((b"GIF87a", b"GIF89a")):
            return int.from_bytes(content[6:8], "little"), int.from_bytes(content[8:10], "little")
        return None
    if mime_type == "image/webp":
        return _webp_dimensions(content)
    return None


def _validated_image_info(content: bytes, declared_mime_type: str | None) -> tuple[str, int, int]:
    if not content:
        raise ToolValidationError("uploaded file is empty")
    if len(content) > upload_max_bytes():
        raise ToolValidationError(f"uploaded file must be less than or equal to {upload_max_bytes()} bytes")
    declared = (declared_mime_type or "").split(";", 1)[0].strip().lower()
    if declared and declared not in ALLOWED_IMAGE_MIME_TYPES:
        raise ToolValidationError("uploaded file must be a PNG, JPEG, GIF or WebP image")
    detected = detect_image_mime_type(content)
    if detected is None:
        raise ToolValidationError("uploaded file content is not a supported image")
    if declared and declared != detected:
        raise ToolValidationError("uploaded file MIME type does not match its content")
    dimensions = detect_image_dimensions(content, detected)
    if dimensions is None:
        raise ToolValidationError("uploaded image dimensions could not be determined")
    width, height = dimensions
    if width < 1 or height < 1:
        raise ToolValidationError("uploaded image dimensions must be greater than 0")
    return detected, width, height


def detect_image_mime_type(content: bytes) -> str | None:
    if content.startswith(b"\x89PNG\r\n\x1a\n"):
        return "image/png"
    if content.startswith(b"\xff\xd8\xff"):
        return "image/jpeg"
    if content.startswith((b"GIF87a", b"GIF89a")):
        return "image/gif"
    if len(content) >= 12 and content[:4] == b"RIFF" and content[8:12] == b"WEBP":
        return "image/webp"
    return None


def _jpeg_dimensions(content: bytes) -> tuple[int, int] | None:
    if not content.startswith(b"\xff\xd8"):
        return None
    offset = 2
    while offset + 3 < len(content):
        if content[offset] != 0xFF:
            offset += 1
            continue
        while offset < len(content) and content[offset] == 0xFF:
            offset += 1
        if offset >= len(content):
            return None
        marker = content[offset]
        offset += 1
        if marker in {0x01, *range(0xD0, 0xD8), 0xD8, 0xD9}:
            continue
        if offset + 2 > len(content):
            return None
        segment_length = int.from_bytes(content[offset : offset + 2], "big")
        if segment_length < 2 or offset + segment_length > len(content):
            return None
        if marker in {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}:
            if segment_length < 7:
                return None
            height = int.from_bytes(content[offset + 3 : offset + 5], "big")
            width = int.from_bytes(content[offset + 5 : offset + 7], "big")
            return width, height
        offset += segment_length
    return None


def _webp_dimensions(content: bytes) -> tuple[int, int] | None:
    if len(content) < 20 or content[:4] != b"RIFF" or content[8:12] != b"WEBP":
        return None
    offset = 12
    while offset + 8 <= len(content):
        chunk_type = content[offset : offset + 4]
        chunk_size = int.from_bytes(content[offset + 4 : offset + 8], "little")
        data_offset = offset + 8
        if data_offset + chunk_size > len(content):
            return None
        data = content[data_offset : data_offset + chunk_size]
        if chunk_type == b"VP8X" and len(data) >= 10:
            width = int.from_bytes(data[4:7], "little") + 1
            height = int.from_bytes(data[7:10], "little") + 1
            return width, height
        if chunk_type == b"VP8 " and len(data) >= 10 and data[3:6] == b"\x9d\x01\x2a":
            width = int.from_bytes(data[6:8], "little") & 0x3FFF
            height = int.from_bytes(data[8:10], "little") & 0x3FFF
            return width, height
        if chunk_type == b"VP8L" and len(data) >= 5 and data[0] == 0x2F:
            b0, b1, b2, b3 = data[1], data[2], data[3], data[4]
            width = 1 + (((b1 & 0x3F) << 8) | b0)
            height = 1 + (((b3 & 0x0F) << 10) | (b2 << 2) | ((b1 & 0xC0) >> 6))
            return width, height
        offset = data_offset + chunk_size + (chunk_size % 2)
    return None


def _inline_image_dimensions(attachments: list[dict[str, object]]) -> dict[str, tuple[int, int]]:
    dimensions: dict[str, tuple[int, int]] = {}
    for raw in attachments:
        if not isinstance(raw, dict) or not raw.get("relative_path"):
            continue
        if _attachment_disposition(raw.get("disposition")) != "inline":
            continue
        content_id = _content_id(raw.get("content_id"))
        if not content_id:
            continue
        uploaded = resolve_uploaded_file(str(raw.get("relative_path") or ""))
        if uploaded.width is not None and uploaded.height is not None:
            dimensions[content_id] = (uploaded.width, uploaded.height)
    return dimensions


def _html_attributes(attrs_text: str) -> dict[str, str]:
    attrs: dict[str, str] = {}
    for match in _HTML_ATTR_RE.finditer(attrs_text):
        name = match.group("name").lower()
        raw_value = match.group("value")
        if raw_value is None:
            attrs[name] = ""
        elif raw_value[:1] in {"'", '"'} and raw_value[-1:] == raw_value[:1]:
            attrs[name] = raw_value[1:-1]
        else:
            attrs[name] = raw_value
    return attrs


def _img_has_size_constraint(attrs: dict[str, str]) -> bool:
    if "width" in attrs or "height" in attrs:
        return True
    style = attrs.get("style")
    return bool(style and _STYLE_SIZE_RE.search(style))


def _validate_relative_path(relative_path: str, credential_id: int) -> list[str]:
    if "\\" in relative_path or Path(relative_path).is_absolute() or PureWindowsPath(relative_path).drive:
        raise ToolValidationError("relative_path must be a safe upload path")
    parts = relative_path.split("/")
    if len(parts) != 5 or any(part in {"", ".", ".."} for part in parts):
        raise ToolValidationError("relative_path must use uploads/{credential_id}/{YYYYMMDD}/{uuid}/{filename}")
    prefix, owner_id, date_text, upload_id, filename = parts
    if prefix != UPLOAD_RELATIVE_PREFIX:
        raise ToolValidationError("relative_path must start with uploads/")
    if owner_id != str(credential_id):
        raise ToolValidationError("relative_path does not belong to this endpoint")
    if not (date_text.isdigit() and len(date_text) == 8):
        raise ToolValidationError("relative_path contains an invalid upload date")
    try:
        UUID(upload_id)
    except ValueError as exc:
        raise ToolValidationError("relative_path contains an invalid upload id") from exc
    if sanitize_filename(filename) != filename:
        raise ToolValidationError("relative_path contains an unsafe filename")
    return parts


def _attachment_disposition(value: object) -> str:
    disposition = str(value or "attachment").strip().lower()
    if disposition not in {"attachment", "inline"}:
        raise ToolValidationError("attachment disposition must be attachment or inline")
    return disposition


def _content_id(value: object) -> str | None:
    if value is None:
        return None
    content_id = str(value).strip()
    if not content_id:
        return None
    if any(char in content_id for char in "\r\n<>"):
        raise ToolValidationError("content_id contains invalid characters")
    return content_id


def _ensure_inside(path: Path, root: Path) -> None:
    if path != root and root not in path.parents:
        raise ToolValidationError("upload path escaped the configured upload directory")


def _remove_upload_and_empty_parents(path: Path, root: Path) -> None:
    path.unlink(missing_ok=True)
    parent = path.parent
    while parent != root:
        try:
            parent.rmdir()
        except FileNotFoundError:
            pass
        except OSError:
            break
        parent = parent.parent


def _check_upload_quota(credential_root: Path, *, incoming_bytes: int) -> None:
    if not credential_root.exists():
        return
    cutoff = time.time() - upload_ttl_seconds()
    total_bytes = 0
    file_count = 0
    for path in credential_root.rglob("*"):
        try:
            stat = path.lstat()
        except FileNotFoundError:
            continue
        if path.is_dir() or path.is_symlink():
            continue
        if stat.st_mtime < cutoff:
            continue
        file_count += 1
        total_bytes += stat.st_size
    if file_count + 1 > upload_quota_files():
        raise ToolValidationError(
            f"upload quota exceeded: {file_count} files already stored, "
            f"limit is {upload_quota_files()} files per credential"
        )
    if total_bytes + incoming_bytes > upload_quota_bytes():
        raise ToolValidationError(
            f"upload quota exceeded: {total_bytes + incoming_bytes} bytes total would exceed "
            f"the {upload_quota_bytes()} bytes per-credential limit"
        )


def _positive_int_env(name: str, default: int) -> int:
    raw = os.environ.get(name)
    if raw is None:
        return default
    try:
        value = int(raw)
    except ValueError as exc:
        raise ToolValidationError(f"{name} must be an integer") from exc
    if value < 1:
        raise ToolValidationError(f"{name} must be greater than 0")
    return value

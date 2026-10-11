from __future__ import annotations

import os
import time

import pytest

from conftest import PNG_BYTES
from outlook_mcp.tools.common import ToolValidationError
from outlook_mcp.uploads import (
    DEFAULT_UPLOAD_TTL_SECONDS,
    cleanup_expired_uploads,
    detect_image_dimensions,
    normalize_inline_image_html,
    resolve_uploaded_file,
    reset_upload_credential_id,
    save_uploaded_file,
    set_upload_credential_id,
)


JPEG_BYTES = b"\xff\xd8\xff\xe0\x00\x10JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00" + b"\xff\xc0\x00\x11\x08" + (25).to_bytes(2, "big") + (40).to_bytes(2, "big") + b"\x03\x01\x11\x00\x02\x11\x00\x03\x11\x00\xff\xd9"
GIF_BYTES = b"GIF89a" + (12).to_bytes(2, "little") + (34).to_bytes(2, "little") + b"\x00\x00\x00"
WEBP_BYTES = b"RIFF" + (22).to_bytes(4, "little") + b"WEBPVP8X" + (10).to_bytes(4, "little") + b"\x00\x00\x00\x00" + (63).to_bytes(3, "little") + (47).to_bytes(3, "little")


def test_save_and_resolve_uploaded_image(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))

    uploaded = save_uploaded_file(
        credential_id=12,
        filename="../unsafe photo.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )
    resolved = resolve_uploaded_file(uploaded.relative_path, credential_id=12)

    assert uploaded.relative_path.startswith("uploads/12/")
    assert ".." not in uploaded.relative_path
    assert str(tmp_path) not in uploaded.relative_path
    assert resolved.path.read_bytes() == PNG_BYTES
    assert resolved.mime_type == "image/png"
    assert resolved.size == len(PNG_BYTES)
    assert resolved.width == 320
    assert resolved.height == 180


@pytest.mark.parametrize(
    ("content", "mime_type", "dimensions"),
    [
        (PNG_BYTES, "image/png", (320, 180)),
        (JPEG_BYTES, "image/jpeg", (40, 25)),
        (GIF_BYTES, "image/gif", (12, 34)),
        (WEBP_BYTES, "image/webp", (64, 48)),
    ],
)
def test_image_dimensions_are_parsed(content: bytes, mime_type: str, dimensions: tuple[int, int]) -> None:
    assert detect_image_dimensions(content, mime_type) == dimensions


def test_normalize_inline_image_html_adds_default_dimensions(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    uploaded = save_uploaded_file(
        credential_id=12,
        filename="photo.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )

    token = set_upload_credential_id(12)
    try:
        html = normalize_inline_image_html(
            '<p><img alt="photo" src="cid:photo1"></p>',
            [{"relative_path": uploaded.relative_path, "disposition": "inline", "content_id": "photo1"}],
        )
    finally:
        reset_upload_credential_id(token)

    assert html == '<p><img alt="photo" src="cid:photo1" width="320" height="180"></p>'


def test_normalize_inline_image_html_shrinks_large_images(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    large_png = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR" + (1200).to_bytes(4, "big") + (900).to_bytes(4, "big") + b"\x08\x02\x00\x00\x00"
    uploaded = save_uploaded_file(
        credential_id=12,
        filename="wide.png",
        mime_type="image/png",
        content=large_png,
    )

    token = set_upload_credential_id(12)
    try:
        html = normalize_inline_image_html(
            '<img src="cid:wide-photo">',
            [{"relative_path": uploaded.relative_path, "disposition": "inline", "content_id": "wide-photo"}],
        )
    finally:
        reset_upload_credential_id(token)

    assert html == '<img src="cid:wide-photo" width="600" style="max-width:600px;height:auto;">'


@pytest.mark.parametrize(
    "body",
    [
        '<img src="cid:photo1" width="12">',
        '<img src="cid:photo1" height="12">',
        '<img src="cid:photo1" style="max-width:100px;">',
        '<img src="cid:missing">',
    ],
)
def test_normalize_inline_image_html_keeps_existing_constraints(monkeypatch, tmp_path, body: str) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    uploaded = save_uploaded_file(
        credential_id=12,
        filename="photo.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )

    token = set_upload_credential_id(12)
    try:
        html = normalize_inline_image_html(
            body,
            [{"relative_path": uploaded.relative_path, "disposition": "inline", "content_id": "photo1"}],
        )
    finally:
        reset_upload_credential_id(token)

    assert html == body


@pytest.mark.parametrize(
    "relative_path",
    [
        "../photo.png",
        "/uploads/12/20260522/abc/photo.png",
        "C:/uploads/12/20260522/abc/photo.png",
        "uploads/12/20260522/abc/../photo.png",
    ],
)
def test_resolve_uploaded_file_rejects_unsafe_paths(monkeypatch, tmp_path, relative_path: str) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))

    with pytest.raises(ToolValidationError):
        resolve_uploaded_file(relative_path, credential_id=12)


def test_resolve_uploaded_file_rejects_cross_credential_and_expired(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    uploaded = save_uploaded_file(
        credential_id=12,
        filename="photo.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )

    with pytest.raises(ToolValidationError, match="does not belong"):
        resolve_uploaded_file(uploaded.relative_path, credential_id=13)

    expired_at = time.time() - DEFAULT_UPLOAD_TTL_SECONDS - 1
    os.utime(uploaded.path, (expired_at, expired_at))
    with pytest.raises(ToolValidationError, match="expired"):
        resolve_uploaded_file(uploaded.relative_path, credential_id=12)
    assert not uploaded.path.exists()


def test_cleanup_expired_uploads_removes_files_and_empty_directories(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    expired = save_uploaded_file(
        credential_id=12,
        filename="expired.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )
    current = save_uploaded_file(
        credential_id=12,
        filename="current.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )
    expired_at = time.time() - DEFAULT_UPLOAD_TTL_SECONDS - 1
    os.utime(expired.path, (expired_at, expired_at))

    assert cleanup_expired_uploads() == 1
    assert not expired.path.exists()
    assert not expired.path.parent.exists()
    assert current.path.exists()


def test_save_uploaded_file_rejects_bad_mime_empty_and_oversize(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_MAX_BYTES", "8")

    with pytest.raises(ToolValidationError, match="empty"):
        save_uploaded_file(credential_id=12, filename="empty.png", mime_type="image/png", content=b"")
    with pytest.raises(ToolValidationError, match="PNG, JPEG, GIF or WebP"):
        save_uploaded_file(credential_id=12, filename="note.txt", mime_type="text/plain", content=PNG_BYTES[:8])
    with pytest.raises(ToolValidationError, match="less than or equal"):
        save_uploaded_file(credential_id=12, filename="photo.png", mime_type="image/png", content=PNG_BYTES)


# --- sanitize_filename Unicode handling ---


def test_sanitize_filename_keeps_unicode_letters() -> None:
    from outlook_mcp.uploads import sanitize_filename

    assert sanitize_filename("月度报告.pdf") == "月度报告.pdf"
    assert sanitize_filename("..\\月度报告.pdf") == "月度报告.pdf"
    assert sanitize_filename("résumé 2026.txt") == "résumé_2026.txt"
    assert sanitize_filename(".hidden") == "hidden"
    with pytest.raises(ToolValidationError):
        sanitize_filename("..")
    with pytest.raises(ToolValidationError):
        sanitize_filename("")


def test_unicode_filename_round_trips_through_upload(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))

    uploaded = save_uploaded_file(
        credential_id=12,
        filename="月度报告.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )
    resolved = resolve_uploaded_file(uploaded.relative_path, credential_id=12)

    assert uploaded.name == "月度报告.png"
    assert resolved.content == PNG_BYTES


# --- per-credential upload quota ---


def test_save_uploaded_file_enforces_file_count_quota(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_QUOTA_FILES", "1")

    save_uploaded_file(credential_id=12, filename="a.png", mime_type="image/png", content=PNG_BYTES)
    with pytest.raises(ToolValidationError, match="files per credential"):
        save_uploaded_file(credential_id=12, filename="b.png", mime_type="image/png", content=PNG_BYTES)

    save_uploaded_file(credential_id=13, filename="c.png", mime_type="image/png", content=PNG_BYTES)


def test_save_uploaded_file_enforces_byte_quota(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_QUOTA_BYTES", str(len(PNG_BYTES) + 1))

    save_uploaded_file(credential_id=12, filename="a.png", mime_type="image/png", content=PNG_BYTES)
    with pytest.raises(ToolValidationError, match="bytes per-credential limit"):
        save_uploaded_file(credential_id=12, filename="b.png", mime_type="image/png", content=PNG_BYTES)


def test_save_uploaded_file_quota_ignores_expired_files(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_QUOTA_FILES", "1")
    expired = save_uploaded_file(
        credential_id=12, filename="old.png", mime_type="image/png", content=PNG_BYTES
    )
    expired_at = time.time() - DEFAULT_UPLOAD_TTL_SECONDS - 1
    os.utime(expired.path, (expired_at, expired_at))

    save_uploaded_file(credential_id=12, filename="new.png", mime_type="image/png", content=PNG_BYTES)


@pytest.mark.skipif(os.name == "nt", reason="POSIX permissions")
def test_upload_root_directory_permissions(monkeypatch, tmp_path) -> None:
    import stat

    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    save_uploaded_file(credential_id=7, filename="a.png", mime_type="image/png", content=PNG_BYTES)
    assert stat.S_IMODE(os.stat(tmp_path / "uploads").st_mode) == 0o700

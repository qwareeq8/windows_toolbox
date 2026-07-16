"""Security boundary tests for the embedded web frontend."""

import sys
from typing import cast

import pytest
from PySide6.QtCore import QUrl
from PySide6.QtWebEngineCore import QWebEnginePage, QWebEngineUrlRequestInfo

from virelo.app import webview


def test_frozen_executable_cannot_enable_development_mode(monkeypatch):
    """A packaged executable ignores the development-mode environment flag."""
    monkeypatch.setenv("VIRELO_DEV", "1")
    monkeypatch.delattr(sys, "frozen", raising=False)
    assert webview._is_dev_mode() is True

    monkeypatch.setattr(sys, "frozen", True, raising=False)
    assert webview._is_dev_mode() is False


def test_release_resources_use_an_explicit_allowlist(tmp_path, monkeypatch):
    """Release requests stay in the frontend bundle plus required Qt assets."""
    frontend = tmp_path / "frontend" / "dist"
    frontend.mkdir(parents=True)
    inside = frontend / "assets" / "app.js"
    inside.parent.mkdir()
    inside.write_text("Virelo", encoding="utf-8")
    outside = tmp_path / "private.txt"
    outside.write_text("Private", encoding="utf-8")
    monkeypatch.setattr(webview, "resource_path", lambda relative: str(tmp_path / relative))

    assert webview._is_allowed_release_resource(QUrl.fromLocalFile(str(inside))) is True
    assert webview._is_allowed_release_resource(QUrl.fromLocalFile(str(outside))) is False
    assert webview._is_allowed_release_resource(QUrl("qrc:///qtwebchannel/qwebchannel.js")) is True
    assert webview._is_allowed_release_resource(QUrl("qrc:///qtwebchannel/other.js")) is False
    assert webview._is_allowed_release_resource(QUrl("data:image/png;base64,AA==")) is True
    assert webview._is_allowed_release_resource(QUrl("https://example.com/app.js")) is False


class _NavigationPage:
    """Exercise VireloWebPage's pure navigation policy without a GUI process."""

    def __init__(self):
        self._trusted_data_url = None
        self.pending_url = None

    def setHtml(self, html):
        encoded = webview._SET_HTML_DATA_PREFIX + QUrl.toPercentEncoding(html).data()
        self.pending_url = QUrl.fromEncoded(encoded)


def test_internal_error_html_is_allowed_once(monkeypatch):
    """Only the exact one-shot data navigation created for error HTML is allowed."""
    monkeypatch.delenv("VIRELO_DEV", raising=False)
    fake_page = _NavigationPage()
    page = cast(webview.VireloWebPage, fake_page)
    html = "<!doctype html><html><body>Missing frontend.</body></html>"
    encoded = webview._SET_HTML_DATA_PREFIX + QUrl.toPercentEncoding(html).data()
    url = QUrl.fromEncoded(encoded)

    webview.VireloWebPage.set_trusted_error_html(page, html)

    assert fake_page._trusted_data_url == encoded
    assert (
        webview.VireloWebPage.acceptNavigationRequest(
            page,
            fake_page.pending_url,
            QWebEnginePage.NavigationType.NavigationTypeTyped,
            True,
        )
        is True
    )
    assert fake_page._trusted_data_url is None
    assert (
        webview.VireloWebPage.acceptNavigationRequest(
            page,
            url,
            QWebEnginePage.NavigationType.NavigationTypeTyped,
            True,
        )
        is False
    )


@pytest.mark.parametrize(
    ("navigation_type", "is_main_frame"),
    [
        (QWebEnginePage.NavigationType.NavigationTypeOther, True),
        (QWebEnginePage.NavigationType.NavigationTypeTyped, False),
    ],
)
def test_error_html_gate_rejects_wrong_navigation_context(
    navigation_type, is_main_frame, monkeypatch
):
    """The error-document exception applies only to a typed main-frame load."""
    monkeypatch.delenv("VIRELO_DEV", raising=False)
    fake_page = _NavigationPage()
    page = cast(webview.VireloWebPage, fake_page)
    html = "<p>Internal error.</p>"
    encoded = webview._SET_HTML_DATA_PREFIX + QUrl.toPercentEncoding(html).data()
    fake_page._trusted_data_url = encoded

    assert (
        webview.VireloWebPage.acceptNavigationRequest(
            page,
            QUrl.fromEncoded(encoded),
            navigation_type,
            is_main_frame,
        )
        is False
    )
    assert fake_page._trusted_data_url is None


class _RequestInfo:
    """Minimal request object for interceptor behavior tests."""

    def __init__(self, url):
        self._url = url
        self.blocked = False

    def requestUrl(self):
        return self._url

    def resourceType(self):
        return "script"

    def block(self, blocked):
        self.blocked = blocked


def test_release_interceptor_blocks_disallowed_requests(monkeypatch):
    """The release interceptor blocks network resources before Chromium loads them."""
    monkeypatch.setattr(
        webview,
        "_is_allowed_release_resource",
        lambda url: url.scheme() == "data",
    )
    allowed = _RequestInfo(QUrl("data:text/plain,Virelo"))
    blocked = _RequestInfo(QUrl("https://example.com/app.js"))
    interceptor = cast(webview.VireloRequestInterceptor, object())

    webview.VireloRequestInterceptor.interceptRequest(
        interceptor, cast(QWebEngineUrlRequestInfo, allowed)
    )
    webview.VireloRequestInterceptor.interceptRequest(
        interceptor, cast(QWebEngineUrlRequestInfo, blocked)
    )

    assert allowed.blocked is False
    assert blocked.blocked is True

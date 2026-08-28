from __future__ import annotations

import sys
import threading

from frontend_server import build_parser, create_app_server, run_worker


WINDOW_TITLE = "Teacher Feedback Desk"


def main() -> None:
    """Run the app in its own native window instead of a browser tab.

    Same server, same interface: this starts frontend_server's HTTP server on a
    background thread and shows it in a pywebview window (WKWebView on macOS,
    WebView2 on Windows). Closing the window shuts the server down.

    This is also the entry point a packaged build will use, so the --worker
    relaunch used for child processes has to be handled here too.
    """
    if len(sys.argv) > 1 and sys.argv[1] == "--worker":
        run_worker(sys.argv[2:])
        return

    # Imported lazily so frontend_server.py keeps working without pywebview installed.
    import webview

    parser = build_parser()
    # Port 0 lets the OS pick a free port, so a browser-mode server (or a stray old
    # instance) on 8765 can never block the window from opening. The window navigates
    # to whatever port was actually bound, so nothing else cares which one it is.
    parser.set_defaults(port=0, no_browser=True)
    args = parser.parse_args()

    server = create_app_server(args)
    host, port = server.server_address[:2]
    url = f"http://{host}:{port}"
    print(f"Feedback frontend (window): {url}")

    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()

    try:
        webview.create_window(WINDOW_TITLE, url, width=1440, height=900, min_size=(1000, 640))
        # Blocks until the window is closed. On macOS the GUI loop must own the main
        # thread, which is why the server runs on the background thread and not this one.
        webview.start()
    finally:
        server.shutdown()
        server.server_close()


if __name__ == "__main__":
    main()

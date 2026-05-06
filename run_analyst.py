"""
Launcher wrapper for the Energy Data Analyst Streamlit app.
-------------------------------------------------------------
Starts the Streamlit server, opens the browser, and monitors it.
When the browser tab/window is closed, the server is shut down automatically.
"""

import subprocess
import sys
import time
import socket
import webbrowser
import signal
import os


def find_free_port():
    """Find a free port on localhost."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("", 0))
        return s.getsockname()[1]


def wait_for_server(port, timeout=30):
    """Wait until the Streamlit server is accepting connections."""
    start = time.time()
    while time.time() - start < timeout:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=1):
                return True
        except OSError:
            time.sleep(0.3)
    return False


def main():
    port = find_free_port()
    app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "analyst.py")

    print("=" * 60)
    print("  Energy Data Analyst - Starting...")
    print("=" * 60)
    print(f"\n  Port : {port}")
    print(f"  URL  : http://localhost:{port}")
    print("\n  >>  Close THIS window to stop the application.")
    print("=" * 60)

    # Build the streamlit command
    cmd = [
        sys.executable, "-m", "streamlit", "run", app_path,
        "--server.port", str(port),
        "--server.headless", "true",
        "--browser.gatherUsageStats", "false",
        "--server.address", "127.0.0.1",
    ]

    # Start Streamlit as a subprocess
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0,
    )

    def shutdown(*_args):
        """Terminate the Streamlit process tree."""
        print("\n\n  Shutting down the server...")
        try:
            if os.name == "nt":
                # On Windows, kill the entire process tree
                subprocess.call(
                    ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
            else:
                proc.terminate()
                proc.wait(timeout=5)
        except Exception:
            proc.kill()
        print("  Server stopped. You can close this window.")

    # Register cleanup handlers
    signal.signal(signal.SIGINT, shutdown)
    signal.signal(signal.SIGTERM, shutdown)
    if os.name == "nt":
        signal.signal(signal.SIGBREAK, shutdown)

    # Wait for server to be ready, then open browser
    if wait_for_server(port):
        url = f"http://localhost:{port}"
        print(f"\n  [OK] Server ready - opening browser at {url}\n")
        webbrowser.open(url)
    else:
        print("\n  [FAIL] Server did not start within 30 seconds.")
        shutdown()
        sys.exit(1)

    # Keep running until the subprocess exits or the user closes the window
    try:
        # Stream Streamlit output so the user can see logs if needed
        for line in proc.stdout:
            # Only print important lines, skip noise
            stripped = line.strip()
            if stripped and not stripped.startswith("You can now view"):
                pass  # suppress most streamlit output to keep console clean
        proc.wait()
    except KeyboardInterrupt:
        shutdown()


if __name__ == "__main__":
    main()

import webview
import threading
import time
import requests
import os
import sys

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.abspath(".")

os.environ["FLASK_APP_BASE"] = base_path

from app import app

def start_flask():
    # Crucial: use_reloader must be False in an EXE
    app.run(port=5000, debug=False, use_reloader=False)

def wait_for_server(url, timeout=15):
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            # We use a head request to minimize data overhead
            r = requests.head(url)
            if r.status_code == 200:
                return True
        except:
            pass
        time.sleep(0.5)
    return False

if __name__ == "__main__":
    flask_thread = threading.Thread(target=start_flask)
    flask_thread.daemon = True
    flask_thread.start()

    if wait_for_server("http://127.0.0.1:5000"):
        # Explicitly using 'edgechromium' prevents many string/byte encoding errors on Windows
        webview.create_window("Paysly Tool", "http://127.0.0.1:5000", width=1100, height=700)
        webview.start(gui='edgechromium')








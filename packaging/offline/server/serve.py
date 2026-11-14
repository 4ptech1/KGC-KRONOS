# KRONOS offline: the same tiny server, for computers with Python instead of Node.js.
import functools, http.server, os, socketserver, sys, threading, webbrowser

PORT = 5181
ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'app')

class Handler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {**http.server.SimpleHTTPRequestHandler.extensions_map,
                      '.js': 'text/javascript', '.mjs': 'text/javascript', '.wasm': 'application/wasm',
                      '.woff2': 'font/woff2', '.woff': 'font/woff', '.svg': 'image/svg+xml'}
    def log_message(self, *args):
        pass
    def end_headers(self):
        self.send_header('Cache-Control', 'no-cache')
        super().end_headers()

# One thread per connection: the browser fetches Python's files side by side, and a server that
# answers one connection at a time stalls on them (found by KRONOS's tests, D86).
socketserver.ThreadingTCPServer.daemon_threads = True
# Mac and Linux keep a just-closed port reserved for a while, so allow re-using it. Not on Windows,
# where that would let two copies share the port without noticing.
socketserver.ThreadingTCPServer.allow_reuse_address = os.name != 'nt'
try:
    server = socketserver.ThreadingTCPServer(('127.0.0.1', PORT), functools.partial(Handler, directory=ROOT))
except OSError:
    print(f'\nPort {PORT} is busy. KRONOS may already be running: open http://localhost:{PORT}/ in your browser.\n')
    sys.exit(1)
url = f'http://localhost:{PORT}/'
print(f'\nKRONOS Offline, version dev\nKRONOS is running at {url}\nKeep this window open while you use it. Close it (or press Ctrl+C) to stop.\n')
threading.Timer(0.5, lambda: webbrowser.open(url)).start()
try:
    server.serve_forever()
except KeyboardInterrupt:
    pass

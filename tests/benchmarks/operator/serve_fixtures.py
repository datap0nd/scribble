from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
from urllib.parse import unquote, urlsplit
import argparse

def handler(root):
    root=Path(root).resolve(strict=True)
    class Handler(SimpleHTTPRequestHandler):
        def __init__(self,*a,**kw):super().__init__(*a,directory=str(root),**kw)
        def do_GET(self):
            part=unquote(urlsplit(self.path).path).lstrip('/') or 'index.html'
            target=(root/part).resolve()
            if target.parent!=root or target.name not in {'index.html','operations.html','archive.html'}:
                self.send_error(404);return
            super().do_GET()
        def do_HEAD(self):self.do_GET()
        def list_directory(self,path):self.send_error(404)
    return Handler

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--root',required=True);p.add_argument('--port',type=int,default=8765);a=p.parse_args()
    server=ThreadingHTTPServer(('127.0.0.1',a.port),handler(a.root))
    print(f'Fixture-only server http://127.0.0.1:{a.port}/index.html',flush=True)
    server.serve_forever()

import BaseHTTPServer, SimpleHTTPServer
import ssl
import os

web_dir = "dist"
os.chdir(web_dir)

httpd = BaseHTTPServer.HTTPServer(('localhost', 3000), SimpleHTTPServer.SimpleHTTPRequestHandler)
httpd.socket = ssl.wrap_socket (httpd.socket, certfile='../python-server/server.pem', server_side=True)
httpd.serve_forever()

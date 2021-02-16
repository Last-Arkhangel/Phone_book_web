#!/usr/bin/env python3
# -*- coding: utf-8 -*-


from functools import partial
from http.server import CGIHTTPRequestHandler
import base64
import os
import ssl
import http.server
from threading import Thread
from socketserver import ThreadingMixIn


class ThreadingHTTPServer(ThreadingMixIn, http.server.HTTPServer):
    pass


class AuthHTTPRequestHandler(CGIHTTPRequestHandler):
    """ Main class to present webpages and authentication. """

    def __init__(self, *args, **kwargs):
        username  = kwargs.pop("username")
        password  = kwargs.pop("password")
        self._auth = base64.b64encode(f"логин:пароль".encode()).decode()
        super().__init__(*args)

    def do_HEAD(self):
        self.send_response(200)
        self.send_header("Content-type", "text/html")
        self.end_headers()

    def do_AUTHHEAD(self):
        self.send_response(401)
        self.send_header("WWW-Authenticate", 'Basic realm="Phone book"')
        self.send_header("Content-type", "text/html")
        self.end_headers()

    def do_GET(self):
        """ Present frontpage with user authentication. """
        if self.headers.get("Authorization") == None:
            self.do_AUTHHEAD()
            self.wfile.write(b"no auth header received")
        elif self.headers.get("Authorization") == "Basic " + self._auth:
            CGIHTTPRequestHandler.do_GET(self)
        else:
            self.do_AUTHHEAD()
            self.wfile.write(self.headers.get("Authorization").encode())
            self.wfile.write(b"not authenticated")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--cgi", action="store_true", help="Run as CGI Server")
    parser.add_argument(
        "--bind",
        "-b",
        metavar="ADDRESS",
        default="192.168.1.1",
        help="Specify alternate bind address " "[default: localhost]",
    )
    parser.add_argument(
        "port",
        action="store",
        default=8008,
        type=int,
        nargs="?",
        help="Specify alternate port [default: 8008]",
    )
    parser.add_argument("--username", "-u", metavar="USERNAME")
    parser.add_argument("--password", "-p", metavar="PASSWORD")
    args = parser.parse_args()
    handler_class = partial(
        AuthHTTPRequestHandler,
        username=args.username,
        password=args.password,
    )
    httpd = ThreadingHTTPServer((args.bind, args.port), handler_class)
    httpd.socket = ssl.wrap_socket (httpd.socket, keyfile="ssl/server.key", certfile="ssl/server.crt", server_side=True)

    httpd.serve_forever()
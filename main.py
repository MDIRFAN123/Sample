"""Anaconda Enterprise entry point for the PDF QR Code Extractor."""

import argparse
import logging
import os
import shlex
import sys

logging.basicConfig(
    level=os.environ.get("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)

from flask import Response
from werkzeug.middleware.dispatcher import DispatcherMiddleware
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.serving import run_simple

from app import app


def runtime_parser():
    parser = argparse.ArgumentParser(description="PDF QR Code Extractor")
    parser.add_argument(
        "--anaconda-project-host",
        action="append",
        default=[],
        help="Hostname allowed by the Anaconda Enterprise deployment",
    )
    parser.add_argument(
        "--anaconda-project-port",
        type=int,
        default=int(os.environ.get("ANACONDA_PROJECT_PORT", "0")),
        help="Runtime-provided HTTP port",
    )
    parser.add_argument(
        "--anaconda-project-iframe-hosts",
        default=os.environ.get("ANACONDA_PROJECT_IFRAME_HOSTS", ""),
        help="Space-separated origins allowed to embed the application",
    )
    parser.add_argument(
        "--anaconda-project-no-browser",
        action="store_true",
        help="Accepted for Anaconda Enterprise command compatibility",
    )
    parser.add_argument(
        "--anaconda-project-use-xheaders",
        action="store_true",
        help="Trust reverse-proxy forwarding headers",
    )
    parser.add_argument(
        "--anaconda-project-url-prefix",
        default=os.environ.get("ANACONDA_PROJECT_URL_PREFIX", ""),
        help="Runtime URL prefix",
    )
    parser.add_argument(
        "--anaconda-project-address",
        default=os.environ.get("ANACONDA_PROJECT_ADDRESS", "0.0.0.0"),
        help="Runtime bind address",
    )
    return parser


def normalize_prefix(value):
    prefix = (value or "").strip().strip("/")
    return f"/{prefix}" if prefix else ""


def configure_application(args):
    app.config.update(
        PREFERRED_URL_SCHEME="https",
        PROJECT_HOSTS=list(args.anaconda_project_host),
        PROJECT_PORT=args.anaconda_project_port,
    )
    if args.anaconda_project_host:
        app.config["TRUSTED_HOSTS"] = list(args.anaconda_project_host)

    if args.anaconda_project_use_xheaders:
        app.wsgi_app = ProxyFix(
            app.wsgi_app,
            x_for=1,
            x_proto=1,
            x_host=1,
            x_port=1,
            x_prefix=1,
        )

    iframe_hosts = shlex.split(args.anaconda_project_iframe_hosts or "")

    @app.after_request
    def deployment_headers(response):
        frame_ancestors = ["'self'", *iframe_hosts]
        response.headers["Content-Security-Policy"] = (
            "frame-ancestors " + " ".join(frame_ancestors)
        )
        return response

    prefix = normalize_prefix(args.anaconda_project_url_prefix)
    if not prefix:
        return app

    not_found = Response("Not Found", status=404)
    return DispatcherMiddleware(not_found, {prefix: app})


def main(argv=None):
    args = runtime_parser().parse_args(argv)
    application = configure_application(args)
    run_simple(
        hostname=args.anaconda_project_address,
        port=args.anaconda_project_port,
        application=application,
        use_debugger=False,
        use_reloader=False,
        threaded=True,
    )


if __name__ == "__main__":
    main(sys.argv[1:])

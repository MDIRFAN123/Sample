"""Anaconda Enterprise entry point for the COPS application."""
import argparse
import os

from werkzeug.middleware.proxy_fix import ProxyFix

os.environ.setdefault("TEAM_APP", "COPS")
os.environ.setdefault("ENTERPRISE_MODE", "true")
os.environ.setdefault("AUTO_DOWNLOAD_ENABLED", "true")
os.environ.setdefault("AUTO_UPLOAD_ENABLED", "false")

from app import create_app


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--anaconda-project-host", action="append", default=[])
    parser.add_argument("--anaconda-project-port", type=int, default=8086)
    parser.add_argument("--anaconda-project-iframe-hosts", default="")
    parser.add_argument("--anaconda-project-no-browser", action="store_true")
    parser.add_argument("--anaconda-project-use-xheaders", action="store_true")
    parser.add_argument("--anaconda-project-url-prefix", default="")
    parser.add_argument("--anaconda-project-address", default="0.0.0.0")
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    application = create_app()
    if args.anaconda_project_use_xheaders:
        application.wsgi_app = ProxyFix(
            application.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1
        )
    prefix = "/" + args.anaconda_project_url_prefix.strip("/") if args.anaconda_project_url_prefix else ""
    if prefix:
        application.config["APPLICATION_ROOT"] = prefix
        original_app = application.wsgi_app

        def prefixed_app(environ, start_response):
            path = environ.get("PATH_INFO", "")
            if path == prefix or path.startswith(prefix + "/"):
                environ["PATH_INFO"] = path[len(prefix):] or "/"
            environ["SCRIPT_NAME"] = prefix
            return original_app(environ, start_response)

        application.wsgi_app = prefixed_app
    application.run(host=args.anaconda_project_address, port=args.anaconda_project_port, debug=False)

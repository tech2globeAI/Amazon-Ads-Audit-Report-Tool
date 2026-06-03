#!/usr/bin/env python3
"""Web UI for Amazon Ads Audit Report generator."""

from __future__ import annotations

import traceback
from pathlib import Path

from flask import Flask, render_template, request, send_file
from io import BytesIO

from amazon_ads_audit import generate_audit_report_bytes

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    uploaded = request.files.get("bulk_sheet")
    if not uploaded or not uploaded.filename:
        return render_template(
            "index.html",
            error="Please choose an Amazon Ads bulk sheet (.xlsx) to upload.",
        ), 400

    filename = uploaded.filename
    if not filename.lower().endswith((".xlsx", ".xlsm")):
        return render_template(
            "index.html",
            error="Only Excel files (.xlsx, .xlsm) are supported.",
        ), 400

    try:
        file_bytes = uploaded.read()
        if not file_bytes:
            return render_template(
                "index.html",
                error="The uploaded file is empty.",
            ), 400

        out_bytes, out_name = generate_audit_report_bytes(file_bytes, filename)
        return send_file(
            BytesIO(out_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=out_name,
        )
    except Exception as exc:
        app.logger.error("Generate failed: %s", traceback.format_exc())
        return render_template(
            "index.html",
            error=str(exc),
        ), 500


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Run Amazon Ads Audit web UI")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=5000)
    parser.add_argument("--debug", action="store_true")
    args = parser.parse_args()
    print(f"Open in browser: http://{args.host}:{args.port}")
    app.run(host=args.host, port=args.port, debug=args.debug)


if __name__ == "__main__":
    main()

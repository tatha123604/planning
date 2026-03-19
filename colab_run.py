"""
Helper to run the FastAPI app on Google Colab with a public URL via ngrok.

Usage inside a Colab cell (after cloning the repo and `cd` into it):
    !pip install -r requirements-colab.txt
    import os
    os.environ["NGROK_AUTHTOKEN"] = "<your-ngrok-token>"  # required for stable URL
    !python colab_run.py --port 8000 --region in

The script opens an ngrok tunnel, prints the public URL, and then starts uvicorn.
"""

from __future__ import annotations

import argparse
import os
import sys

try:
    from pyngrok import conf, ngrok
except ModuleNotFoundError:  # pragma: no cover - only executed in missing-dep scenarios
    sys.stderr.write(
        "pyngrok is not installed. Run `pip install -r requirements-colab.txt` first.\n"
    )
    raise

import uvicorn

from src.app import app


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Expose FastAPI app on Colab via ngrok")
    parser.add_argument("--port", type=int, default=8000, help="Local port to serve FastAPI")
    parser.add_argument(
        "--region",
        type=str,
        default=os.environ.get("NGROK_REGION"),
        help="(optional) ngrok region; leave unset for ngrok v3/default",
    )
    parser.add_argument(
        "--auth-token",
        type=str,
        default=os.environ.get("NGROK_AUTHTOKEN"),
        help="ngrok auth token (set NGROK_AUTHTOKEN env var or pass here)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.auth_token:
        conf.get_default().auth_token = args.auth_token
    else:
        sys.stderr.write(
            "Warning: NGROK_AUTHTOKEN not set. ngrok may refuse to open a stable tunnel.\n"
        )

    connect_kwargs = {"addr": args.port, "proto": "http"}
    if args.region:
        connect_kwargs["region"] = args.region

    try:
        public_tunnel = ngrok.connect(**connect_kwargs)
    except Exception as exc:  # pragma: no cover - runtime/network dependent
        # Retry without region if the ngrok binary rejects the field (v3+)
        if "field region not found" in str(exc):
            public_tunnel = ngrok.connect(args.port, "http")
        else:
            raise

    print(f"Public URL: {public_tunnel.public_url}")
    print("Starting uvicorn... (Ctrl+C to stop)")

    uvicorn.run(app, host="0.0.0.0", port=args.port)


if __name__ == "__main__":
    main()

from __future__ import annotations

from sqlmodel import Session

from .app import refresh_ssts_snapshot
from .db import engine, init_db


def main() -> int:
    init_db()
    with Session(engine) as session:
        result = refresh_ssts_snapshot(session, force=True)
    print(result.get("message", "SSTS sync completed."))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

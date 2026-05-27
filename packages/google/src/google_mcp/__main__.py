import sys

from google_mcp.server import build


def main() -> None:
    app = sys.argv[1].lower() if len(sys.argv) > 1 else "sheets"
    build(app).run()


if __name__ == "__main__":
    main()

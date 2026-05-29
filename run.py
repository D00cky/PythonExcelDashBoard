import os

from waitress import serve

from app import create_app


def main() -> None:
    port = int(os.environ.get("PORT", "8000"))
    host = os.environ.get("HOST", "0.0.0.0")
    serve(create_app(), host=host, port=port)


if __name__ == "__main__":
    main()

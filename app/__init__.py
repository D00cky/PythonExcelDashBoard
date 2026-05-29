from flask import Flask


def create_app(config: dict | None = None) -> Flask:
    app = Flask(__name__)
    if config:
        app.config.update(config)

    from app.routes.main import bp as main_bp

    app.register_blueprint(main_bp)

    return app

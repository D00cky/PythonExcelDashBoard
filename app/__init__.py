from flask import Flask


def create_app(config: dict | None = None) -> Flask:
    config = dict(config or {})
    instance_path = config.pop("INSTANCE_PATH", None)
    app = Flask(__name__, instance_path=instance_path)
    app.config.update(config)

    from app.routes.main import bp as main_bp

    app.register_blueprint(main_bp)

    return app

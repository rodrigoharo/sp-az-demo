import os


def require_setting(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        raise RuntimeError(f"Missing required application setting: {name}")
    return value


def optional_setting(name: str) -> str | None:
    value = os.environ.get(name)
    return value if value else None


from starlette.config import Config


class DbConfig:
    user: str
    password: str
    server: str
    port: str
    db: str

    def __init__(self, user: str, password: str, server: str, port: str, db: str) -> None:
        self.user = user
        self.password = password
        self.server = server
        self.port = port
        self.db = db

    def get_path(self) -> str:
        return f'postgresql://{self.user}:{self.password}@{self.server}:{self.port}/{self.db}'


_config = Config(".env")

db_config = DbConfig(
    user=_config("POSTGRES_USER"),
    password=_config("POSTGRES_USER"),
    server=_config("POSTGRES_SERVER"),
    port=_config("POSTGRES_PORT"),
    db=_config("POSTGRES_DB"))

upload_dir = _config('UPLOAD_DIR')

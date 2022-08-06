from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from starlette.config import Config
from .config import db_config

config = Config(".env")

Base = declarative_base()

ECHO_LOG = True

engine = create_engine(db_config.get_path(), echo=ECHO_LOG)

Session = sessionmaker(bind=engine)
session = Session()

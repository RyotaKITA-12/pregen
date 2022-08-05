from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from starlette.config import Config

config = Config(".env")

Base = declarative_base()

user = config("POSTGRES_USER")
password = config("POSTGRES_USER")
server = config("POSTGRES_SERVER")
port = config("POSTGRES_PORT")
db = config("POSTGRES_DB")
PATH = f'postgresql://{user}:{password}@{server}:{port}/{db}'

ECHO_LOG = True

engine = create_engine(PATH, echo=ECHO_LOG)

Session = sessionmaker(bind=engine)
session = Session()

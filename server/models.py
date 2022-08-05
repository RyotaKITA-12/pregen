from datetime import datetime
import hashlib

from sqlalchemy import Column, String, DateTime
from sqlalchemy.sql.functions import current_timestamp
from sqlalchemy.dialects.mysql import INTEGER

from db import Base


class User(Base):
    """
    id         : 主キー
    username   : 名前
    password   : パスワード
    mail       : メールアドレス
    created_at : 登録日
    """
    __tablename__ = 'user'
    id = Column(
        'id',
        INTEGER(unsigned=True),
        primary_key=True,
        autoincrement=True,
    )
    username = Column('username', String(256))
    password = Column('password', String(256))
    mail = Column('mail', String(256))
    created_at = Column(
        'created_at',
        DateTime,
        default=datetime.now(),
        nullable=False,
        server_default=current_timestamp(),
    )

    def __init__(self, username, password, mail):
        self.username = username
        self.password = hashlib.md5(password.encode()).hexdigest()
        self.mail = mail
        self.created_at = datetime.now()

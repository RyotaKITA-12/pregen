import db
from models import User

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.requests import Request

app = FastAPI(title='pregen')
app.mount(
    '/server/views',
    StaticFiles(directory="/server/views"),
    name='static'
)
templates = Jinja2Templates(directory="views/templates")
jinja_env = templates.env


def test(request: Request):
    user = db.session.query(User).filter(User.username == 'admin').first()
    db.session.close()
    return templates.TemplateResponse('test.html',
                                      {'request': request,
                                       'user': user})

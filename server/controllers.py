from sys import prefix
from fastapi import FastAPI
from router.files import router as filesRouter
from fastapi.staticfiles import StaticFiles

app = FastAPI(title='pregen', prefix='/api')

# /static/template.pptx : テンプレートダウンロード用
app.mount('/static', StaticFiles(directory="tmp/data"), name='files')

app.include_router(filesRouter)

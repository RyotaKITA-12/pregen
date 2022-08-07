from sys import prefix
from fastapi import FastAPI
from router.files import router as filesRouter
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title='pregen', prefix='/api')
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"]
)

# /static/template.pptx : テンプレートダウンロード用
app.mount('/static', StaticFiles(directory="tmp/pptx/data"), name='files')

app.include_router(filesRouter)

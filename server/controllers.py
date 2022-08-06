from sys import prefix
from fastapi import FastAPI
from router.files import router as filesRouter

app = FastAPI(title='pregen', prefix='/api')

app.include_router(filesRouter)

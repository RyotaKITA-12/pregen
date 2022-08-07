import os
import shutil
from typing import BinaryIO
from fastapi import APIRouter, HTTPException, UploadFile
from fastapi.responses import FileResponse
from config import upload_dir

import app


router = APIRouter(prefix='/files')


@router.post("/upload_pptx")
async def upload_pptx(file: UploadFile):
    extension = file.filename.split(".")[-1] in ("pptx")
    if not extension:
        raise HTTPException(status_code=400, detail="invalid file format")
    save_file(file=file.file, filename=file.filename)
    app.main()
    return {
        'message': 'successfully saved the file',
        'filename': file.filename
    }


@router.get("{filename:path}")
async def download_file(filename: str):
    response = FileResponse(
        path=os.path.join(upload_dir, filename),
        filename=filename
    )
    return response


def save_file(filename: str, file: BinaryIO, dir: str = upload_dir) -> bool:
    if file:
        os.makedirs(dir, exist_ok=True)
        with open(os.path.join(dir, filename), 'w+b') as upload_dir:
            shutil.copyfileobj(file, upload_dir)
        return True
    return False

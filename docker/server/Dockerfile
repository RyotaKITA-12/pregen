FROM python:3.10-alpine

WORKDIR /server


RUN apk add --no-cache postgresql-libs
RUN apk add build-base jpeg-dev zlib-dev openjpeg-dev
RUN apk add gcc musl-dev postgresql-dev gfortran openblas-dev lapack-dev linux-headers
RUN apk add wget

COPY ./server/requirements.txt .
RUN pip install --upgrade pip setuptools
RUN wget http://gensen.dl.itc.u-tokyo.ac.jp/soft/pytermextract-0_02.zip \
    && unzip pytermextract-0_02.zip -d /pytermextract \
    && rm pytermextract-0_02.zip \
    && cd /pytermextract \
    && python setup.py install

RUN pip install --no-cache-dir --upgrade -r /server/requirements.txt

COPY . .

CMD ["uvicorn", "main:app", "--reload", "--host", "0.0.0.0", "--port", "8080"]

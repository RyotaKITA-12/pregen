FROM python:3.10-alpine

WORKDIR /server

COPY ./server/requirements.txt .

RUN apk add --no-cache postgresql-libs \
    && apk add build-base \
    && apk add --no-cache --virtual .build-deps gcc musl-dev postgresql-dev \
    && pip install --no-cache-dir --upgrade -r /server/requirements.txt \
    && apk --purge del .build-deps

COPY . .

CMD ["uvicorn", "main:app", "--reload", "--host", "0.0.0.0", "--port", "8080"]

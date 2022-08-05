FROM python:3.10-alpine

WORKDIR /app

COPY . .

RUN apk add --no-cache postgresql-libs \
    && apk add --no-cache --virtual .build-deps gcc musl-dev postgresql-dev \
    && pip install --no-cache-dir --upgrade -r /app/requirements.txt \
    && apk --purge del .build-deps

CMD ["uvicorn", "main:app", "--reload", "--host", "0.0.0.0", "--port", "8080"]

FROM node:18-alpine

WORKDIR /client
COPY . .

RUN npm install -g npm \
    npm install @material-ui/core \
    && rm -rf /var/lib/apt/lists/*

FROM node:20-bookworm

RUN apt-get update \
    && apt-get install -y \
       libreoffice \
       poppler-utils \
       tesseract-ocr \
       tesseract-ocr-ben \
       tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package.json ./
RUN npm install

COPY server.js ./

EXPOSE 10000

CMD ["npm", "start"]

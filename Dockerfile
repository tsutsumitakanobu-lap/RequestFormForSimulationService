# Dockerfile
FROM node:20-alpine

WORKDIR /workspace

# git を追加 ← これを追加！
RUN apk add --no-cache git

# claspとTypeScriptをグローバルインストール
RUN npm install -g @google/clasp typescript

COPY package*.json ./
RUN npm install 2>/dev/null || true

CMD ["sh"]
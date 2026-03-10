# Dockerfile
FROM node:20-alpine

WORKDIR /workspace

RUN apk add --no-cache git

RUN git config --global user.email "tsutsumi_takanobu@lapsys.co.jp" && \
    git config --global user.name "tsutsumitakanobu-lap"

# claspとTypeScriptをグローバルインストール
RUN npm install -g @google/clasp typescript

COPY package*.json ./
RUN npm install 2>/dev/null || true

CMD ["sh"]
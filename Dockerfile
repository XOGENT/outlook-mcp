FROM node:22-slim

WORKDIR /app

COPY package*.json ./
RUN npm ci --omit=dev

COPY server/ server/

ENV MCP_OUTLOOK_DATA_DIR=/data
ENV MCP_OUTLOOK_WORK_DIR=/data/downloads
ENV MCP_OUTLOOK_HEADLESS=true
ENV NODE_ENV=production

RUN mkdir -p /data/downloads && chown -R node:node /data

USER node

VOLUME /data

ENTRYPOINT ["node", "server/index.js"]

# PPTX render server for the Management Report pipeline
FROM node:20-slim

WORKDIR /app

# Install deps first (better layer caching)
COPY package.json package-lock.json* ./
RUN npm install --omit=dev --no-audit --no-fund

# App source
COPY pptx-api-server.mjs ./
COPY public ./public

# Generated decks land here (ephemeral; served immediately via /download)
RUN mkdir -p /app/output

ENV NODE_ENV=production
ENV PORT=3456
EXPOSE 3456

CMD ["node", "pptx-api-server.mjs"]

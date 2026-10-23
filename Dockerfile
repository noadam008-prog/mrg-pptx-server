# PPTX render server for the Management Report pipeline
FROM node:20-slim

WORKDIR /app

# Install exactly the locked deps (npm ci), production only
COPY package.json package-lock.json ./
RUN npm ci --omit=dev --no-audit --no-fund

# App source
COPY pptx-api-server.mjs ./
COPY public ./public

# Generated decks land here (ephemeral; served via /d/<id>/… links, purged after DOWNLOAD_TTL_HOURS)
RUN mkdir -p /app/output && chown -R node:node /app/output

ENV NODE_ENV=production
ENV PORT=3456
EXPOSE 3456

# Run as the unprivileged "node" user, not root (2026-10-23 security review)
USER node

CMD ["node", "pptx-api-server.mjs"]

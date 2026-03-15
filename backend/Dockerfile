FROM node:22-slim
# Install wget for healthcheck
RUN apt-get update && apt-get install -y --no-install-recommends wget && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY src/ ./src/

# Pre-create log directories so they exist at startup whether or not a volume
# is mounted.  Done as root (before USER switch) then chown'd to the node user
# (UID 1000) so the app can write to them without privilege escalation.
RUN mkdir -p /app/logs/feedback /app/logs/frontend \
    && chown -R node:node /app

# IC2: Run as non-root user for security (node user exists in official Node.js images)
USER node

CMD ["node", "src/server.js"]


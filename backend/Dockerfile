FROM node:22-slim
# Install wget for healthcheck
RUN apt-get update && apt-get install -y --no-install-recommends wget && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY src/ ./src/

# IC2: Run as non-root user for security (node user exists in official Node.js images)
RUN chown -R node:node /app
USER node

CMD ["node", "src/server.js"]


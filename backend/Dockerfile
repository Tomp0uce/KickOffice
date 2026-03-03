FROM node:22-alpine
# Install wget for healthcheck (not included in alpine by default)
RUN apk add --no-cache wget

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY src/ ./src/

CMD ["node", "src/server.js"]


FROM node:22-slim AS build-stage
WORKDIR /build

# Copy office-word-diff first (local dependency at same level as frontend)
COPY office-word-diff/ ./office-word-diff/

# Copy frontend package files
COPY frontend/package.json frontend/package-lock.json ./frontend/

# Install dependencies (from frontend directory where package.json expects ../office-word-diff)
WORKDIR /build/frontend
# Use npm install instead of npm ci for better compatibility with local file dependencies
RUN npm install --no-audit --no-fund

# Copy frontend source
COPY frontend/ ./

ARG VITE_BACKEND_URL=http://192.168.50.10:3003
ARG VITE_APP_VERSION=
ARG VITE_REQUEST_TIMEOUT_MS=45000
ENV VITE_BACKEND_URL=$VITE_BACKEND_URL
ENV VITE_APP_VERSION=$VITE_APP_VERSION
ENV VITE_REQUEST_TIMEOUT_MS=$VITE_REQUEST_TIMEOUT_MS

RUN npm run build

FROM nginx:stable
COPY --from=build-stage /build/frontend/dist /usr/share/nginx/html
COPY frontend/nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]

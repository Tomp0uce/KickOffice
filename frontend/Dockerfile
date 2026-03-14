FROM node:22-slim AS build-stage
WORKDIR /build

# Copy frontend package files
COPY frontend/package.json frontend/package-lock.json ./frontend/

# Install dependencies
WORKDIR /build/frontend
# ARCH-L1: Use npm ci for reproducible builds (local deps removed in Phase 4A)
RUN npm ci --no-audit --no-fund

# Copy frontend source
COPY frontend/ ./

# IH2: Replace with your actual server IP (e.g., http://192.168.x.x:3003 or http://localhost:3003)
ARG VITE_BACKEND_URL=http://localhost:3003
ARG VITE_APP_VERSION=
ARG VITE_REQUEST_TIMEOUT_MS=45000
ENV VITE_BACKEND_URL=$VITE_BACKEND_URL
ENV VITE_APP_VERSION=$VITE_APP_VERSION
ENV VITE_REQUEST_TIMEOUT_MS=$VITE_REQUEST_TIMEOUT_MS

RUN npm run build

# IC2: Use nginx-unprivileged for non-root security (listens on 8080 by default)
FROM nginxinc/nginx-unprivileged:stable
COPY --from=build-stage /build/frontend/dist /usr/share/nginx/html
COPY frontend/nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 8080
CMD ["nginx", "-g", "daemon off;"]

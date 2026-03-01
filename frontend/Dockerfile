FROM node:22-alpine AS build-stage
WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci

COPY . .

ARG VITE_BACKEND_URL=http://192.168.50.10:3003
ARG VITE_APP_VERSION=
ARG VITE_REQUEST_TIMEOUT_MS=45000
ENV VITE_BACKEND_URL=$VITE_BACKEND_URL
ENV VITE_APP_VERSION=$VITE_APP_VERSION
ENV VITE_REQUEST_TIMEOUT_MS=$VITE_REQUEST_TIMEOUT_MS

RUN npm run build

FROM nginx:alpine
COPY --from=build-stage /app/dist /usr/share/nginx/html
COPY nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]

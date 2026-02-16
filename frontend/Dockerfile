FROM node:22-alpine AS build-stage
WORKDIR /app

COPY package.json ./
RUN npm install

COPY . .

ARG VITE_BACKEND_URL=http://192.168.50.10:3003
ARG VITE_APP_VERSION=
ENV VITE_BACKEND_URL=$VITE_BACKEND_URL
ENV VITE_APP_VERSION=$VITE_APP_VERSION

RUN npm run build

FROM nginx:alpine
COPY --from=build-stage /app/dist /usr/share/nginx/html
COPY nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]

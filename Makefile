.PHONY: build up up-build

# Extract current short commit hash (or default to 1.0.0 if not available)
GIT_HASH := $(shell git rev-parse --short HEAD 2>/dev/null || echo "1.0.0")

export APP_VERSION := 1.0.0+$(GIT_HASH)

## Build containers without starting
build:
	@echo "Building KickOffice version $(APP_VERSION)..."
	docker-compose build

## Start containers normally
up:
	@echo "Starting KickOffice version $(APP_VERSION)..."
	docker-compose up -d

## Rebuild and start containers
up-build:
	@echo "Building and starting KickOffice version $(APP_VERSION)..."
	docker-compose up -d --build

#!/usr/bin/env bash

set -euo pipefail

APP_NAME="excel-to-word"
IMAGE_NAME="${APP_NAME}:latest"
CONTAINER_NAME="${APP_NAME}"
PORT_HOST="${PORT:-3000}"
PORT_CONTAINER="3002"

echo "[1/4] Building Docker image: ${IMAGE_NAME}"
docker build -t "${IMAGE_NAME}" .

echo "[2/4] Stopping and removing existing container if present"
if docker ps -a --format '{{.Names}}' | grep -q "^${CONTAINER_NAME}$"; then
  docker rm -f "${CONTAINER_NAME}" || true
fi

echo "[3/4] Starting container ${CONTAINER_NAME} on port ${PORT_HOST} -> ${PORT_CONTAINER}"
docker run -d \
  --name "${CONTAINER_NAME}" \
  -p "${PORT_HOST}:${PORT_CONTAINER}" \
  --restart unless-stopped \
  "${IMAGE_NAME}"

echo "[4/4] Done. Service should be reachable at: http://localhost:${PORT_HOST}/"



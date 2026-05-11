FROM node:20-slim

# Install system dependencies for PDF rendering and OCR
RUN apt-get update && apt-get install -y --no-install-recommends \
    poppler-utils \
    qpdf \
    imagemagick \
    tesseract-ocr \
    tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

# Fix ImageMagick policy to allow PDF operations
RUN sed -i 's/rights="none" pattern="PDF"/rights="read|write" pattern="PDF"/' /etc/ImageMagick-6/policy.xml 2>/dev/null || true

WORKDIR /opt/render/project/src

# Copy package files first for better caching
COPY package*.json ./

# Install ALL dependencies (including dev) for the build step
RUN npm ci

# Copy source and build.
# `npm run build` (script/build.ts) handles copying server/estimator-data.json
# into dist/ on its own. A previous version of this Dockerfile copied a stale
# root-level estimator-data.json over the newly-built one — don't add a cp here.
COPY . .
RUN npm run build

# Prune dev dependencies after build
RUN npm prune --omit=dev

# Create data directory
RUN mkdir -p data

ENV NODE_ENV=production
ENV PORT=10000

EXPOSE 10000

CMD ["node", "dist/index.cjs"]

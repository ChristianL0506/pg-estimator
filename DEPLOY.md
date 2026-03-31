# Picou Group Takeoff & Estimating Tool — Quick Deploy

## Requirements
- Ubuntu 22.04+ server (DigitalOcean $24/month recommended)
- Node.js 20+
- Anthropic API key (console.anthropic.com)
- Tools: poppler-utils, imagemagick, tesseract-ocr, qpdf

## Install Dependencies
```bash
apt update && apt upgrade -y
curl -fsSL https://deb.nodesource.com/setup_20.x | bash -
apt install -y nodejs poppler-utils imagemagick tesseract-ocr qpdf build-essential
```

## Deploy the App
```bash
mkdir -p /opt/pg-estimator
cd /opt/pg-estimator
# Upload/copy the pg-unified folder here

cd pg-unified
npm install
npm run build
cp estimator-data.json dist/estimator-data.json
```

## Configure
```bash
export ANTHROPIC_API_KEY="sk-ant-your-key-here"
```

## Start
```bash
NODE_ENV=production node dist/index.cjs
# App runs at http://YOUR_SERVER_IP:5000
# Login: admin / picougroup
```

## Keep Running (PM2)
```bash
npm install -g pm2
ANTHROPIC_API_KEY="sk-ant-your-key-here" pm2 start dist/index.cjs --name pg-estimator
pm2 startup
pm2 save
```

## Backup
```bash
# Weekly: copy the database
cp data/pg-unified.db ~/backups/pg-unified-$(date +%Y%m%d).db
```

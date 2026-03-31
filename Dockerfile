# Stage 1: Build frontend static files
FROM node:20-slim AS frontend-build
WORKDIR /app/frontend
COPY python/frontend/package.json python/frontend/package-lock.json ./
RUN npm ci
COPY python/frontend/ ./
RUN npm run build

# Stage 2: Python runtime
FROM python:3.12-slim
WORKDIR /app

# Install Python dependencies
COPY python/requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy backend source code
COPY python/config.py python/main.py python/models.py ./
COPY python/document_parser.py python/openai_client.py ./
COPY python/excel_processor_enhanced.py ./

# Copy Excel template
COPY python/assets/ ./assets/

# Copy frontend static build from stage 1
COPY --from=frontend-build /app/frontend/out ./frontend/out

# Create required directories
RUN mkdir -p uploads temp

# Expose port (Render sets $PORT)
EXPOSE 10000

# Start server
CMD uvicorn main:app --host 0.0.0.0 --port ${PORT:-10000} --timeout-keep-alive 1500

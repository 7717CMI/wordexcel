# Word-Excel Processor

A production-ready application for extracting market research data from Word documents and generating Excel files with preserved macros.

## Features

- **Document Processing**: Extract structured data from Word documents
- **Excel Generation**: Create Excel files with preserved macros and formatting
- **Bulk Processing**: Process multiple files simultaneously with WebSocket updates
- **Real-time Updates**: Live progress tracking via WebSocket connections
- **Data Review**: Review and edit extracted data before Excel generation

## Tech Stack

### Backend
- FastAPI (Python)
- OpenAI API for data extraction
- python-docx for Word processing
- openpyxl for Excel generation
- WebSocket for real-time updates

### Frontend
- Next.js 14 (TypeScript)
- React 18
- Tailwind CSS
- WebSocket client for live updates

## Prerequisites

- Python 3.8+
- Node.js 16+
- npm or yarn
- OpenAI API key

## Installation

1. **Install Python dependencies:**
```bash
cd python
pip install -r requirements.txt
```

2. **Install and build frontend:**
```bash
cd python/frontend
npm install
npm run build
```

3. **Configure environment:**
Create a `python/.env` file:
```env
OPENAI_API_KEY=your_api_key_here
```

The frontend calls the API with same-origin relative paths, so no API URL
needs to be configured. Other optional settings:

| Variable | Default | Purpose |
| --- | --- | --- |
| `OPENAI_API_KEY` | _(required)_ | Key used for extraction |
| `OPENAI_MODEL` | `gpt-4o-mini` | Extraction model |
| `OPENAI_MAX_TOKENS` | `4000` | Response cap; raise for very large documents |
| `MAX_DOCUMENT_CHARS` | `200000` | Document text is truncated to this before the model call |
| `MAX_FILE_SIZE` | `52428800` | Upload limit in bytes (50 MB) |
| `MAX_BULK_FILES` | `50` | Files per bulk batch |
| `DOWNLOAD_CONFIRM_TIMEOUT` | `30` | Seconds to wait for the browser to confirm a download |
| `DOWNLOAD_RETENTION_SECONDS` | `300` | How long a generated file stays on disk after being served |
| `CORS_ORIGINS` | _(empty)_ | Comma-separated extra origins; unnecessary for same-origin deploys |

## Running the Application

### Development Mode

1. **Start backend:**
```bash
cd python
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

2. **Start frontend:**
```bash
cd python/frontend
npm run dev
```

### Production Mode

The frontend is a static export (`next.config.js` sets `output: 'export'`),
which FastAPI serves directly. Build it once, then run only the backend:

```bash
cd python/frontend && npm ci && npm run build
cd .. && uvicorn main:app --host 0.0.0.0 --port 8000
```

The whole app is then available on port 8000. Run a single worker: bulk jobs
and the WebSocket progress channel hold state in process memory, so multiple
workers would deliver progress to the wrong connection.

## Deployment (Render)

`render.yaml` provisions a single Docker web service. The build needs both
Node (for the Next.js static export) and Python (for FastAPI), which Render's
native Python runtime cannot provide, so the multi-stage `Dockerfile` is the
supported path.

1. Point Render at this repo; it picks up `render.yaml` as a Blueprint.
2. Set `OPENAI_API_KEY` in the Render dashboard (it is marked `sync: false`
   so it is never committed).
3. Deploy. Render injects `$PORT`; the container binds to it, and
   `/api/health` is used as the health check.

Uploaded and generated files live on the container's ephemeral disk and are
cleaned up automatically, so no persistent disk is required.

## API Endpoints

- `GET /api/health` - Health check
- `POST /api/upload` - Upload a Word document, returns a `fileId`
- `POST /api/process` - Extract data from an uploaded document
- `POST /api/generate-excel` - Generate the .xlsm from extracted data
- `POST /api/independent-bulk-process` - Accept a batch; returns immediately
  and reports progress over the WebSocket
- `GET /api/download/{filename}` - Download a generated file
- `WS /ws` - WebSocket endpoint for real-time updates

Any unmatched non-`/api` path serves the frontend (SPA routing).

## Project Structure

```
wordexcel/
├── python/                 # Backend application
│   ├── main.py            # FastAPI application
│   ├── document_parser.py # Word document processing
│   ├── excel_processor_enhanced.py # Excel generation
│   ├── openai_client.py  # OpenAI integration
│   ├── models.py         # Pydantic models
│   ├── config.py         # Configuration
│   ├── requirements.txt  # Python dependencies
│   └── frontend/         # Next.js frontend
│       ├── app/          # App router pages
│       ├── components/   # React components
│       └── package.json  # Node dependencies
├── Dockerfile            # Multi-stage build used by Render
├── render.yaml           # Render Blueprint
└── README.md             # This file
```

## Deployment Considerations

### Security
- Always use HTTPS in production
- Secure your OpenAI API key
- Implement rate limiting
- Add authentication if needed

### Performance
- Use a reverse proxy (Nginx/Apache)
- Enable caching where appropriate
- Consider CDN for static assets
- Monitor memory usage

### Scaling
- Use PM2 cluster mode for multiple instances
- Consider containerization with Docker
- Implement load balancing for high traffic
- Use cloud storage for file uploads

## Troubleshooting

### Common Issues

1. **Render deploy fails with "no open ports detected":**
   - Check the logs for a startup traceback. The app starts without
     `OPENAI_API_KEY` (extraction requests fail individually instead), so a
     boot failure points at a different misconfiguration.

2. **"Failed to process document: OPENAI_API_KEY is not configured":**
   - Set `OPENAI_API_KEY` in the Render dashboard and redeploy.

3. **Extraction fails with "AI response was truncated":**
   - The document produced more JSON than `OPENAI_MAX_TOKENS` allows.
     Raise it, or lower `MAX_DOCUMENT_CHARS`.

4. **"Extracted data is incomplete":**
   - The model could not find the expected market research fields in the
     document. Check the server log for the parsed keys.

5. **Build errors:**
   - Clear node_modules and reinstall
   - Check Node.js and Python versions
   - Review TypeScript errors

## Support

For issues or questions, please check the documentation or create an issue in the repository.

## License

[Your License Here]

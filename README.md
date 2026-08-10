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
- DeepSeek (OpenAI-compatible API) for data extraction
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
- DeepSeek API key (or an OpenAI key with `LLM_PROVIDER=openai`)

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
DEEPSEEK_API_KEY=your_api_key_here
```

The frontend calls the API with same-origin relative paths, so no API URL
needs to be configured. Other optional settings:

| Variable | Default | Purpose |
| --- | --- | --- |
| `LLM_PROVIDER` | `deepseek` | `deepseek` or `openai` |
| `DEEPSEEK_API_KEY` | _(required for deepseek)_ | Key used for extraction |
| `OPENAI_API_KEY` | _(required for openai)_ | Key used when `LLM_PROVIDER=openai` |
| `LLM_MODEL` | `deepseek-chat` / `gpt-4o-mini` | Extraction model |
| `LLM_BASE_URL` | provider default | Override the API endpoint |
| `LLM_MAX_TOKENS` | `4000` | Response cap; raise for very large documents |
| `LLM_TIMEOUT` | `120` | Per-request timeout in seconds |
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

`next dev` serves the app on port 3000 while the API lives on 8000, so the
dev server proxies `/api/*` to the backend (see `next.config.js`). Next's
proxy does not forward WebSocket upgrades, so the progress socket needs an
explicit URL. Both default to `http://127.0.0.1:8000`; override if the
backend is elsewhere:

```bash
API_PROXY_TARGET=http://127.0.0.1:8000 \
NEXT_PUBLIC_WS_URL=ws://127.0.0.1:8000/ws \
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
2. Set `DEEPSEEK_API_KEY` in the Render dashboard (it is marked `sync: false`
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
│   ├── llm_client.py     # DeepSeek/OpenAI integration
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
- Keep the API key in the environment; never commit `python/.env`
- There is no authentication - anyone who can reach the service can spend
  your API credits. Put it behind auth before exposing it publicly
- Consider rate limiting the upload and bulk endpoints

### Scaling
- Bulk job state and WebSocket connections live in process memory, so the
  service runs as a single worker. Multiple workers need shared state
  (Redis) and a job queue before they will work correctly
- Files are written to local disk and cleaned up on a timer; a multi-instance
  deployment would need object storage instead

## Troubleshooting

### Common Issues

1. **Render deploy fails with "no open ports detected":**
   - Check the logs for a startup traceback. The app starts without
     an API key (extraction requests fail individually instead), so a
     boot failure points at a different misconfiguration.

2. **"Failed to process document: DEEPSEEK_API_KEY is not configured":**
   - Set `DEEPSEEK_API_KEY` in the Render dashboard and redeploy.

3. **Extraction fails with "AI response was truncated":**
   - The document produced more JSON than `LLM_MAX_TOKENS` allows.
     Raise `LLM_MAX_TOKENS`, or lower `MAX_DOCUMENT_CHARS`.

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

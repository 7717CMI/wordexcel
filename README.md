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

### Quick Start (Windows)

```batch
deploy-production.bat
```

### Quick Start (Linux/Mac)

```bash
chmod +x deploy-production.sh
./deploy-production.sh
```

### Manual Installation

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
Create `python/.env` file:
```env
OPENAI_API_KEY=your_api_key_here
API_URL=http://localhost:8000
NEXT_PUBLIC_API_URL=http://localhost:8000
```

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

#### Using PM2 (Recommended)

```bash
# Install PM2 globally
npm install -g pm2

# Start all services
pm2 start ecosystem.config.js

# View logs
pm2 logs

# Stop all services
pm2 stop all
```

#### Manual Start

1. **Backend:**
```bash
cd python
uvicorn main:app --host 0.0.0.0 --port 8000 --workers 4
```

2. **Frontend:**
```bash
cd python/frontend
npm start
```

## API Endpoints

- `POST /extract` - Extract data from a Word document
- `POST /generate-excel` - Generate Excel from extracted data
- `POST /process-file` - Process single file (extract + generate)
- `POST /bulk-process` - Process multiple files
- `WS /ws` - WebSocket endpoint for real-time updates

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
├── ecosystem.config.js    # PM2 configuration
├── deploy-production.bat  # Windows deployment
├── deploy-production.sh   # Linux/Mac deployment
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

## Monitoring

- Check logs in `logs/` directory
- Monitor WebSocket connections
- Track API response times
- Set up alerts for errors

## Troubleshooting

### Common Issues

1. **WebSocket connection fails:**
   - Check firewall settings
   - Ensure ports 8000 and 3000 are open
   - Verify CORS settings

2. **File processing errors:**
   - Check file permissions
   - Ensure temp/uploads directories exist
   - Verify OpenAI API key is valid

3. **Build errors:**
   - Clear node_modules and reinstall
   - Check Node.js and Python versions
   - Review TypeScript errors

## Support

For issues or questions, please check the documentation or create an issue in the repository.

## License

[Your License Here]

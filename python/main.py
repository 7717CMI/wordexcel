import os
import time
import logging
import json
import asyncio
from typing import Dict, Any
from fastapi import FastAPI, File, UploadFile, HTTPException, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

from config import Config
from models import ProcessingResult, FileUploadResponse, ApiResponse
from document_parser import (
    extract_text_from_word,
    preprocess_text,
    normalize_extracted_data,
    validate_extracted_data,
    calculate_confidence
)
from openai_client import extract_market_data
from excel_processor_enhanced import ExcelProcessorEnhanced

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize FastAPI app
app = FastAPI(
    title="Word to Excel Processor API",
    description="AI-powered market research data extraction and Excel generation",
    version="1.0.0"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# WebSocket connection manager
class ConnectionManager:
    def __init__(self):
        self.active_connections: list[WebSocket] = []
        self.pending_downloads: dict[str, asyncio.Event] = {}  # Track pending downloads

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket: WebSocket):
        self.active_connections.remove(websocket)

    async def send_personal_message(self, message: str, websocket: WebSocket):
        await websocket.send_text(message)

    async def broadcast(self, message: str):
        for connection in self.active_connections:
            try:
                await connection.send_text(message)
            except Exception as e:
                logger.error(f"WebSocket broadcast error: {e}")

    async def wait_for_download_confirmation(self, file_id: str, timeout: int = 30):
        """Wait for frontend to confirm file download"""
        if file_id not in self.pending_downloads:
            self.pending_downloads[file_id] = asyncio.Event()
        
        try:
            await asyncio.wait_for(self.pending_downloads[file_id].wait(), timeout=timeout)
            logger.info(f"Download confirmed for file: {file_id}")
            return True
        except asyncio.TimeoutError:
            logger.warning(f"Download confirmation timeout for file: {file_id}")
            return False
        finally:
            # Clean up
            if file_id in self.pending_downloads:
                del self.pending_downloads[file_id]

    def confirm_download(self, file_id: str):
        """Called when frontend confirms download"""
        if file_id in self.pending_downloads:
            self.pending_downloads[file_id].set()
            logger.info(f"Download confirmation received for file: {file_id}")

manager = ConnectionManager()

# Ensure required directories exist
Config.ensure_directories()

@app.get("/")
async def root():
    """Health check endpoint"""
    return {
        "message": "Word to Excel Processor API",
        "status": "running",
        "version": "1.0.0"
    }

@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    """WebSocket endpoint for real-time progress updates"""
    await manager.connect(websocket)
    try:
        while True:
            # Receive messages from frontend
            data = await websocket.receive_text()
            try:
                message = json.loads(data)
                if message.get('type') == 'download_confirmed':
                    file_id = message.get('file_id')
                    if file_id:
                        manager.confirm_download(file_id)
                        logger.info(f"Download confirmation received for: {file_id}")
            except json.JSONDecodeError:
                logger.warning(f"Invalid JSON received: {data}")
    except WebSocketDisconnect:
        manager.disconnect(websocket)

@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    """Upload Word document endpoint"""
    try:
        # Validate file type
        if not file.filename.lower().endswith(('.docx', '.doc')):
            raise HTTPException(status_code=400, detail="Only Word documents (.docx, .doc) are allowed")

        # Validate file size
        if file.size > Config.MAX_FILE_SIZE:
            raise HTTPException(status_code=400, detail="File size exceeds maximum limit")

        # Generate unique file ID
        file_id = f"{int(time.time() * 1000)}_{file.filename}"
        file_path = os.path.join(Config.UPLOAD_DIR, file_id)

        # Save uploaded file
        with open(file_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)

        logger.info(f"File uploaded successfully: {file_id}")

        return FileUploadResponse(
            success=True,
            fileId=file_id,
            message="File uploaded successfully"
        )

    except HTTPException:
        raise
    except Exception as error:
        logger.error(f"Upload error: {error}")
        raise HTTPException(status_code=500, detail="Failed to upload file")

@app.post("/api/process")
async def process_document(data: Dict[str, Any]):
    """Process uploaded document endpoint"""
    try:
        file_id = data.get('fileId')
        if not file_id:
            raise HTTPException(status_code=400, detail="File ID is required")

        file_path = os.path.join(Config.UPLOAD_DIR, file_id)
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="File not found")

        logger.info(f"Processing document: {file_id}")

        # Extract text from Word document
        extracted_text = extract_text_from_word(file_path)
        logger.info(f"Text extracted, length: {len(extracted_text)}")

        # Preprocess text
        processed_text = preprocess_text(extracted_text)
        logger.info(f"Text preprocessed, length: {len(processed_text)}")

        # Extract market data using AI with retry mechanism
        extracted_data = None
        max_retries = 2
        
        for attempt in range(max_retries):
            try:
                logger.info(f"AI extraction attempt {attempt + 1}/{max_retries}")
                extracted_data = extract_market_data(processed_text)
                logger.info("AI extraction completed")
                logger.info(f"Raw extracted data keys: {list(extracted_data.keys()) if isinstance(extracted_data, dict) else 'Not a dict'}")
                
                # Check if AI extraction returned valid data
                if extracted_data and isinstance(extracted_data, dict) and 'market' in extracted_data:
                    logger.info("AI extraction successful with valid data structure")
                    break
                else:
                    logger.warning(f"AI extraction attempt {attempt + 1} returned invalid data structure")
                    if attempt < max_retries - 1:
                        logger.info("Retrying AI extraction...")
                        continue
                    else:
                        raise Exception("AI extraction failed after all retry attempts - invalid data structure")
                        
            except Exception as ai_error:
                logger.error(f"AI extraction attempt {attempt + 1} failed: {ai_error}")
                if attempt < max_retries - 1:
                    logger.info("Retrying AI extraction...")
                    continue
                else:
                    raise Exception(f"AI extraction failed after all retry attempts: {str(ai_error)}")
        
        if not extracted_data:
            raise Exception("AI extraction failed - no data returned")

        # Normalize extracted data
        normalized_data = normalize_extracted_data(extracted_data)
        logger.info("Data normalization completed")

        # Validate extracted data
        validation_result = validate_extracted_data(normalized_data)
        logger.info(f"Data validation result: {validation_result}")

        # Calculate confidence score
        confidence = calculate_confidence(normalized_data)
        logger.info(f"Confidence score: {confidence}")

        # Convert Pydantic model to dictionary if needed
        data_for_response = normalized_data
        if hasattr(normalized_data, 'dict'):
            data_for_response = normalized_data.dict()
        elif hasattr(normalized_data, 'model_dump'):
            data_for_response = normalized_data.model_dump()

        # Clean up uploaded file after successful processing
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                logger.info(f"Cleaned up uploaded file after processing: {file_id}")
        except Exception as cleanup_error:
            logger.warning(f"Failed to cleanup uploaded file {file_id}: {cleanup_error}")

        return ProcessingResult(
            success=True,
            data=data_for_response,
            confidence=confidence
        )

    except HTTPException:
        raise
    except Exception as error:
        logger.error(f"Processing error: {error}")
        
        # Clean up uploaded file on error
        try:
            if 'file_id' in locals() and 'file_path' in locals():
                if os.path.exists(file_path):
                    os.remove(file_path)
                    logger.info(f"Cleaned up uploaded file after error: {file_id}")
        except Exception as cleanup_error:
            logger.warning(f"Failed to cleanup uploaded file {file_id} after error: {cleanup_error}")
        
        raise HTTPException(status_code=500, detail="Failed to process document")

@app.post("/api/generate-excel")
async def generate_excel(data: Dict[str, Any]):
    """Generate Excel file endpoint"""
    try:
        file_id = data.get('fileId')
        extracted_data = data.get('data')

        if not file_id or not extracted_data:
            raise HTTPException(status_code=400, detail="File ID and data are required")

        if not all(key in extracted_data for key in ['market', 'segments', 'players']):
            raise HTTPException(status_code=400, detail="Invalid data structure")

        template_path = Config.EXCEL_TEMPLATE_PATH
        logger.info(f'Using template path: {template_path}')
        logger.info(f'Current working directory: {os.getcwd()}')

        # Use enhanced Excel processor for better macro preservation
        excel_processor = ExcelProcessorEnhanced(template_path)
        logger.info('Populating Excel template with extracted data...')
        excel_processor.populate_data(extracted_data)

        # Use market name for filename (same as D2 cell content)
        market_name = extracted_data.get('market', {}).get('market_name', 'Market Data')
        
        # Clean market name for filename (remove only problematic characters, keep spaces)
        import re
        clean_market_name = re.sub(r'[<>:"/\\|?*\x00]', '', market_name)  # Cross-platform filename sanitization
        clean_market_name = clean_market_name.strip()  # Remove leading/trailing whitespace
        
        # Fallback to timestamp if market name is empty
        if not clean_market_name:
            timestamp = int(time.time() * 1000)
            clean_market_name = f"Market Data {timestamp}"
        
        output_filename = f"{clean_market_name}.xlsm"
        
        # Ensure temp directory exists
        temp_dir = "temp"
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
        
        output_path = os.path.join(temp_dir, output_filename)

        logger.info('Saving populated Excel file with enhanced macro preservation...')
        excel_processor.save_file(output_path)

        # Clean up
        excel_processor.cleanup()

        logger.info('Excel file generated successfully with preserved macros')

        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename=\"{output_filename}\"",
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }
        )

    except HTTPException:
        raise
    except Exception as error:
        logger.error(f"Excel generation error: {error}")
        raise HTTPException(status_code=500, detail="Failed to generate Excel file")

@app.post("/api/independent-bulk-process")
async def independent_bulk_process(files: list[UploadFile] = File(...)):
    """Process multiple files independently with simple error handling"""
    try:
        if not files:
            raise HTTPException(status_code=400, detail="No files provided")
        
        if len(files) > 50:
            raise HTTPException(status_code=400, detail="Maximum 50 files allowed per batch")
        
        logger.info(f"Starting bulk processing of {len(files)} files")
        
        # Send initial progress update via WebSocket
        await manager.broadcast(json.dumps({
            'type': 'bulk_start',
            'total_files': len(files),
            'message': f'Starting bulk processing of {len(files)} files...'
        }))
        
        # Process files sequentially
        results = []
        successful = []
        failed = []
        
        for i, file in enumerate(files):
            logger.info(f"Processing file {i+1}/{len(files)}: {file.filename}")
            file_id = None
            
            # Send file start notification via WebSocket
            await manager.broadcast(json.dumps({
                'type': 'file_start',
                'file_index': i,
                'filename': file.filename,
                'message': f'Starting file {i+1}/{len(files)}: {file.filename}'
            }))
            
            try:
                # Step 1: Save uploaded file
                logger.info(f"Step 1: Saving uploaded file {file.filename}")
                file_id = await _save_uploaded_file(file)
                logger.info(f"File saved successfully: {file_id}")
                
                # Step 2: Process document
                logger.info(f"Step 2: Processing document {file_id}")
                processing_result = await _process_document(file_id)
                
                if not processing_result.success:
                    logger.error(f"Document processing failed for {file.filename}: {processing_result.error}")
                    raise Exception(f"Document processing failed: {processing_result.error}")
                
                logger.info(f"Document processing successful for {file.filename}")
                
                # Step 3: Generate Excel
                logger.info(f"Step 3: Generating Excel for {file.filename}")
                data_for_excel = processing_result.data
                if hasattr(processing_result.data, 'dict'):
                    data_for_excel = processing_result.data.dict()
                elif hasattr(processing_result.data, 'model_dump'):
                    data_for_excel = processing_result.data.model_dump()
                
                logger.info(f"Data structure for Excel: {type(data_for_excel)} with keys: {list(data_for_excel.keys()) if isinstance(data_for_excel, dict) else 'Not a dict'}")
                
                excel_result = await _generate_excel(file_id, data_for_excel)
                
                # Check if Excel generation failed
                if excel_result.get('error'):
                    logger.error(f"Excel generation failed for {file.filename}: {excel_result['error']}")
                    raise Exception(f"Excel generation failed: {excel_result['error']}")
                
                logger.info(f"Excel generation successful for {file.filename}: {excel_result['filename']}")
                
                # Clean up uploaded file
                try:
                    uploaded_file_path = os.path.join(Config.UPLOAD_DIR, file_id)
                    if os.path.exists(uploaded_file_path):
                        os.remove(uploaded_file_path)
                        logger.info(f"Cleaned up uploaded file: {file_id}")
                except Exception as cleanup_error:
                    logger.warning(f"Failed to cleanup uploaded file {file_id}: {cleanup_error}")
                
                # Add successful result
                result = {
                    'filename': file.filename,
                    'file_id': file_id,
                    'market_name': data_for_excel.get('market', {}).get('market_name', 'Unknown Market'),
                    'excel_filename': excel_result['filename'],
                    'excel_path': excel_result['path'],
                    'success': True,
                    'status': 'completed'
                }
                results.append(result)
                successful.append(result)
                logger.info(f"Successfully processed file {i+1}: {file.filename}")
                
                # Send real-time update via WebSocket
                file_id = excel_result['filename']
                await manager.broadcast(json.dumps({
                    'type': 'file_completed',
                    'file_index': i,
                    'filename': file.filename,
                    'excel_filename': excel_result['filename'],
                    'download_url': f'/api/download/{excel_result["filename"]}',
                    'market_name': data_for_excel.get('market', {}).get('market_name', 'Unknown Market'),
                    'success': True,
                    'file_id': file_id,  # Add file_id for download confirmation
                    'message': f'File {i+1}/{len(files)} completed: {file.filename}'
                }))
                
                # Wait for download confirmation before proceeding to next file
                logger.info(f"Waiting for download confirmation for: {file_id}")
                download_confirmed = await manager.wait_for_download_confirmation(file_id, timeout=60)
                
                if download_confirmed:
                    logger.info(f"Download confirmed for {file_id}, proceeding to next file")
                else:
                    logger.warning(f"Download confirmation timeout for {file_id}, proceeding anyway")
                
            except Exception as error:
                logger.error(f"Error processing file {file.filename}: {error}")
                logger.error(f"Error type: {type(error).__name__}")
                logger.error(f"Error details: {str(error)}")
                
                # Clean up uploaded file on failure
                if file_id:
                    try:
                        uploaded_file_path = os.path.join(Config.UPLOAD_DIR, file_id)
                        if os.path.exists(uploaded_file_path):
                            os.remove(uploaded_file_path)
                            logger.info(f"Cleaned up uploaded file after failure: {file_id}")
                    except Exception as cleanup_error:
                        logger.warning(f"Failed to cleanup uploaded file {file_id} after failure: {cleanup_error}")
                
                # Add failed result
                result = {
                    'filename': file.filename,
                    'file_id': file_id,
                    'market_name': 'Unknown Market',
                    'excel_filename': None,
                    'excel_path': None,
                    'success': False,
                    'error': str(error),
                    'status': 'failed'
                }
                results.append(result)
                failed.append(result)
                
                # Send real-time update via WebSocket for failed file
                await manager.broadcast(json.dumps({
                    'type': 'file_failed',
                    'file_index': i,
                    'filename': file.filename,
                    'error': str(error),
                    'success': False,
                    'message': f'File {i+1}/{len(files)} failed: {file.filename} - {str(error)}'
                }))
        
        logger.info(f"Bulk processing completed: {len(successful)} successful, {len(failed)} failed")
        
        # Send final completion update via WebSocket
        await manager.broadcast(json.dumps({
            'type': 'bulk_complete',
            'total_files': len(files),
            'successful': len(successful),
            'failed': len(failed),
            'message': f'Bulk processing completed: {len(successful)} successful, {len(failed)} failed'
        }))
        
        return ApiResponse(
            success=True,
            message=f"Bulk processing completed: {len(successful)} successful, {len(failed)} failed",
            data={
                'total_files': len(files),
                'successful': len(successful),
                'failed': len(failed),
                'results': results
            }
        )
        
    except HTTPException:
        raise
    except Exception as error:
        logger.error(f"Bulk processing error: {error}")
        return ApiResponse(
            success=False,
            message=f"Bulk processing failed: {str(error)}",
            data={
                'total_files': len(files) if 'files' in locals() else 0,
                'successful': 0,
                'failed': len(files) if 'files' in locals() else 0,
                'results': []
            }
        )


async def _process_single_file_independently(file: UploadFile, index: int) -> Dict[str, Any]:
    """Process a single file independently with error isolation"""
    try:
        logger.info(f"Starting independent processing of file {index + 1}: {file.filename}")
        
        # Step 1: Save uploaded file
        file_id = await _save_uploaded_file(file)
        
        # Step 2: Process document
        processing_result = await _process_document(file_id)
        
        if not processing_result.success:
            raise Exception(f"Document processing failed: {processing_result.error or 'Unknown error'}")
        
        # Step 3: Generate Excel
        data_for_excel = processing_result.data
        if hasattr(processing_result.data, 'dict'):
            data_for_excel = processing_result.data.dict()
        elif hasattr(processing_result.data, 'model_dump'):
            data_for_excel = processing_result.data.model_dump()
        
        excel_result = await _generate_excel(file_id, data_for_excel)
        
        # Clean up uploaded file
        try:
            uploaded_file_path = os.path.join(Config.UPLOAD_DIR, file_id)
            if os.path.exists(uploaded_file_path):
                os.remove(uploaded_file_path)
                logger.info(f"Cleaned up uploaded file: {file_id}")
        except Exception as cleanup_error:
            logger.warning(f"Failed to cleanup uploaded file {file_id}: {cleanup_error}")
        
        logger.info(f"Successfully processed file {index + 1}: {file.filename}")
        
        return {
            'filename': file.filename,
            'file_id': file_id,
            'market_name': data_for_excel.get('market', {}).get('market_name', 'Unknown Market'),
            'excel_filename': excel_result['filename'],
            'excel_path': excel_result['path'],
            'success': True,
            'status': 'completed'
        }
        
    except Exception as error:
        logger.error(f"Error processing file {file.filename}: {error}")
        
        # Clean up uploaded file on failure
        try:
            if 'file_id' in locals():
                uploaded_file_path = os.path.join(Config.UPLOAD_DIR, file_id)
                if os.path.exists(uploaded_file_path):
                    os.remove(uploaded_file_path)
                    logger.info(f"Cleaned up uploaded file after failure: {file_id}")
        except Exception as cleanup_error:
            logger.warning(f"Failed to cleanup uploaded file {file_id} after failure: {cleanup_error}")
        
        return {
            'filename': file.filename,
            'file_id': file_id if 'file_id' in locals() else None,
            'market_name': 'Unknown Market',
            'excel_filename': None,
            'excel_path': None,
            'success': False,
            'error': str(error),
            'status': 'failed'
        }

# Remove timeout functions - they cause connection issues

async def _save_uploaded_file(file: UploadFile) -> str:
    """Save uploaded file"""
    try:
        # Validate file type
        if not file.filename.lower().endswith(('.docx', '.doc')):
            logger.error(f"Invalid file type for {file.filename}: Only Word documents (.docx, .doc) are allowed")
            raise HTTPException(status_code=400, detail="Only Word documents (.docx, .doc) are allowed")
        
        # Validate file size (50MB limit)
        content = await file.read()
        if len(content) > Config.MAX_FILE_SIZE:
            logger.error(f"File too large for {file.filename}: {len(content)} bytes exceeds {Config.MAX_FILE_SIZE} limit")
            raise HTTPException(status_code=400, detail="File size exceeds 50MB limit")
        
        # Generate unique filename
        timestamp = int(time.time() * 1000)
        filename = f"{timestamp}_{file.filename}"
        file_path = os.path.join(Config.UPLOAD_DIR, filename)
        
        # Save file
        with open(file_path, "wb") as buffer:
            buffer.write(content)
        
        logger.info(f"File saved successfully: {filename} ({len(content)} bytes)")
        return filename
        
    except HTTPException:
        raise
    except Exception as error:
        logger.error(f"Error saving file {file.filename}: {error}")
        raise HTTPException(status_code=500, detail="Failed to upload file")

async def _process_document(file_id: str) -> ProcessingResult:
    """Process document independently"""
    try:
        file_path = os.path.join(Config.UPLOAD_DIR, file_id)
        if not os.path.exists(file_path):
            logger.error(f"File not found: {file_path}")
            raise Exception("File not found")

        logger.info(f"Processing document: {file_id}")
        
        # Extract text from Word document
        logger.info(f"Step 1: Extracting text from {file_id}")
        text = extract_text_from_word(file_path)
        if not text:
            logger.error(f"Failed to extract text from {file_id}")
            raise Exception("Failed to extract text from document")
        
        logger.info(f"Text extracted successfully from {file_id}: {len(text)} characters")

        # Preprocess text
        logger.info(f"Step 2: Preprocessing text from {file_id}")
        preprocessed_text = preprocess_text(text)
        if not preprocessed_text:
            logger.error(f"Failed to preprocess text from {file_id}")
            raise Exception("Failed to preprocess text")
        
        logger.info(f"Text preprocessed successfully from {file_id}: {len(preprocessed_text)} characters")

        # Extract market data using AI with retry mechanism
        logger.info(f"Step 3: AI extraction for {file_id} (attempt 1/2)")
        extracted_data = extract_market_data(preprocessed_text)
        
        # Validate extracted data structure
        if not extracted_data or not isinstance(extracted_data, dict):
            logger.warning(f"First AI extraction failed for {file_id}, retrying...")
            logger.info(f"Step 3: AI extraction for {file_id} (attempt 2/2)")
            extracted_data = extract_market_data(preprocessed_text)
        
        if not extracted_data or not isinstance(extracted_data, dict):
            logger.error(f"AI extraction failed for {file_id}: Invalid data structure")
            raise Exception("Failed to extract valid data structure from AI")
        
        logger.info(f"AI extraction successful for {file_id}: {list(extracted_data.keys())}")
        
        if 'market' not in extracted_data:
            logger.error(f"Missing 'market' key in extracted data for {file_id}")
            raise Exception("Missing 'market' key in data")

        # Normalize extracted data
        logger.info(f"Step 4: Normalizing data for {file_id}")
        normalized_data = normalize_extracted_data(extracted_data)
        if not normalized_data:
            logger.error(f"Failed to normalize extracted data for {file_id}")
            raise Exception("Failed to normalize extracted data")
        
        logger.info(f"Data normalization successful for {file_id}")

        # Validate extracted data
        logger.info(f"Step 5: Validating data for {file_id}")
        is_valid = validate_extracted_data(normalized_data)
        if not is_valid:
            logger.error(f"Data validation failed for {file_id}")
            raise Exception("Invalid data structure for Excel processing")
        
        logger.info(f"Data validation successful for {file_id}")

        # Calculate confidence score
        logger.info(f"Step 6: Calculating confidence for {file_id}")
        confidence = calculate_confidence(normalized_data)
        logger.info(f"Confidence score for {file_id}: {confidence}")

        # Convert Pydantic model to dictionary if needed
        data_for_response = normalized_data
        if hasattr(normalized_data, 'dict'):
            data_for_response = normalized_data.dict()
        elif hasattr(normalized_data, 'model_dump'):
            data_for_response = normalized_data.model_dump()

        logger.info(f"Document processing completed successfully for {file_id}")
        return ProcessingResult(
            success=True,
            data=data_for_response,
            confidence=confidence
        )

    except Exception as error:
        logger.error(f"Document processing error for {file_id}: {error}")
        logger.error(f"Error type: {type(error).__name__}")
        return ProcessingResult(
            success=False,
            data=None,
            confidence=0,
            error=str(error)
        )

async def _generate_excel(file_id: str, data: Dict[str, Any]) -> Dict[str, Any]:
    """Generate Excel file independently"""
    try:
        if not data:
            logger.error(f"No data provided for Excel generation for {file_id}")
            raise Exception("No data provided for Excel generation")

        template_path = Config.EXCEL_TEMPLATE_PATH
        logger.info(f'Using template path: {template_path} for {file_id}')

        # Validate data structure
        if not all(key in data for key in ['market', 'segments', 'players']):
            logger.error(f"Invalid data structure for Excel processing for {file_id}. Keys: {list(data.keys()) if hasattr(data, 'keys') else 'No keys method'}")
            raise Exception("Invalid data structure for Excel processing")

        logger.info(f"Data structure validation passed for {file_id}")

        # Use enhanced Excel processor for better macro preservation
        excel_processor = ExcelProcessorEnhanced(template_path)
        logger.info(f'Populating Excel template with extracted data for {file_id}...')
        excel_processor.populate_data(data)

        # Use market name for filename (same as D2 cell content)
        market_name = data.get('market', {}).get('market_name', 'Market Data')
        
        # Clean market name for filename (remove only problematic characters, keep spaces)
        import re
        clean_market_name = re.sub(r'[<>:"/\\|?*\x00]', '', market_name)  # Cross-platform filename sanitization
        clean_market_name = clean_market_name.strip()  # Remove leading/trailing whitespace
        
        # Fallback to timestamp if market name is empty
        if not clean_market_name:
            timestamp = int(time.time() * 1000)
            clean_market_name = f"Market Data {timestamp}"
        
        output_filename = f"{clean_market_name}.xlsm"
        
        # Ensure temp directory exists
        temp_dir = "temp"
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
        
        output_path = os.path.join(temp_dir, output_filename)

        logger.info(f'Saving populated Excel file with enhanced macro preservation for {file_id}...')
        excel_processor.save_file(output_path)

        # Clean up
        excel_processor.cleanup()

        logger.info(f'Excel file generated successfully for {file_id}: {output_filename}')

        return {
            'filename': output_filename,
            'path': output_path
        }

    except Exception as error:
        logger.error(f"Excel generation error for {file_id}: {error}")
        logger.error(f"Error type: {type(error).__name__}")
        logger.error(f"Data structure: {data}")
        raise error

@app.get("/api/download/{filename}")
async def download_file(filename: str):
    """Download generated Excel file and clean up after download"""
    try:
        # Ensure temp directory exists
        temp_dir = "temp"
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
        
        file_path = os.path.join(temp_dir, filename)
        
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="File not found")
        
        # Create a response that will clean up the file after download
        def cleanup_file():
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                    logger.info(f"Cleaned up downloaded Excel file: {filename}")
            except Exception as cleanup_error:
                logger.warning(f"Failed to cleanup Excel file {filename}: {cleanup_error}")
        
        # Use a custom response that triggers cleanup after download
        response = FileResponse(
            path=file_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename=\"{filename}\"",
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }
        )
        
        # Schedule cleanup after response is sent
        import asyncio
        asyncio.create_task(delayed_cleanup(file_path, filename))
        
        return response
        
    except HTTPException:
        raise
    except Exception as error:
        logger.error(f"Download error: {error}")
        raise HTTPException(status_code=500, detail="Failed to download file")

async def delayed_cleanup(file_path: str, filename: str):
    """Clean up file after a short delay to ensure download completes"""
    try:
        # Wait 5 seconds to ensure download completes
        await asyncio.sleep(5)
        
        if os.path.exists(file_path):
            os.remove(file_path)
            logger.info(f"Cleaned up downloaded Excel file: {filename}")
        else:
            logger.info(f"Excel file already cleaned up: {filename}")
            
    except Exception as cleanup_error:
        logger.warning(f"Failed to cleanup Excel file {filename}: {cleanup_error}")

@app.post("/api/check-generated-files")
async def check_generated_files(request: dict):
    """Check if files were generated in the temp directory"""
    try:
        temp_dir = "temp"
        if not os.path.exists(temp_dir):
            return {"files": []}
        
        # Get all .xlsm files in temp directory
        generated_files = []
        for filename in os.listdir(temp_dir):
            if filename.endswith('.xlsm'):
                generated_files.append(filename)
        
        return {"files": generated_files}
        
    except Exception as error:
        logger.error(f"Error checking generated files: {error}")
        return {"files": [], "error": str(error)}

@app.post("/api/cleanup-files")
async def cleanup_files():
    """Clean up orphaned files in uploads and temp directories"""
    try:
        cleanup_stats = {
            'uploads_cleaned': 0,
            'temp_cleaned': 0,
            'total_size_freed': 0
        }
        
        # Clean up uploads directory (Word files older than 1 hour)
        uploads_dir = Config.UPLOAD_DIR
        if os.path.exists(uploads_dir):
            current_time = time.time()
            for filename in os.listdir(uploads_dir):
                file_path = os.path.join(uploads_dir, filename)
                if os.path.isfile(file_path):
                    file_age = current_time - os.path.getmtime(file_path)
                    if file_age > 3600:  # 1 hour
                        file_size = os.path.getsize(file_path)
                        os.remove(file_path)
                        cleanup_stats['uploads_cleaned'] += 1
                        cleanup_stats['total_size_freed'] += file_size
                        logger.info(f"Cleaned up old uploaded file: {filename}")
        
        # Clean up temp directory (Excel files older than 1 hour)
        temp_dir = "temp"
        if os.path.exists(temp_dir):
            current_time = time.time()
            for filename in os.listdir(temp_dir):
                file_path = os.path.join(temp_dir, filename)
                if os.path.isfile(file_path):
                    file_age = current_time - os.path.getmtime(file_path)
                    if file_age > 3600:  # 1 hour
                        file_size = os.path.getsize(file_path)
                        os.remove(file_path)
                        cleanup_stats['temp_cleaned'] += 1
                        cleanup_stats['total_size_freed'] += file_size
                        logger.info(f"Cleaned up old Excel file: {filename}")
        
        logger.info(f"Cleanup completed: {cleanup_stats}")
        return {"success": True, "stats": cleanup_stats}
        
    except Exception as error:
        logger.error(f"Cleanup files error: {error}")
        return {"success": False, "error": str(error)}

@app.post("/api/cleanup-all-files")
async def cleanup_all_files():
    """Clean up ALL files in uploads and temp directories (immediate cleanup)"""
    try:
        cleanup_stats = {
            'uploads_cleaned': 0,
            'temp_cleaned': 0,
            'total_size_freed': 0
        }
        
        # Clean up ALL files in uploads directory
        uploads_dir = Config.UPLOAD_DIR
        if os.path.exists(uploads_dir):
            for filename in os.listdir(uploads_dir):
                file_path = os.path.join(uploads_dir, filename)
                if os.path.isfile(file_path):
                    file_size = os.path.getsize(file_path)
                    os.remove(file_path)
                    cleanup_stats['uploads_cleaned'] += 1
                    cleanup_stats['total_size_freed'] += file_size
                    logger.info(f"Cleaned up uploaded file: {filename}")
        
        # Clean up ALL files in temp directory
        temp_dir = "temp"
        if os.path.exists(temp_dir):
            for filename in os.listdir(temp_dir):
                file_path = os.path.join(temp_dir, filename)
                if os.path.isfile(file_path):
                    file_size = os.path.getsize(file_path)
                    os.remove(file_path)
                    cleanup_stats['temp_cleaned'] += 1
                    cleanup_stats['total_size_freed'] += file_size
                    logger.info(f"Cleaned up Excel file: {filename}")
        
        logger.info(f"Immediate cleanup completed: {cleanup_stats}")
        return {"success": True, "stats": cleanup_stats}
        
    except Exception as error:
        logger.error(f"Immediate cleanup error: {error}")
        return {"success": False, "error": str(error)}

async def startup_cleanup():
    """Clean up orphaned files on server startup"""
    try:
        logger.info("Starting cleanup of orphaned files...")
        
        cleanup_stats = {
            'uploads_cleaned': 0,
            'temp_cleaned': 0,
            'total_size_freed': 0
        }
        
        # Clean up uploads directory (Word files older than 30 minutes)
        uploads_dir = Config.UPLOAD_DIR
        if os.path.exists(uploads_dir):
            current_time = time.time()
            for filename in os.listdir(uploads_dir):
                file_path = os.path.join(uploads_dir, filename)
                if os.path.isfile(file_path):
                    file_age = current_time - os.path.getmtime(file_path)
                    if file_age > 1800:  # 30 minutes
                        file_size = os.path.getsize(file_path)
                        os.remove(file_path)
                        cleanup_stats['uploads_cleaned'] += 1
                        cleanup_stats['total_size_freed'] += file_size
                        logger.info(f"Startup cleanup - removed uploaded file: {filename}")
        
        # Clean up temp directory (Excel files older than 30 minutes)
        temp_dir = "temp"
        if os.path.exists(temp_dir):
            current_time = time.time()
            for filename in os.listdir(temp_dir):
                file_path = os.path.join(temp_dir, filename)
                if os.path.isfile(file_path):
                    file_age = current_time - os.path.getmtime(file_path)
                    if file_age > 1800:  # 30 minutes
                        file_size = os.path.getsize(file_path)
                        os.remove(file_path)
                        cleanup_stats['temp_cleaned'] += 1
                        cleanup_stats['total_size_freed'] += file_size
                        logger.info(f"Startup cleanup - removed Excel file: {filename}")
        
        logger.info(f"Startup cleanup completed: {cleanup_stats}")
        
    except Exception as error:
        logger.error(f"Startup cleanup error: {error}")

# Serve Next.js static build from FastAPI (for Render single-service deployment)
frontend_out_dir = os.path.join(os.path.dirname(__file__), "frontend", "out")
if os.path.exists(frontend_out_dir):
    from fastapi.responses import HTMLResponse

    # Serve static assets (_next, images, etc.)
    next_static_dir = os.path.join(frontend_out_dir, "_next")
    if os.path.exists(next_static_dir):
        app.mount("/_next", StaticFiles(directory=next_static_dir), name="next_static")

    # Catch-all route for frontend pages (must be after all /api routes)
    @app.get("/{full_path:path}")
    async def serve_frontend(full_path: str):
        # Try exact file first (e.g. favicon.ico)
        file_path = os.path.join(frontend_out_dir, full_path)
        if full_path and os.path.isfile(file_path):
            return FileResponse(file_path)
        # Fallback to index.html for SPA routing
        index_path = os.path.join(frontend_out_dir, "index.html")
        if os.path.exists(index_path):
            with open(index_path, "r") as f:
                return HTMLResponse(content=f.read())
        return HTMLResponse(content="Frontend not built. Run: cd frontend && npm run build", status_code=404)

if __name__ == "__main__":
    # Ensure required directories exist
    Config.ensure_directories()

    # Run startup cleanup
    import asyncio
    asyncio.run(startup_cleanup())

    uvicorn.run(
        "main:app",
        host=Config.HOST,
        port=Config.PORT,
        reload=True,
        reload_excludes=["temp/*", "uploads/*", "*.xlsm", "*.xlsx"],
        # Add connection keep-alive settings - 25 minutes
        timeout_keep_alive=1500,
        timeout_graceful_shutdown=1500
    )

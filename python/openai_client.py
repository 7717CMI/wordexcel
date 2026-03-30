import json
import logging
from typing import Dict, Any
from openai import OpenAI
from config import Config

logger = logging.getLogger(__name__)

# Initialize OpenAI client
client = OpenAI(api_key=Config.OPENAI_API_KEY)

# Prompt template for extracting market data from Word documents
EXTRACTION_PROMPT = """
You are an expert market research analyst. Extract the following information from the provided market research document:

REQUIRED FIELDS:
1. Market Name: The specific market/industry name
2. Base Year: The reference year for current market size (e.g., 2025)
3. Start Year: The beginning year of the forecast period (e.g., 2020)
4. End Year: The final year of the forecast period (e.g., 2032)
5. Market Size (Base Year): Current market size with currency and unit (e.g., "USD 150 Mn", "USD 0.15 Bn")
6. Market Size (End Year): Forecasted market size with currency and unit (e.g., "USD 290 Mn")
7. CAGR: Compound Annual Growth Rate as percentage (e.g., "9.50%")
8. Currency Unit: The currency used (e.g., "USD")

DRIVERS AND RESTRAINTS:
- Extract the TOP 2 market drivers (factors driving market growth) as SHORT OUTLINES ONLY (5-10 words max)
  - DO NOT write full sentences or explanations - just the key point as a brief outline
- Extract the TOP 2 market restraints/challenges (factors hindering market growth) - be concise, max 1-2 sentences each
- If more than 2 are mentioned, select the most important ones

SEGMENTATION DATA:
- Identify segmentation categories mentioned in the document (e.g., "By Technology", "By Application", "By End-User", "By Product Type")
- EXCLUDE any regional/geographical/country-based segmentation (e.g., "By Region", "By Country", "By Geography", etc.)
- For each category, list the specific segments/items
- If percentage shares are provided for segments, include them
- Use the exact category names as found in the document - don't rename them

KEY PLAYERS:
- Extract company names mentioned as key players, market leaders, or major competitors
- List them as individual company names

IMPORTANT: 
- Extract exactly what is stated in the document
- If a field is not found, use null or empty string
- Maintain the exact format and structure as found
- Convert all numbers to appropriate types
- Preserve currency and unit information as found
- Return ONLY the JSON object - no markdown, no code blocks, no explanations

Return the data in this JSON format (adapt the structure to match what you find):
{
  "market": {
    "market_name": "string",
    "base_year": number,
    "start_year": number,
    "end_year": number,
    "size_base_raw": "string",
    "size_forecast_raw": "string",
    "cagr_percent_display": "string",
    "currency_unit": "string",
    "driver_1": "string (SHORT outline only, 5-10 words max)",
    "driver_2": "string (SHORT outline only, 5-10 words max)",
    "restraint_1": "string (first key market restraint/challenge)",
    "restraint_2": "string (second key market restraint/challenge)"
  },
  "segments": {
    "category_name_1": {
      "header": "string (exact name from document)",
      "items": ["string"],
      "shares": [number] (if available)
    },
    "category_name_2": {
      "header": "string (exact name from document)",
      "items": ["string"]
    }
    // Add more categories as found in the document
  },
  "players": {
    "header": "Key Players",
    "players": ["string"]
  }
}

CRITICAL REQUIREMENTS:
- Extract exactly what is stated in the document
- If a field is not found, use null or empty string
- Maintain the exact format and structure
- Convert all numbers to appropriate types
- Preserve currency and unit information as found
- Return ONLY the JSON object - no markdown, no code blocks, no explanations
"""

def extract_market_data(document_text: str) -> Dict[str, Any]:
    """Extract data using OpenAI"""
    try:
        logger.info('=== AI EXTRACTION DEBUG ===')
        logger.info(f'Document text length: {len(document_text)}')
        logger.info(f'Document text preview (first 500 chars): {document_text[:500]}')
        logger.info(f'Document text preview (last 500 chars): {document_text[-500:]}')
        
        # Create chat completion
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": "You are a market research data extraction specialist. Extract data accurately and return valid JSON."
                },
                {
                    "role": "user",
                    "content": f"{EXTRACTION_PROMPT}\n\nDocument Text:\n{document_text}"
                }
            ],
            temperature=0.1,  # Low temperature for consistent extraction
            max_tokens=2000,
        )

        # Extract response content
        response_content = response.choices[0].message.content
        if not response_content:
            raise Exception("No response from OpenAI")

        logger.info('=== AI RESPONSE DEBUG ===')
        logger.info(f'Raw AI response: {response_content}')
        logger.info(f'Response length: {len(response_content)}')
        logger.info(f'Response preview: {response_content[:200]}...')

        # Try to parse the JSON response
        try:
            # Clean the response to extract JSON content
            json_content = response_content.strip()
            
            # Remove markdown code blocks if present
            if json_content.startswith('```json'):
                json_content = json_content.replace('```json', '', 1)
            if json_content.startswith('```'):
                json_content = json_content.replace('```', '', 1)
            if json_content.endswith('```'):
                json_content = json_content.replace('```', '', 1)
            
            # Clean up any remaining whitespace
            json_content = json_content.strip()
            
            logger.info(f'Cleaned JSON content: {json_content}')
            
            # Parse JSON
            extracted_data = json.loads(json_content)
            logger.info(f'Parsed data: {json.dumps(extracted_data, indent=2)}')
            
            return extracted_data
            
        except json.JSONDecodeError as parse_error:
            logger.error(f"Failed to parse OpenAI response: {parse_error}")
            logger.error(f"Raw response: {response_content}")
            raise Exception("Invalid response format from AI model")
            
    except Exception as error:
        logger.error(f"OpenAI API error: {error}")
        raise error

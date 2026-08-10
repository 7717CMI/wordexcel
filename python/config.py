import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# All relative paths resolve against this file's directory rather than the
# process working directory, so the app behaves identically whether it is
# started from the repo root, from python/, or from /app inside Docker.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _resolve(path: str) -> str:
    """Resolve a possibly-relative path against BASE_DIR."""
    return path if os.path.isabs(path) else os.path.normpath(os.path.join(BASE_DIR, path))


# Provider defaults. DeepSeek exposes an OpenAI-compatible API, so the same
# SDK drives both -- only the base URL, key and model name differ.
_PROVIDER_DEFAULTS = {
    'deepseek': {'base_url': 'https://api.deepseek.com', 'model': 'deepseek-chat'},
    'openai': {'base_url': None, 'model': 'gpt-4o-mini'},
}


class Config:
    # LLM Configuration
    LLM_PROVIDER = os.getenv('LLM_PROVIDER', 'deepseek').strip().lower()
    if LLM_PROVIDER not in _PROVIDER_DEFAULTS:
        raise ValueError(
            f"Unsupported LLM_PROVIDER {LLM_PROVIDER!r}; "
            f"expected one of {sorted(_PROVIDER_DEFAULTS)}"
        )

    DEEPSEEK_API_KEY = os.getenv('DEEPSEEK_API_KEY', '')
    OPENAI_API_KEY = os.getenv('OPENAI_API_KEY', '')

    # Resolved per-provider settings; each is individually overridable.
    LLM_API_KEY = DEEPSEEK_API_KEY if LLM_PROVIDER == 'deepseek' else OPENAI_API_KEY
    LLM_BASE_URL = os.getenv('LLM_BASE_URL') or _PROVIDER_DEFAULTS[LLM_PROVIDER]['base_url']
    LLM_MODEL = os.getenv('LLM_MODEL') or _PROVIDER_DEFAULTS[LLM_PROVIDER]['model']
    LLM_MAX_TOKENS = int(os.getenv('LLM_MAX_TOKENS', 4000))
    LLM_TIMEOUT = float(os.getenv('LLM_TIMEOUT', 120))

    # Name of the env var the user needs to set, used in error messages.
    LLM_KEY_ENV_VAR = 'DEEPSEEK_API_KEY' if LLM_PROVIDER == 'deepseek' else 'OPENAI_API_KEY'

    # Cap the document text sent to the model so a large report cannot exceed
    # the context window (roughly 4 chars per token).
    MAX_DOCUMENT_CHARS = int(os.getenv('MAX_DOCUMENT_CHARS', 200000))

    # File Upload Configuration
    MAX_FILE_SIZE = int(os.getenv('MAX_FILE_SIZE', 52428800))  # 50MB default
    MAX_BULK_FILES = int(os.getenv('MAX_BULK_FILES', 50))
    UPLOAD_DIR = _resolve(os.getenv('UPLOAD_DIR', './uploads'))
    TEMP_DIR = _resolve(os.getenv('TEMP_DIR', './temp'))

    # How long the bulk processor waits for the browser to confirm a download
    # before moving on to the next file.
    DOWNLOAD_CONFIRM_TIMEOUT = int(os.getenv('DOWNLOAD_CONFIRM_TIMEOUT', 30))

    # How long a generated Excel file stays on disk after being served.
    DOWNLOAD_RETENTION_SECONDS = int(os.getenv('DOWNLOAD_RETENTION_SECONDS', 300))

    # Excel Template Configuration
    EXCEL_TEMPLATE_PATH = _resolve(
        os.getenv('EXCEL_TEMPLATE_PATH', './assets/Bonding_Neodymium_Magnet_Market.xlsm')
    )

    # CORS: comma-separated origins. The frontend is served from the same
    # origin in the default deployment, so no cross-origin access is needed.
    CORS_ORIGINS = [
        o.strip() for o in os.getenv('CORS_ORIGINS', '').split(',') if o.strip()
    ]

    # Server Configuration
    HOST = os.getenv('HOST', '0.0.0.0')
    PORT = int(os.getenv('PORT', 8000))

    # Ensure upload directory exists
    @classmethod
    def ensure_directories(cls):
        """Ensure required directories exist"""
        os.makedirs(cls.UPLOAD_DIR, exist_ok=True)
        os.makedirs(cls.TEMP_DIR, exist_ok=True)

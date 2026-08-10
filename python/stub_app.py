"""Run the real app with the OpenAI call stubbed out.

Used to exercise upload -> process -> generate-excel -> download end to end
without spending API credits. Not imported by the application itself.

    uvicorn stub_app:app
"""
import main

# Shape mirrors what gpt-4o-mini returns for the extraction prompt, including
# the category_name_* segment keys that normalize_extracted_data remaps.
STUB_RESPONSE = {
    "market": {
        "market_name": "Bonding Neodymium Magnet Market",
        "base_year": 2025,
        "start_year": 2020,
        "end_year": 2032,
        "size_base_raw": "USD 150 Mn",
        "size_forecast_raw": "USD 290 Mn",
        "cagr_percent_display": "9.50%",
        "currency_unit": "USD",
        "driver_1": "Rising EV traction motor demand",
        "driver_2": "Consumer electronics miniaturization",
        "restraint_1": "Volatile rare earth material prices.",
        "restraint_2": "Supply chain concentration in China.",
    },
    "segments": {
        "By Technology": {"header": "By Technology", "items": ["Compression Bonded", "Injection Bonded"]},
        "By Application": {
            "header": "By Application",
            "items": ["Automotive", "Consumer Electronics", "Industrial"],
        },
    },
    "players": {
        "header": "Key Players",
        "players": ["Daido Steel Co Ltd", "TDK Corporation", "Nichia Corporation", "Galaxy Magnet"],
    },
}


def _stub_extract(document_text: str):
    import copy
    return copy.deepcopy(STUB_RESPONSE)


main.extract_market_data = _stub_extract
app = main.app

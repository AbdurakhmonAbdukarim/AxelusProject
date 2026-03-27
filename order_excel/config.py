# =============================================================================
# CONFIG.PY  —  Central configuration. Edit values here to change behavior.
# =============================================================================

# ---------------------------------------------------------------------------
# Telegram Bot
# ---------------------------------------------------------------------------
TELEGRAM_BOT_TOKEN = "8772012180:AAGIbKsekXz1RgiGR45VHcutKtt4rcSgRFM"   # Get from @BotFather on Telegram

# ---------------------------------------------------------------------------
# Zoho API credentials
# ---------------------------------------------------------------------------
ZOHO_ACCESS_TOKEN  = "1000.3374525a75dcbdd86115a4b3a0dcb58b.55d85f6473fab51ab8e87ec04ad7b6c4"   # Refreshed automatically
ZOHO_REFRESH_TOKEN = "1000.2816b922b5aa02f23ccbfe575774ea6c.61a053c1726f7d5d6a842b1be2f091bc"  # Never expires — from get_zoho_token.py
ZOHO_CLIENT_ID     = "1000.HKSTC4EOA2K294N9605P78PTSX96YQ"      # From api-console.zoho.com
ZOHO_CLIENT_SECRET = "220f4a58e520e65016f3884baf19678c2dafe8fd95"  # From api-console.zoho.com
ZOHO_ORG_ID        = "839172316"         # Zoho organization ID
ZOHO_BASE_URL      = "https://www.zohoapis.com/inventory/v1"



# ---------------------------------------------------------------------------
# File paths
# ---------------------------------------------------------------------------
INPUT_FILE  = "orders.xlsx"     # Source Excel file
OUTPUT_DIR  = "./output"        # Folder where per-manufacturer files are saved

# ---------------------------------------------------------------------------
# Source Excel column names  (change if your headers differ)
# ---------------------------------------------------------------------------
COL_ITEM_NAME    = "Item Name"
COL_SKU          = "SKU"
COL_MANUFACTURER = "Manufacturer"
COL_ORDER_QTY    = "ZAKAZ SONI"

# ---------------------------------------------------------------------------
# Output Excel layout — column headers (order matters)
# ---------------------------------------------------------------------------
OUTPUT_HEADERS = [
    "ITEM PHOTO",
    "ITEM NAME",
    "ITEM MODEL",
    "TOTAL CTNS",
    "PCS/CTN",
    "TOTAL QTY",
    "UNIT PRICE",
    "TOTAL AMOUNT PRICE",
    "CBM",
    "TOTAL CBM",
    "GROSS WEIGHT",
    "NET WEIGHT",
    "BARCODE",        # barcode image (scannable)
    "SKU",            # SKU text value stored here
    "SHIPPING MARK",
]

# Column widths (same order as OUTPUT_HEADERS)
OUTPUT_COL_WIDTHS = [18, 22, 20, 11, 9, 10, 11, 16, 8, 10, 13, 12, 32, 15, 18]

# ---------------------------------------------------------------------------
# Row / image sizing
# ---------------------------------------------------------------------------
ROW_HEIGHT_PT      = 100    # Data row height in points
PHOTO_WIDTH_PX     = 100    # Item photo width in Excel
PHOTO_HEIGHT_PX    = 80    # Item photo height in Excel
BARCODE_WIDTH_PX   = 280   # Barcode image width in Excel
BARCODE_HEIGHT_PX  = 80    # Barcode image height in Excel

# ---------------------------------------------------------------------------
# Barcode settings  (python-barcode Code128)
# ---------------------------------------------------------------------------
BARCODE_MODULE_WIDTH  = 0.4
BARCODE_MODULE_HEIGHT = 12
BARCODE_QUIET_ZONE    = 2
BARCODE_FONT_SIZE     = 7
BARCODE_TEXT_DISTANCE = 2

# ---------------------------------------------------------------------------
# Shipping mark template  — {sku} is replaced with the item SKU
# ---------------------------------------------------------------------------
SHIPPING_MARK_TEMPLATE = "AN-203\n{sku}"
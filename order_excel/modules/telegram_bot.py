# =============================================================================
# modules/telegram_bot.py
# =============================================================================

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import asyncio
import logging
import tempfile

from telegram import Update, Message
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

import config
from modules.reader       import load_source
from modules.zoho         import load_cache, get_item_photo, get_units_per_box
from modules.barcodes     import make_barcode
from modules.excel_writer import save_manufacturer_workbook

logging.basicConfig(format="%(asctime)s [%(levelname)s] %(name)s: %(message)s", level=logging.INFO)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Live status message — edits one Telegram message in-place
# ---------------------------------------------------------------------------

class LiveLog:
    def __init__(self, message: Message):
        self._message = message
        self._header  = "📋 <b>Progress</b>\n\n"
        self._lines   = []
        self._text    = ""

    async def set(self, line: str):
        if self._lines:
            self._lines[-1] = line
        else:
            self._lines.append(line)
        await self._flush()

    async def add(self, line: str):
        self._lines.append(line)
        await self._flush()

    async def _flush(self):
        new_text = self._header + "\n".join(self._lines)
        if new_text == self._text:
            return
        self._text = new_text
        try:
            await self._message.edit_text(new_text, parse_mode="HTML")
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Hello! I process order Excel files.\n\n"
        "📎 Send me your <b>.xlsx</b> file and I will:\n"
        "  • Filter rows with a ZAKAZ SONI quantity\n"
        "  • Group them by Manufacturer\n"
        "  • Fetch item photos from Zoho Inventory\n"
        "  • Generate barcodes for each item\n"
        "  • Send back one Excel file per manufacturer",
        parse_mode="HTML",
    )


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📋 <b>Required columns in your .xlsx file:</b>\n\n"
        f"  • <code>{config.COL_ITEM_NAME}</code>\n"
        f"  • <code>{config.COL_SKU}</code>\n"
        f"  • <code>{config.COL_MANUFACTURER}</code>\n"
        f"  • <code>{config.COL_ORDER_QTY}</code>",
        parse_mode="HTML",
    )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _clean_sku(raw) -> str:
    """Clean SKU string — strip .0 but preserve leading zeros."""
    s = str(raw).strip() if raw is not None else ""
    if s.endswith(".0") and not s.startswith("0"):
        try:
            return str(int(float(s)))
        except (ValueError, TypeError):
            pass
    return s


def _item_in_zoho(sku: str, item_name: str) -> bool:
    """Check if item exists in Zoho cache by SKU or name."""
    from modules.zoho import _cache_by_sku, _cache_by_name, _clean
    clean_sku  = _clean(sku)
    clean_name = _clean(item_name)
    return (
        (bool(clean_sku)  and clean_sku  in _cache_by_sku)  or
        (bool(clean_name) and clean_name in _cache_by_name)
    )


# ---------------------------------------------------------------------------
# Main document handler
# ---------------------------------------------------------------------------

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document

    if not doc.file_name.endswith(".xlsx"):
        await update.message.reply_text("⚠️ Please send an <b>.xlsx</b> file.", parse_mode="HTML")
        return

    status_msg = await update.message.reply_text("⏳ Starting...", parse_mode="HTML")
    live = LiveLog(status_msg)

    with tempfile.TemporaryDirectory() as tmp_dir:

        # ── Step 1: Download ─────────────────────────────────────────────
        await live.add(f"⬇️ Downloading <code>{doc.file_name}</code>...")
        input_path = os.path.join(tmp_dir, doc.file_name)
        tg_file    = await context.bot.get_file(doc.file_id)
        await tg_file.download_to_drive(input_path)
        await live.set(f"✅ Downloaded: <code>{doc.file_name}</code>")

        # ── Step 2: Read Excel ───────────────────────────────────────────
        await live.add("📖 Reading Excel...")
        try:
            df = load_source(input_path)
        except SystemExit as e:
            await live.set(f"❌ Error reading file:\n<code>{e}</code>")
            return

        manufacturers = df[config.COL_MANUFACTURER].unique().tolist()
        await live.set(
            f"✅ Excel read: <b>{len(df)}</b> rows, "
            f"<b>{len(manufacturers)}</b> manufacturer(s)"
        )

        use_zoho = bool(
            config.ZOHO_ACCESS_TOKEN and
            config.ZOHO_ACCESS_TOKEN != "YOUR_ACCESS_TOKEN_HERE" and
            config.ZOHO_ORG_ID and
            config.ZOHO_ORG_ID != "YOUR_ORG_ID_HERE"
        )

        # ── Step 3: Load Zoho cache ──────────────────────────────────────
        if use_zoho:
            await live.add("🔄 Connecting to Zoho Inventory...")
            zoho_log_lines = []

            async def zoho_log(msg: str):
                zoho_log_lines.append(msg)
                visible = zoho_log_lines[-3:]
                await live.set("\n   ".join(["🔄 Loading Zoho catalog..."] + visible))

            await load_cache(log_fn=zoho_log)
            await live.set(
                f"✅ Zoho catalog loaded "
                f"({zoho_log_lines[-1] if zoho_log_lines else 'done'})"
            )
        else:
            await live.add("⚪️ Zoho skipped (no token configured)")

        # ── Step 4: Process each manufacturer ───────────────────────────
        output_dir = os.path.join(tmp_dir, "output")
        os.makedirs(output_dir, exist_ok=True)
        original_output_dir = config.OUTPUT_DIR
        config.OUTPUT_DIR   = output_dir
        output_files        = []

        for mfr_idx, (manufacturer, group) in enumerate(df.groupby(config.COL_MANUFACTURER), 1):

            # Build rows_data: list of (key, sku, item_name)
            rows_data = []
            for _, row in group.iterrows():
                raw_sku   = row.get(config.COL_SKU, "")
                item_name = str(row.get(config.COL_ITEM_NAME, ""))
                sku       = _clean_sku(raw_sku)
                key       = sku if sku.lower() not in ("nan", "", "none") else item_name
                rows_data.append((key, sku, item_name))

            # Dedupe by key
            seen = {}
            for key, sku, item_name in rows_data:
                if key not in seen:
                    seen[key] = (sku, item_name)
            unique_rows = [(k, s, n) for k, (s, n) in seen.items()]

            await live.add(
                f"🏭 <b>[{mfr_idx}/{len(manufacturers)}] {manufacturer}</b> "
                f"— {len(unique_rows)} item(s)"
            )

            # ── Fetch photos + units_per_box ─────────────────────────────
            photo_map         = {}
            units_per_box_map = {}
            barcode_map       = {}

            if use_zoho:
                await live.set(
                    f"🏭 <b>[{mfr_idx}/{len(manufacturers)}] {manufacturer}</b>\n"
                    f"   ⚡ Fetching {len(unique_rows)} photos & box units..."
                )

                # Photos first, then UPB — sequential to avoid 429
                photo_results = await asyncio.gather(
                    *[get_item_photo(sku, item_name) for _, sku, item_name in unique_rows],
                    return_exceptions=True
                )
                upb_results = await asyncio.gather(
                    *[get_units_per_box(sku, item_name) for _, sku, item_name in unique_rows],
                    return_exceptions=True
                )

                found = 0
                for (key, sku, item_name), photo, upb in zip(unique_rows, photo_results, upb_results):
                    photo_map[key]         = None if isinstance(photo, Exception) else photo
                    units_per_box_map[key] = None if isinstance(upb,   Exception) else upb
                    if photo_map[key]:
                        found += 1

                await live.set(
                    f"🏭 <b>[{mfr_idx}/{len(manufacturers)}] {manufacturer}</b>\n"
                    f"   ✅ Photos: {found}/{len(unique_rows)} found"
                )
            else:
                await live.set(
                    f"🏭 <b>[{mfr_idx}/{len(manufacturers)}] {manufacturer}</b>\n"
                    f"   ⚪️ Photos skipped"
                )

            # ── Barcodes ─────────────────────────────────────────────────
            await live.set(
                f"🏭 <b>[{mfr_idx}/{len(manufacturers)}] {manufacturer}</b>\n"
                f"   🔲 Generating {len(unique_rows)} barcode(s)..."
            )
            loop = asyncio.get_event_loop()
            bc_results = await asyncio.gather(
                *[loop.run_in_executor(None, make_barcode, sku, item_name)
                  for key, sku, item_name in rows_data],
                return_exceptions=True
            )
            for (key, sku, item_name), bc in zip(rows_data, bc_results):
                barcode_map[key] = bc if not isinstance(bc, Exception) else None

            # ── Write Excel ───────────────────────────────────────────────
            await live.set(
                f"🏭 <b>[{mfr_idx}/{len(manufacturers)}] {manufacturer}</b>\n"
                f"   📝 Writing Excel file..."
            )
            out_path = save_manufacturer_workbook(
                manufacturer, group, photo_map, barcode_map, units_per_box_map
            )

            # ── Build item report ─────────────────────────────────────────
            missing_lines = []
            if use_zoho:
                for key, sku, item_name in unique_rows:
                    in_zoho  = _item_in_zoho(sku, item_name)
                    no_photo = photo_map.get(key) is None
                    no_upb   = units_per_box_map.get(key) is None
                    label    = (sku + " " if sku.lower() not in ("nan","","none") else "") + item_name

                    if not in_zoho:
                        missing_lines.append("- " + label + " : not found in Zoho")
                    else:
                        issues = []
                        if no_photo: issues.append("no photo")
                        if no_upb:   issues.append("no pcs/box")
                        if issues:
                            missing_lines.append("! " + label + " : " + ", ".join(issues))

            if missing_lines:
                report = "<b>" + manufacturer + "</b> - missing:\n" + "\n".join(missing_lines)
            else:
                report = "<b>" + manufacturer + "</b> - all items complete"

            output_files.append((manufacturer, out_path, report))

            await live.set(
                f"✅ <b>[{mfr_idx}/{len(manufacturers)}] {manufacturer}</b> — done"
            )

        config.OUTPUT_DIR = original_output_dir

        # ── Step 5: Send files ───────────────────────────────────────────
        if not output_files:
            await live.add("⚠️ No output files were generated.")
            return

        await live.add(f"📤 Sending {len(output_files)} file(s)...")

        for i, (manufacturer, path, report) in enumerate(output_files, 1):
            await live.set(f"📤 Sending {i}/{len(output_files)}: <b>{manufacturer}</b>...")
            with open(path, "rb") as f:
                await context.bot.send_document(
                    chat_id=update.effective_chat.id,
                    document=f,
                    filename=os.path.basename(path),
                    caption="📦 " + manufacturer,
                )
            # Send item report after each file
            await context.bot.send_message(
                chat_id=update.effective_chat.id,
                text=report,
                parse_mode="HTML",
            )

        await live.set(f"✅ Sent {len(output_files)} file(s)")
        await live.add("🎉 <b>All done!</b>")


# ---------------------------------------------------------------------------
# Unknown message
# ---------------------------------------------------------------------------

async def handle_unknown(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📎 Send me an .xlsx file, or use /help.")


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------

def run_bot():
    token = config.TELEGRAM_BOT_TOKEN
    if not token or token == "YOUR_BOT_TOKEN_HERE":
        raise ValueError("TELEGRAM_BOT_TOKEN is not set in config.py!")

    app = ApplicationBuilder().token(token).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help",  cmd_help))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_unknown))

    log.info("Bot is running... Press Ctrl+C to stop.")
    app.run_polling()
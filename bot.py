"""
Merchant Turnover Telegram Bot
================================
Accepts GPay Business & Paytm merchant CSVs, compiles them into a
colour-coded Excel report with daily turnovers and a grand total.

Commands
--------
/start   – welcome & instructions
/status  – show what has been collected so far
/compile – generate and send the Excel report
/reset   – clear all data for a fresh run
"""

import io
import logging
import os
from collections import defaultdict
from datetime import datetime

import pandas as pd
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN", "")

logging.basicConfig(
    format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
#  CSV DETECTION & PARSING
# ══════════════════════════════════════════════════════════════════════════════

def _clean(val: str) -> str:
    """Strip whitespace and stray single/double quotes from a cell value."""
    return str(val).strip().strip("'\"")


def detect_csv_type(df: pd.DataFrame) -> str:
    """Return 'gpay', 'paytm', or 'unknown' based on column signatures."""
    cols = {c.strip().lower() for c in df.columns}
    if "transaction_date" in cols or "merchant_name" in cols:
        return "paytm"
    if "creation time" in cols or "paid via" in cols:
        return "gpay"
    return "unknown"


def parse_gpay_csv(df: pd.DataFrame, shop_name: str) -> list[dict]:
    """
    GPay Business export columns we care about:
        Creation time | Type | Amount | Status

    Rules:
    • Status  == 'Settled'
    • Type    == 'UPI'        (skip 'Daily collections' = SoundPod fee)
    • Amount  >  0            (skip fee deductions which are negative)
    """
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
    mask = (
        (df["Status"].str.strip() == "Settled")
        & (df["Type"].str.strip() == "UPI")
        & (df["Amount"] > 0)
    )
    df = df[mask]

    records: list[dict] = []
    for _, row in df.iterrows():
        try:
            date = pd.to_datetime(_clean(row["Creation time"])).date()
            records.append(
                {
                    "shop": shop_name,
                    "merchant": "GPay",
                    "date": date,
                    "amount": float(row["Amount"]),
                    "payer": _clean(row.get("Payer/Receiver", "")),
                    "txn_id": _clean(row.get("Transaction ID", "")),
                }
            )
        except Exception as exc:
            logger.warning("Skipping GPay row – %s", exc)

    return records


def parse_paytm_csv(df: pd.DataFrame) -> list[dict]:
    """
    Paytm merchant export columns we care about:
        Transaction_Date | Merchant_Name | Amount | Status

    Rules:
    • Status == 'SUCCESS'
    • Amount >  0

    Note: Paytm exports often wrap values in single quotes;
    _clean() strips them.
    """
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    # Status may be wrapped in quotes
    df["_status"] = df["Status"].apply(_clean)
    df = df[df["_status"] == "SUCCESS"]

    df["Amount"] = pd.to_numeric(
        df["Amount"].apply(lambda x: _clean(x).replace(",", "")),
        errors="coerce",
    )
    df = df[df["Amount"] > 0]

    records: list[dict] = []
    for _, row in df.iterrows():
        try:
            date = pd.to_datetime(_clean(row["Transaction_Date"])).date()
            shop = _clean(row.get("Merchant_Name", "Unknown"))
            records.append(
                {
                    "shop": shop,
                    "merchant": "Paytm",
                    "date": date,
                    "amount": float(row["Amount"]),
                    "payer": _clean(row.get("Customer_VPA", "")),
                    "txn_id": _clean(row.get("Transaction_ID", "")),
                }
            )
        except Exception as exc:
            logger.warning("Skipping Paytm row – %s", exc)

    return records


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL REPORT GENERATION
# ══════════════════════════════════════════════════════════════════════════════

# Colour palette
C = {
    "header_bg":      "263238",  # dark blue-grey  (column headers)
    "title_bg":       "1A237E",  # deep indigo     (sheet title)
    "gpay_row":       "E3F2FD",  # light blue      (GPay data rows)
    "gpay_accent":    "1565C0",  # dark blue       (GPay label / font)
    "paytm_row":      "E8F5E9",  # light green     (Paytm data rows)
    "paytm_accent":   "2E7D32",  # dark green      (Paytm label / font)
    "subtotal_bg":    "FFF9C4",  # light yellow    (per-shop subtotal)
    "subtotal_font":  "E65100",  # deep orange     (subtotal label)
    "grand_bg":       "BF360C",  # deep orange-red (grand total)
    "white":          "FFFFFF",
    "grey_border":    "B0BEC5",
}

_thin  = Side(style="thin",   color=C["grey_border"])
_thick = Side(style="medium", color="78909C")


def _border(top=_thin, bottom=_thin, left=_thin, right=_thin):
    return Border(top=top, bottom=bottom, left=left, right=right)


def _hdr(ws, row: int, col: int, value, bg: str = "header_bg",
         fg: str = "white", size: int = 10, bold: bool = True,
         h_align: str = "center") -> None:
    cell = ws.cell(row=row, column=col, value=value)
    cell.font      = Font(bold=bold, color=C[fg], size=size, name="Calibri")
    cell.fill      = PatternFill("solid", fgColor=C[bg])
    cell.alignment = Alignment(horizontal=h_align, vertical="center",
                                wrap_text=True)


def _data(ws, row: int, col: int, value, bg: str, bold: bool = False,
          fmt: str | None = None, fg: str = "header_bg",
          h_align: str = "center") -> None:
    cell = ws.cell(row=row, column=col, value=value)
    cell.font      = Font(bold=bold, color=C[fg], name="Calibri", size=10)
    cell.fill      = PatternFill("solid", fgColor=C[bg])
    cell.alignment = Alignment(horizontal=h_align, vertical="center")
    if fmt:
        cell.number_format = fmt


def _set_widths(ws, widths: dict[str, float]) -> None:
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w


def _freeze(ws, cell: str) -> None:
    ws.freeze_panes = cell


# ── Sheet helpers ──────────────────────────────────────────────────────────────

def _build_dashboard(wb: Workbook, shop_daily: dict, shop_txns: dict) -> None:
    ws = wb.active
    ws.title = "Dashboard"
    ws.sheet_view.showGridLines = False

    # Title
    ws.merge_cells("A1:F1")
    cell = ws["A1"]
    cell.value = (
        f"  Merchant Daily Turnover Report  •  "
        f"Generated {datetime.now().strftime('%d %b %Y  %H:%M')}"
    )
    cell.font      = Font(bold=True, size=13, color=C["white"], name="Calibri")
    cell.fill      = PatternFill("solid", fgColor=C["title_bg"])
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 32

    # Column headers
    col_labels = ["Shop Name", "Merchant", "Date", "Transactions", "Daily Turnover (₹)", ""]
    for c, lbl in enumerate(col_labels, 1):
        _hdr(ws, 2, c, lbl)
        ws.cell(row=2, column=c).border = _border(
            top=Side(style="medium", color="455A64"),
            bottom=Side(style="medium", color="455A64"),
        )
    ws.row_dimensions[2].height = 22

    row = 3
    grand_total = 0.0

    for (shop, merchant), daily in sorted(shop_daily.items(),
                                          key=lambda x: x[0][0]):
        row_bg     = "gpay_row"    if merchant == "GPay"   else "paytm_row"
        accent_key = "gpay_accent" if merchant == "GPay"   else "paytm_accent"
        accent_hex = C[accent_key]

        shop_total  = sum(daily.values())
        grand_total += shop_total
        txns_total  = len(shop_txns[(shop, merchant)])
        first_row   = row

        for date in sorted(daily.keys()):
            day_txns = sum(
                1 for r in shop_txns[(shop, merchant)] if r["date"] == date
            )
            day_amt  = daily[date]
            is_first = (row == first_row)

            # Shop name only in the first row of this group
            name_cell = ws.cell(row=row, column=1,
                                value=shop if is_first else "")
            name_cell.fill      = PatternFill("solid", fgColor=C[row_bg])
            name_cell.font      = Font(bold=is_first, color=accent_hex,
                                       name="Calibri", size=10)
            name_cell.alignment = Alignment(horizontal="left",
                                             vertical="center",
                                             indent=1)

            # Merchant
            mc = ws.cell(row=row, column=2, value=merchant if is_first else "")
            mc.fill      = PatternFill("solid", fgColor=C[row_bg])
            mc.font      = Font(bold=True, color=accent_hex,
                                 name="Calibri", size=10)
            mc.alignment = Alignment(horizontal="center", vertical="center")

            _data(ws, row, 3, date.strftime("%d %b %Y"), row_bg,
                  h_align="center")
            _data(ws, row, 4, day_txns,  row_bg, h_align="center")
            _data(ws, row, 5, day_amt,   row_bg, fmt="#,##0.00",
                  h_align="right")

            for c in range(1, 6):
                ws.cell(row=row, column=c).border = _border()
            ws.row_dimensions[row].height = 18
            row += 1

        # Per-shop subtotal
        ws.merge_cells(f"A{row}:B{row}")
        sub = ws.cell(row=row, column=1,
                      value=f"  {shop} — Subtotal")
        sub.font      = Font(bold=True, color=C["subtotal_font"],
                              name="Calibri", size=10)
        sub.fill      = PatternFill("solid", fgColor=C["subtotal_bg"])
        sub.alignment = Alignment(horizontal="left", vertical="center",
                                   indent=1)

        _data(ws, row, 3, "", "subtotal_bg")
        _data(ws, row, 4, txns_total, "subtotal_bg", bold=True,
              fg="subtotal_font", h_align="center")
        _data(ws, row, 5, shop_total, "subtotal_bg", bold=True,
              fg="subtotal_font", fmt="#,##0.00", h_align="right")

        for c in range(1, 6):
            ws.cell(row=row, column=c).border = _border(
                top=Side(style="medium", color="E65100"),
                bottom=Side(style="medium", color="E65100"),
            )
        ws.row_dimensions[row].height = 20
        row += 1

    # Grand total
    ws.merge_cells(f"A{row}:D{row}")
    gt = ws.cell(row=row, column=1, value="  GRAND TOTAL — All Shops")
    gt.font      = Font(bold=True, size=12, color=C["white"], name="Calibri")
    gt.fill      = PatternFill("solid", fgColor=C["grand_bg"])
    gt.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    for c in range(2, 5):
        ws.cell(row=row, column=c).fill = PatternFill("solid",
                                                       fgColor=C["grand_bg"])

    gv = ws.cell(row=row, column=5, value=grand_total)
    gv.font           = Font(bold=True, size=12, color=C["white"],
                              name="Calibri")
    gv.fill           = PatternFill("solid", fgColor=C["grand_bg"])
    gv.number_format  = "#,##0.00"
    gv.alignment      = Alignment(horizontal="right", vertical="center")

    for c in range(1, 6):
        ws.cell(row=row, column=c).border = _border(
            top=Side(style="thick", color="BF360C"),
            bottom=Side(style="thick", color="BF360C"),
        )
    ws.row_dimensions[row].height = 26

    _set_widths(ws, {"A": 24, "B": 12, "C": 16, "D": 16, "E": 22, "F": 4})
    _freeze(ws, "A3")


def _build_shop_summary(wb: Workbook, shop_daily: dict, shop_txns: dict) -> None:
    ws = wb.create_sheet("Shop Summary")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:H1")
    cell = ws["A1"]
    cell.value     = "  Shop-wise Summary"
    cell.font      = Font(bold=True, size=13, color=C["white"], name="Calibri")
    cell.fill      = PatternFill("solid", fgColor=C["title_bg"])
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30

    hdrs = ["#", "Shop Name", "Merchant", "Total Transactions",
            "Total Turnover (₹)", "Date From", "Date To", ""]
    for c, h in enumerate(hdrs, 1):
        _hdr(ws, 2, c, h)
    ws.row_dimensions[2].height = 22

    idx = 1
    grand = 0.0
    for (shop, merchant), daily in sorted(shop_daily.items(),
                                           key=lambda x: x[0][0]):
        row_bg     = "gpay_row"    if merchant == "GPay"   else "paytm_row"
        accent_key = "gpay_accent" if merchant == "GPay"   else "paytm_accent"

        total      = sum(daily.values())
        grand     += total
        txns       = len(shop_txns[(shop, merchant)])
        dates      = sorted(daily.keys())
        r          = idx + 2

        _data(ws, r, 1, idx,    row_bg, h_align="center")
        _data(ws, r, 2, shop,   row_bg, bold=True, fg=accent_key,
              h_align="left")
        _data(ws, r, 3, merchant, row_bg, bold=True, fg=accent_key)
        _data(ws, r, 4, txns,   row_bg, h_align="center")
        _data(ws, r, 5, total,  row_bg, bold=True, fmt="#,##0.00",
              h_align="right")
        _data(ws, r, 6, dates[0].strftime("%d %b %Y"), row_bg)
        _data(ws, r, 7, dates[-1].strftime("%d %b %Y"), row_bg)

        for c in range(1, 8):
            ws.cell(row=r, column=c).border = _border()
        ws.row_dimensions[r].height = 20
        idx += 1

    gt_row = idx + 2
    ws.merge_cells(f"A{gt_row}:D{gt_row}")
    gt = ws.cell(row=gt_row, column=1, value="  GRAND TOTAL")
    gt.font      = Font(bold=True, size=12, color=C["white"], name="Calibri")
    gt.fill      = PatternFill("solid", fgColor=C["grand_bg"])
    gt.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for c in range(2, 5):
        ws.cell(row=gt_row, column=c).fill = PatternFill(
            "solid", fgColor=C["grand_bg"])

    gv = ws.cell(row=gt_row, column=5, value=grand)
    gv.font          = Font(bold=True, size=12, color=C["white"], name="Calibri")
    gv.fill          = PatternFill("solid", fgColor=C["grand_bg"])
    gv.number_format = "#,##0.00"
    gv.alignment     = Alignment(horizontal="right", vertical="center")

    for c in range(6, 8):
        ws.cell(row=gt_row, column=c).fill = PatternFill(
            "solid", fgColor=C["grand_bg"])
    for c in range(1, 8):
        ws.cell(row=gt_row, column=c).border = _border(
            top=Side(style="thick", color="BF360C"),
            bottom=Side(style="thick", color="BF360C"),
        )
    ws.row_dimensions[gt_row].height = 26

    _set_widths(ws, {"A": 5, "B": 24, "C": 12, "D": 20,
                      "E": 22, "F": 14, "G": 14, "H": 4})
    _freeze(ws, "A3")


def _build_all_transactions(wb: Workbook, all_records: list[dict]) -> None:
    ws = wb.create_sheet("All Transactions")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:H1")
    cell = ws["A1"]
    cell.value     = "  All Transactions"
    cell.font      = Font(bold=True, size=13, color=C["white"], name="Calibri")
    cell.fill      = PatternFill("solid", fgColor=C["title_bg"])
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30

    hdrs = ["#", "Shop Name", "Merchant", "Date", "Payer / VPA",
            "Transaction ID", "Amount (₹)", ""]
    for c, h in enumerate(hdrs, 1):
        _hdr(ws, 2, c, h)
    ws.row_dimensions[2].height = 22

    sorted_recs = sorted(all_records, key=lambda x: (x["shop"], x["date"]))
    for i, rec in enumerate(sorted_recs, 1):
        r          = i + 2
        row_bg     = "gpay_row"    if rec["merchant"] == "GPay"   else "paytm_row"
        accent_key = "gpay_accent" if rec["merchant"] == "GPay"   else "paytm_accent"

        _data(ws, r, 1, i,                             row_bg, h_align="center")
        _data(ws, r, 2, rec["shop"],                   row_bg, bold=True,
              fg=accent_key, h_align="left")
        _data(ws, r, 3, rec["merchant"],               row_bg, bold=True,
              fg=accent_key)
        _data(ws, r, 4, rec["date"].strftime("%d %b %Y"), row_bg)
        _data(ws, r, 5, rec.get("payer", ""),          row_bg, h_align="left")
        _data(ws, r, 6, rec.get("txn_id", ""),         row_bg, h_align="left")
        _data(ws, r, 7, rec["amount"],                 row_bg,
              fmt="#,##0.00", h_align="right")

        for c in range(1, 8):
            ws.cell(row=r, column=c).border = _border()
        ws.row_dimensions[r].height = 17

    # Auto-filter on header row
    ws.auto_filter.ref = f"A2:G{len(sorted_recs) + 2}"

    _set_widths(ws, {"A": 5, "B": 24, "C": 12, "D": 16,
                      "E": 30, "F": 36, "G": 16, "H": 4})
    _freeze(ws, "A3")


def generate_excel(all_records: list[dict]) -> bytes:
    """Build a 3-sheet workbook from all collected records and return bytes."""
    # Pre-compute aggregations
    shop_daily: dict = defaultdict(lambda: defaultdict(float))
    shop_txns:  dict = defaultdict(list)

    for rec in all_records:
        key = (rec["shop"], rec["merchant"])
        shop_daily[key][rec["date"]] += rec["amount"]
        shop_txns[key].append(rec)

    wb = Workbook()
    _build_dashboard(wb, shop_daily, shop_txns)
    _build_shop_summary(wb, shop_daily, shop_txns)
    _build_all_transactions(wb, all_records)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════════
#  TELEGRAM BOT HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

WELCOME = (
    "👋 *Welcome to the Merchant Turnover Bot!*\n\n"
    "I compile *GPay Business* and *Paytm* merchant CSVs from multiple shops "
    "into a single, colour-coded Excel report.\n\n"
    "📤 *How to use*\n"
    "1. Send all your CSV files at once (select multiple & send together)\n"
    "2. Paytm CSVs are processed instantly — shop name is read automatically\n"
    "3. For GPay CSVs I'll ask the shop name for each one\n"
    "4. When done, use /compile to get the Excel report\n\n"
    "📋 *Commands*\n"
    "/status   — See what's collected so far\n"
    "/compile  — Generate & download the Excel report\n"
    "/reset    — Clear all data and start fresh\n\n"
    "Go ahead — send your first CSV file! 📁"
)


def _ensure_state(context: ContextTypes.DEFAULT_TYPE) -> None:
    context.user_data.setdefault("records",          [])
    context.user_data.setdefault("pending_csv",      None)   # legacy – kept for compat
    context.user_data.setdefault("gpay_queue",       [])     # list of (df, filename)
    context.user_data.setdefault("asking_gpay_name", False)


async def _prompt_next_gpay(update: Update,
                             context: ContextTypes.DEFAULT_TYPE) -> None:
    """Ask the user for the shop name of the next queued GPay CSV."""
    queue = context.user_data["gpay_queue"]
    if not queue:
        context.user_data["asking_gpay_name"] = False
        remaining = len(context.user_data["records"])
        await update.message.reply_text(
            f"✅ All files processed!  {remaining} transaction(s) collected.\n"
            "Use /compile to generate the Excel report.",
            parse_mode="Markdown",
        )
        return

    _, filename = queue[0]
    context.user_data["asking_gpay_name"] = True
    count = len(queue)
    await update.message.reply_text(
        f"📗 *GPay CSV:* `{filename}`\n"
        f"{'(' + str(count) + ' GPay file(s) left to name)  ' if count > 1 else ''}"
        f"What is the *shop name* for this file?\n"
        "_(Just type the name and send)_",
        parse_mode="Markdown",
    )


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    _ensure_state(context)
    await update.message.reply_text(WELCOME, parse_mode="Markdown")


async def cmd_status(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    _ensure_state(context)
    records = context.user_data["records"]

    if not records:
        await update.message.reply_text(
            "📭 Nothing collected yet.  Send some CSV files to get started!"
        )
        return

    shop_totals: dict[str, float] = defaultdict(float)
    shop_counts: dict[str, int]   = defaultdict(int)
    for r in records:
        key = f"{r['shop']}  ({r['merchant']})"
        shop_totals[key] += r["amount"]
        shop_counts[key] += 1

    lines = ["📊 *Current data:*\n"]
    for key in sorted(shop_totals):
        lines.append(
            f"🏪 *{key}*\n"
            f"   {shop_counts[key]} transactions  •  "
            f"₹{shop_totals[key]:,.2f}"
        )
    lines.append(
        f"\n💰 *Grand Total: ₹{sum(shop_totals.values()):,.2f}*\n\n"
        "Use /compile when ready."
    )
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def cmd_compile(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    _ensure_state(context)
    records = context.user_data["records"]

    if context.user_data["gpay_queue"]:
        remaining = len(context.user_data["gpay_queue"])
        await update.message.reply_text(
            f"⚠️ {remaining} GPay CSV(s) still need a shop name.\n"
            "Reply with the shop name to continue, then use /compile."
        )
        return

    if not records:
        await update.message.reply_text(
            "📭 Nothing to compile yet.  Send some CSV files first!"
        )
        return

    msg = await update.message.reply_text("⏳ Generating your Excel report…")

    try:
        excel_bytes = generate_excel(records)
    except Exception as exc:
        logger.exception("Excel generation failed")
        await msg.edit_text(f"❌ Error generating report:\n`{exc}`",
                            parse_mode="Markdown")
        return

    shops       = {f"{r['shop']} ({r['merchant']})" for r in records}
    grand_total = sum(r["amount"] for r in records)
    fname       = f"Merchant_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    bio      = io.BytesIO(excel_bytes)
    bio.name = fname

    await update.message.reply_document(
        document=bio,
        filename=fname,
        caption=(
            f"✅ *Report ready!*\n"
            f"🏪 Shops: {len(shops)}\n"
            f"📊 Transactions: {len(records)}\n"
            f"💰 Grand Total: ₹{grand_total:,.2f}"
        ),
        parse_mode="Markdown",
    )
    await msg.delete()


async def cmd_reset(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    context.user_data["records"]          = []
    context.user_data["pending_csv"]      = None
    context.user_data["gpay_queue"]       = []
    context.user_data["asking_gpay_name"] = False
    await update.message.reply_text(
        "🔄 All data cleared!  Send fresh CSVs whenever you're ready."
    )


async def handle_document(update: Update,
                           context: ContextTypes.DEFAULT_TYPE) -> None:
    _ensure_state(context)
    doc = update.message.document

    if not doc.file_name.lower().endswith(".csv"):
        await update.message.reply_text(
            "⚠️ That doesn't look like a CSV file.  "
            "Please send a `.csv` export from GPay or Paytm."
        )
        return

    # Download & detect
    tg_file = await doc.get_file()
    bio     = io.BytesIO()
    await tg_file.download_to_memory(bio)
    bio.seek(0)

    try:
        df = pd.read_csv(bio)
    except Exception as exc:
        await update.message.reply_text(
            f"❌ Could not read `{doc.file_name}`: `{exc}`",
            parse_mode="Markdown",
        )
        return

    csv_type = detect_csv_type(df)

    if csv_type == "gpay":
        # Queue the GPay CSV; will ask for name after any in-progress name entry
        context.user_data["gpay_queue"].append((df, doc.file_name))
        await update.message.reply_text(
            f"📗 *GPay CSV queued:* `{doc.file_name}`",
            parse_mode="Markdown",
        )
        # Start asking for names if not already doing so
        if not context.user_data["asking_gpay_name"]:
            await _prompt_next_gpay(update, context)

    elif csv_type == "paytm":
        records = parse_paytm_csv(df)
        if not records:
            await update.message.reply_text(
                f"⚠️ No `SUCCESS` transactions found in `{doc.file_name}`.\n"
                "Make sure you're using the correct merchant export."
            )
            return

        context.user_data["records"].extend(records)
        shops = sorted({r["shop"] for r in records})
        total = sum(r["amount"] for r in records)
        await update.message.reply_text(
            f"✅ *Paytm CSV processed:* `{doc.file_name}`\n"
            f"🏪 Shop(s): `{'`, `'.join(shops)}`\n"
            f"📊 Transactions: {len(records)}  •  💰 ₹{total:,.2f}",
            parse_mode="Markdown",
        )

    else:
        await update.message.reply_text(
            f"❓ `{doc.file_name}` — couldn't identify as GPay or Paytm.\n\n"
            "*GPay export* columns: `Creation time`, `Paid via`\n"
            "*Paytm export* columns: `Transaction_Date`, `Merchant_Name`",
            parse_mode="Markdown",
        )


async def handle_text(update: Update,
                      context: ContextTypes.DEFAULT_TYPE) -> None:
    _ensure_state(context)

    if not context.user_data["asking_gpay_name"]:
        await update.message.reply_text(
            "Send me CSV files, or use /start to see all commands."
        )
        return

    shop_name = update.message.text.strip()
    if not shop_name:
        await update.message.reply_text("Please enter a valid shop name (can't be empty).")
        return

    queue = context.user_data["gpay_queue"]
    if not queue:
        context.user_data["asking_gpay_name"] = False
        return

    # Process the first item in the queue
    df, filename = queue.pop(0)
    records = parse_gpay_csv(df, shop_name)
    context.user_data["records"].extend(records)

    if records:
        total = sum(r["amount"] for r in records)
        await update.message.reply_text(
            f"✅ *GPay processed:* `{filename}`\n"
            f"🏪 Shop: `{shop_name}`  •  📊 {len(records)} txns  •  💰 ₹{total:,.2f}",
            parse_mode="Markdown",
        )
    else:
        await update.message.reply_text(
            f"⚠️ No valid UPI transactions found in `{filename}` for *{shop_name}*.",
            parse_mode="Markdown",
        )

    # Ask for next GPay file's name, or finish
    await _prompt_next_gpay(update, context)


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    if not BOT_TOKEN:
        raise RuntimeError(
            "BOT_TOKEN is not set.  "
            "Copy .env.example → .env and add your token."
        )

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start",   cmd_start))
    app.add_handler(CommandHandler("help",    cmd_start))
    app.add_handler(CommandHandler("status",  cmd_status))
    app.add_handler(CommandHandler("compile", cmd_compile))
    app.add_handler(CommandHandler("reset",   cmd_reset))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info("Bot is running — press Ctrl+C to stop.")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()

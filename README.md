# Merchant Turnover Telegram Bot

Compiles GPay Business and Paytm merchant CSV exports from multiple shops
into a single colour-coded Excel report with daily turnovers and a grand total.

---

## Features

- Auto-detects GPay Business vs Paytm CSV format
- GPay: asks for shop name (not in the file), filters only settled UPI credits, skips SoundPod fees
- Paytm: reads shop name automatically from `Merchant_Name` column
- Excel report has 3 sheets:
  - **Dashboard** — daily turnover per shop, subtotals, grand total
  - **Shop Summary** — one row per shop with date range and totals
  - **All Transactions** — every individual transaction, filterable
- Colour-coded: GPay = blue, Paytm = green
- Works with as many shops / CSV files as you like in one session

---

## Setup (5 minutes)

### 1. Create a Telegram Bot

1. Open Telegram and message **[@BotFather](https://t.me/BotFather)**
2. Send `/newbot`
3. Follow the prompts (pick a name and username)
4. Copy the **bot token** you receive

### 2. Install Python dependencies

```bash
pip install python-telegram-bot pandas openpyxl python-dotenv
```

> Requires Python 3.10 or newer.

### 3. Configure your token

```bash
cp .env.example .env
```

Open `.env` and replace `your_telegram_bot_token_here` with your real token:

```
BOT_TOKEN=123456789:AAxxxxxxxxxxxxxxxxxxxxxxx
```

### 4. Run the bot

```bash
python bot.py
```

You should see:  `Bot is running — press Ctrl+C to stop.`

---

## How to use the bot

1. Open Telegram and find your bot (search by its @username)
2. Send `/start`
3. Send your CSV files one by one:
   - **GPay Business CSV** → bot asks "What's the shop name?" → reply with the name
   - **Paytm CSV** → shop name is read automatically from the file
4. Repeat for every merchant / shop CSV
5. Send `/compile` → bot sends back the Excel report
6. Send `/reset` to clear and start a new session

### Commands

| Command    | What it does                                |
|------------|---------------------------------------------|
| `/start`   | Welcome message and instructions            |
| `/status`  | Show shops collected so far with totals     |
| `/compile` | Generate and download the Excel report      |
| `/reset`   | Clear all data for a fresh run              |

---

## CSV formats supported

### GPay Business export
Export from the GPay Business app → Transactions → Export.

Key columns used: `Creation time`, `Type`, `Amount`, `Status`, `Payer/Receiver`

Rules applied:
- Only rows where `Type = UPI` (skips "Daily collections" / SoundPod fees)
- Only rows where `Status = Settled`
- Only rows where `Amount > 0` (skips fee deductions)

### Paytm merchant export
Export from Paytm for Business dashboard → Reports → Transaction Report.

Key columns used: `Transaction_Date`, `Merchant_Name`, `Amount`, `Status`

Rules applied:
- Only rows where `Status = SUCCESS`
- Only rows where `Amount > 0`

---

## Running in the background (optional)

To keep the bot running after you close your terminal, use `nohup` or `screen`:

```bash
# Option A — nohup
nohup python bot.py > bot.log 2>&1 &

# Option B — screen
screen -S merchant-bot
python bot.py
# Ctrl+A then D to detach; screen -r merchant-bot to re-attach
```

Or use `systemd` / PM2 for a proper always-on service.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `BOT_TOKEN is not set` | Make sure `.env` exists and has your token |
| "Couldn't identify this as GPay or Paytm" | Check that you exported the correct file from the app |
| GPay total seems wrong | Only UPI → Settled → positive amounts are counted; this is correct |
| Bot not responding | Make sure `python bot.py` is still running in your terminal |

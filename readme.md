

# Birthday Reminder Bot

A Python-based script for sending birthday messages via Slack, complete with internationalization (i18n) support, customizable messages, and coupon integration.

## Features

- Automatically detect birthdays and send Slack messages
- Manual message sending option
- Excel-based user and coupon data
- Locale-based message translation via JSON
- Safety mechanisms for test vs. real modes

## File Structure

- `app.py`: Main script that runs the birthday reminder logic
- `birthday_test.xlsx` / `birthday_real.xlsx`: Excel files with birthday data
- `coupon_test.xlsx` / `coupon_real.xlsx`: Excel files with coupon links and codes
- `i18n.json`: Language file containing localized message templates

## Setup

1. Create a `.env` file with your Slack Bot token:
   ```
   SLACK_BOT_TOKEN=xoxb-xxxxxxxxxxxx
   ```

2. Install dependencies:
   ```
   pip install python-dotenv openpyxl slack_sdk
   ```

3. Prepare your Excel files:
   - `birthday_*.xlsx` should contain the following columns:
     ```
     Name | Slack Display Name | Slack ID | Birthday | Join Date | Sent | Locale
     ```
   - `coupon_*.xlsx` should contain:
     ```
     Link | Code | Used
     ```

4. Create `i18n.json` with message templates for each supported locale.

## Running the Script

Run the script with:
```
python app.py
```

You will be prompted to choose:
- Option 1: Automatically send messages for recent birthdays
- Option 2: Manually select users to send messages

## Modes

The app automatically determines if it's in **TEST** or **REAL** mode based on the Excel filename (`*_test.xlsx` or `*_real.xlsx`). In TEST mode:
- Only the user with ID `U0000000000` will receive messages
- All others are blocked for safety

## Translation (i18n)

Message content is dynamically pulled from `i18n.json` based on the user's locale in the birthday file. Supported fields include:
- `birthday_greeting`
- `belated_birthday_greeting`
- `message_body`
- `redeem_link`

## Logging

Console messages indicate:
- Mode (REAL or TEST)
- Files in use
- Recipients of birthday messages
- Blocked messages due to test mode

## Maintenance Tips

- Always verify you're using the correct file mode to prevent accidental messages
- Check that "Sent" values in Excel are normalized (e.g., no extra spaces)
- Keep the `i18n.json` up to date as you add new locales

---
Maintained by Jacob Chen
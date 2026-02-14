Add 25 more small-to-medium trucking companies (5-100 trucks) to my leads.xlsx file. Focus on: Ideal Client Profile:
Family-owned or independent trucking companies with their OWN fleet
5-100 truck fleet size
NO large corporations with IT departments (avoid national carriers like TForce, Purolator, Day & Ross, etc.)
Companies that likely manage dispatch manually via phone/SMS
Active in GTA region (Brampton, Mississauga, Vaughan, Markham, Scarborough, Richmond Hill, Ajax, Pickering, Burlington, Milton, Oakville)
EXCLUDE:
Freight brokers and logistics companies that don't own trucks (they just arrange shipping)
Companies with names containing "Logistics," "Freight Broker," "Brokerage" unless they clearly operate their own fleet
Residential addresses (single driver-operators working from home)
Companies registered to house addresses
PRIORITIZE:
Commercial/industrial addresses (business parks, industrial roads, commercial units)
Companies with "Transport," "Trucking," "Carrier," "Express," "Hauling" in their name
Companies that employ multiple drivers and need active dispatch coordination
Owner-operator focused companies with office/yard locations
Requirements:
Research verified COMMERCIAL addresses and postal codes only
Check for duplicates against existing 76 companies in my file
Focus on companies with dispatchers managing multiple drivers, loads, settlements, detention tracking
Avoid duplicate addresses or companies already in the list
Add them to leads.xlsx with proper formatting (Company Name, Address, City, Province, Postal Code, empty Lob_Status/Lob_Date_Sent/Lob_ID columns).

# Lob Letter Automation

Automate sending physical letters using the Lob.com API. Reads lead data from an Excel file and sends personalized letters using a Lob template.

## Features

- Read leads from Excel file
- Send letters via Lob.com API
- Track sent status to avoid duplicates
- Auto-save progress every 5 letters
- Test and Live mode support
- Error handling and logging
- QR code integration in letters

## Prerequisites

- Python 3.7+
- Lob.com account with API keys
- Lob template created in your account

## Installation

1. Clone this repository:
```bash
git clone https://github.com/jerico-fullscope/lob-letter-automation.git
cd lob-letter-automation
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Configure environment variables:
   - Copy `.env.example` to `.env`
   - Add your Lob API keys
   - Add your Lob template ID
   - Update your return address details

## Configuration

Edit `.env` file with your credentials:

```env
# Lob API Keys
LOB_TEST_API_KEY=test_xxxxxxxxxxxxxxxxxxxxx
LOB_LIVE_API_KEY=live_xxxxxxxxxxxxxxxxxxxxx

# Template ID
LOB_TEMPLATE_ID=tmpl_xxxxxxxxxxxxx

# From Address
FROM_NAME=Your Name
FROM_COMPANY=Your Company
FROM_ADDRESS_LINE1=123 Main Street
FROM_ADDRESS_CITY=Your City
FROM_ADDRESS_STATE=XX
FROM_ADDRESS_ZIP=12345
FROM_ADDRESS_COUNTRY=CA

# Mode (test or live)
MODE=test
```

## Excel File Format

Your `leads.xlsx` file should have these columns:

| Column Name | Description |
|-------------|-------------|
| Company Name | Recipient company name |
| First Name | Recipient first name |
| Address | Street address |
| City | City |
| Province | Province/State code |
| Postal Code | ZIP/Postal code |
| Lob_Status | (Auto-filled) Status: SENT or ERROR |
| Lob_Date_Sent | (Auto-filled) Timestamp when sent |

## Usage

### Test Mode (Recommended First)

1. Set `MODE=test` in `.env`
2. Run the script:
```bash
python send_letters.py
```

Test mode uses your test API key and won't charge your account.

### Live Mode (Production)

1. Set `MODE=live` in `.env`
2. Run the script:
```bash
python send_letters.py
```

**Warning:** Live mode will send real letters and charge your Lob account!

## How It Works

1. Reads `leads.xlsx` from the current directory
2. For each row:
   - Skips if `Lob_Status` is already 'SENT'
   - Sends letter via Lob API using your template
   - Updates `Lob_Status` to 'SENT' or 'ERROR: [reason]'
   - Updates `Lob_Date_Sent` with timestamp
3. Saves progress to `leads_updated.xlsx` every 5 letters
4. Saves final results when complete

## Template Variables

The script passes these merge variables to your Lob template:

- `company_name` - From Excel
- `first_name` - From Excel
- `qr_url` - Configured QR URL
- `date` - Current date (formatted as "Month DD, YYYY")

Make sure your Lob template includes these variables!

## Output

- Console shows real-time progress
- `leads_updated.xlsx` contains updated status for all leads
- Original `leads.xlsx` remains unchanged

## Troubleshooting

**Invalid Address Errors:**
- Check that addresses are complete and valid
- Verify postal codes are correct format
- Lob validates addresses and rejects invalid ones

**API Key Errors:**
- Verify your API keys are correct in `.env`
- Check that keys match the mode (test vs live)

**Template Not Found:**
- Verify your template ID in `.env`
- Ensure template exists in your Lob account

## Safety Features

- Auto-saves progress every 5 letters
- Won't re-send to already-sent leads
- Test mode prevents accidental charges
- Clear visual warning when in LIVE mode

## License

MIT

## Support

For issues or questions, contact FullScope Services or open an issue on GitHub.

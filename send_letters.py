import pandas as pd
import lob
from datetime import datetime
import sys
import os
from dotenv import load_dotenv

# Fix Windows console encoding issues
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

# Load environment variables from .env file
load_dotenv()

# ===== API KEY CONFIGURATION =====
TEST_API_KEY = os.getenv('LOB_TEST_API_KEY')
LIVE_API_KEY = os.getenv('LOB_LIVE_API_KEY')
MODE = os.getenv('MODE', 'test').lower()

# Select API key based on MODE
if MODE == 'live':
    API_KEY = LIVE_API_KEY
    print("⚠️  LIVE MODE - Real letters will be sent and charged!\n")
else:
    API_KEY = TEST_API_KEY
    print("✓ TEST MODE - No charges will be applied\n")

# ===== LOB CONFIGURATION =====
LOB_TEMPLATE_ID = os.getenv('LOB_TEMPLATE_ID')

# Batch size configuration
BATCH_SIZE = int(os.getenv('BATCH_SIZE', '0'))  # 0 means send all

# Load configuration from .env
QR_URL = os.getenv('QR_URL', 'https://youtu.be/GlXtHlBPluc')
FROM_ADDRESS = {
    "name": os.getenv('FROM_NAME'),
    "company": os.getenv('FROM_COMPANY'),
    "address_line1": os.getenv('FROM_ADDRESS_LINE1'),
    "address_city": os.getenv('FROM_ADDRESS_CITY'),
    "address_state": os.getenv('FROM_ADDRESS_STATE'),
    "address_zip": os.getenv('FROM_ADDRESS_ZIP'),
    "address_country": os.getenv('FROM_ADDRESS_COUNTRY', 'CA')
}

# ===== INITIALIZE LOB CLIENT =====
lob.api_key = API_KEY

# ===== LOAD EXCEL FILE =====
input_file = "leads.xlsx"
output_file = "leads.xlsx"  # Save back to same file to track status

try:
    df = pd.read_excel(input_file)
    print(f"✓ Loaded {len(df)} rows from {input_file}")
    if BATCH_SIZE > 0:
        print(f"📦 Batch mode: Will send maximum {BATCH_SIZE} letters this run\n")
    else:
        print(f"📨 Sending all unsent letters\n")
except FileNotFoundError:
    print(f"ERROR: Could not find {input_file}")
    sys.exit(1)

# Ensure required columns exist
required_columns = ['Company Name', 'Address', 'City', 'Province', 'Postal Code']
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    print(f"ERROR: Missing required columns: {missing_columns}")
    sys.exit(1)

# Convert existing columns to string type and fill NaN values
df['Lob_Status'] = df['Lob_Status'].fillna('').astype(str)
df['Lob_Date_Sent'] = df['Lob_Date_Sent'].fillna('').astype(str)
df['Lob_ID'] = df['Lob_ID'].fillna('').astype(str) if 'Lob_ID' in df.columns else pd.Series([''] * len(df))

# ===== PROCESS EACH ROW =====
letters_sent = 0
save_interval = 5

for index, row in df.iterrows():
    company_name = row['Company Name']

    # Check if already sent
    if pd.notna(row['Lob_Status']) and row['Lob_Status'] == 'SENT':
        print(f"Skipping {company_name}... (already sent)")
        continue

    # Skip if Lob_Status has an error (optional - remove this if you want to retry errors)
    if pd.notna(row['Lob_Status']) and row['Lob_Status'].startswith('ERROR'):
        print(f"Skipping {company_name}... (previous error: {row['Lob_Status']})")
        continue

    # Check if batch limit reached
    if BATCH_SIZE > 0 and letters_sent >= BATCH_SIZE:
        print(f"\n📦 Batch limit reached ({BATCH_SIZE} letters sent). Stopping.\n")
        break

    # Prepare recipient address
    to_address = {
        "company": company_name,
        "address_line1": row['Address'],
        "address_city": row['City'],
        "address_state": row['Province'],
        "address_zip": str(row['Postal Code']),
        "address_country": "CA"
    }

    # Prepare merge variables for template
    merge_variables = {
        "company_name": company_name,
        "first_name": "Team",  # Placeholder since template requires it
        "qr_url": QR_URL,
        "date": datetime.now().strftime("%B %d, %Y")
    }

    # Send letter via Lob
    print(f"Sending to {company_name}...", end=" ")

    try:
        letter = lob.Letter.create(
            description=f"Letter to {company_name}",
            to_address=to_address,
            from_address=FROM_ADDRESS,
            file=LOB_TEMPLATE_ID,
            merge_variables=merge_variables,
            color=False,  # Black & white printing
            use_type="marketing"  # or "operational" for transactional mail
        )

        # Success!
        df.at[index, 'Lob_Status'] = 'SENT'
        df.at[index, 'Lob_Date_Sent'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df.at[index, 'Lob_ID'] = letter.id
        letters_sent += 1

        print(f"✓ Success! (ID: {letter.id})")

    except lob.error.InvalidRequestError as e:
        # Handle invalid address or request errors
        error_message = str(e)
        df.at[index, 'Lob_Status'] = f'ERROR: {error_message[:100]}'
        print(f"✗ Failed ({error_message})")

    except Exception as e:
        # Handle any other errors
        error_message = str(e)
        df.at[index, 'Lob_Status'] = f'ERROR: {error_message[:100]}'
        print(f"✗ Failed ({error_message})")

    # Save progress every 5 letters
    if letters_sent > 0 and letters_sent % save_interval == 0:
        df.to_excel(output_file, index=False)
        print(f"\n💾 Progress saved to {output_file} ({letters_sent} letters sent)\n")

# ===== FINAL SAVE =====
df.to_excel(output_file, index=False)
print(f"\n✓ Complete! Saved final results to {output_file}")
print(f"📊 Total letters sent this run: {letters_sent}")

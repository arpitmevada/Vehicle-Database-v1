# Vehicle Database v1 (free pilot)

Upload 4 photos (VIN plate, Hours, Odometer, Front). Photos are shrunk on the phone, saved to a Google Sheet,
and downloaded as Excel with the photos placed in cells. No AI, no paid service.

Files:
- app_photos.py      main app (set this as the Streamlit main file)
- core.py            helper code
- requirements.txt   Python packages
- photo_picker/index.html   phone-side photo resize (must stay in this folder)

Secrets (Streamlit > Manage app > Settings > Secrets):

SHEET_ID = "your-google-sheet-id"
APP_PASSWORD = "your-team-password"
GCP_JSON = '''
{ paste the full service-account JSON here }
'''

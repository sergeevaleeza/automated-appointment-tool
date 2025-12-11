# **[Automated Appointment Tool](https://automated-appointment-tool.streamlit.app/)**
===========================

A Streamlit-based application for processing appointment data from CSV files, matching it with patient information from an Excel file, and generating customized Excel reports split by provider (e.g., doctors). The app allows users to configure provider mappings dynamically and outputs a ZIP file containing individual reports and a processing summary.

## Features
- Upload appointment CSV and patient Excel files.
- Automatically detect unique providers from the data and allow custom short names for output files.
- Fuzzy matching of patient names between files.
- Generate provider-specific Excel reports with formatted patient data, insurance, dates, status, and codes.
- Produce a detailed processing summary with statistics.
- Output everything as a downloadable ZIP archive.
- Generalized for any similar data processing use case (e.g., clinics, scheduling systems).

## Installation
1. Clone the repository:
   
   ```bash
   git clone https://github.com/yourusername/streamlit-appointment-processor.git
   cd streamlit-appointment-processor
   ```
   
3. Install dependencies:
   
   ```bash
   pip install -r requirements.txt
   ```
   
(Requirements: streamlit, pandas, openpyxl, zipfile)

## Usage
1. Run the app:
   
   ```bash
   streamlit run app.py
   ```
   
3. In the web interface:
- Upload your appointment CSV (expected columns: AppointmentTime, Patient, SeenBy, AppointmentStatus, etc.).
- Upload your patient Excel file (expected sheet: "Active"; columns: Name (Last, First), Code, Insurance).
- Enter a period string (e.g., "11_2025") for output file naming.
- Configure short names for each detected provider.
- Click "Process Files" to generate and download the ZIP.

## Example Data
- Appointment CSV: Contains scheduled visits with patient names, dates, providers, and status.
- Patient Excel: List of patients with insurance details for matching.

## Contributing
Pull requests welcome! For major changes, open an issue first.

## License
MIT

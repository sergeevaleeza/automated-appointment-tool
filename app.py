#!/usr/bin/env python3
import streamlit as st
import pandas as pd
from datetime import datetime
from typing import Tuple, Dict, List, Optional
import io
import zipfile
import os

# --- Configuration ---
SUFFIXES = {"jr", "sr", "ii", "iii", "iv", "v", "md", "phd", "psyd", "do"}

# Column positions for output Excel files (0-based)
IDX_NAME = 0            # Column 1: "last Name, First Name"
IDX_INSURANCE = 1       # Column 2: "insurance"
IDX_DATE = 2            # Column 3: "date of visit"
IDX_STATUS = 3          # Column 4: "status"
IDX_CODES = 6           # Column 7: "codes"

class PatientProcessor:
    def __init__(self):
        self.exact_map = {}
        self.last_to_firsts = {}
        self.stats = {
            'total_appointments': 0,
            'matched_patients': 0,
            'doctors_processed': {}
        }
    
    def clean_spaces(self, s: str) -> str:
        """Clean up spacing issues in strings"""
        if pd.isna(s) or s == "":
            return ""
        s = str(s).strip()
        s = s.replace(" ,", ",").replace(",  ", ", ").replace("  ", " ")
        return " ".join(s.split())
    
    def normalize_basic(self, x) -> str:
        """Basic normalization for string comparison"""
        if pd.isna(x):
            return ""
        s = self.clean_spaces(str(x).strip().lower())
        s = s.replace(".", "").replace("'", "")
        return s
    
    def strip_suffixes(self, name_part: str) -> str:
        """Remove common suffixes from name parts"""
        parts = [p for p in name_part.split() if p]
        while parts and parts[-1].lower() in SUFFIXES:
            parts.pop()
        return " ".join(parts)
    
    def parse_patient_name(self, name: str, format_hint: str = "auto") -> Tuple[str, str]:
        """
        Parse patient name with support for complex names.
        Returns (last_name, first_name) both normalized.
        
        format_hint: "auto", "last_first" (Last, First), or "first_last" (First Last)
        """
        if pd.isna(name) or name == "":
            return ("", "")
        
        name = str(name).strip()
        
        # Auto-detect format or use hint
        if format_hint == "auto":
            has_comma = ',' in name
            format_to_use = "last_first" if has_comma else "first_last"
        else:
            format_to_use = format_hint
        
        if format_to_use == "last_first":
            # Format: "Last, First" or "Last,First"
            if ',' in name:
                parts = name.split(',', 1)
                last_name = parts[0].strip()
                first_name = parts[1].strip() if len(parts) > 1 else ""
            else:
                # No comma, treat as last name only
                last_name = name
                first_name = ""
        else:
            # Format: "First Last" or "First Middle Last"
            parts = name.split()
            if len(parts) == 0:
                return ("", "")
            elif len(parts) == 1:
                # Single name - could be first or last
                last_name = parts[0]
                first_name = ""
            else:
                # Multiple parts - first part(s) are first/middle, last is surname
                first_name = " ".join(parts[:-1])
                last_name = parts[-1]
        
        # Handle complex last names like "Russell (Kwon)"
        if '(' in last_name:
            # Extract the main last name before parentheses
            main_last_name = last_name.split('(')[0].strip()
            last_name = main_last_name
        
        # Normalize and strip suffixes
        last_name = self.strip_suffixes(self.normalize_basic(last_name))
        first_name = self.strip_suffixes(self.normalize_basic(first_name))
        
        # Extract first token from first name for matching
        first_token = first_name.split()[0] if first_name else ""
        
        return (last_name, first_token)
    
    def names_prefix_match(self, a_first: str, b_first: str) -> bool:
        """Check if either first name starts with the other"""
        if not a_first or not b_first:
            return False
        return a_first.startswith(b_first) or b_first.startswith(a_first)
    
    def build_mutual_index(self, mutual_df):
        """Build index structures for fast patient lookup"""
        self.exact_map = {}
        self.last_to_firsts = {}
        
        for _, row in mutual_df.iterrows():
            # Assuming mutual file has: name (Last, First), insurance code, insurance name
            # Parse with "last_first" format hint since mutual file uses "Last, First"
            last, first_tok = self.parse_patient_name(row.iloc[0], format_hint="last_first")
            
            if not last:  # Skip if no valid last name
                continue
            
            pair = (last, first_tok) if first_tok else (last, "")
            
            # Store insurance code (col 2) and insurance name (col 3)
            insurance_code = str(row.iloc[1]) if not pd.isna(row.iloc[1]) else ""
            insurance_name = str(row.iloc[2]) if not pd.isna(row.iloc[2]) else ""
            
            if pair not in self.exact_map:
                self.exact_map[pair] = (insurance_code, insurance_name)
            
            if first_tok:
                self.last_to_firsts.setdefault(last, []).append(
                    (first_tok, (insurance_code, insurance_name))
                )
    
    def lookup_mutual(self, last: str, first_token: str) -> Tuple[Optional[str], Optional[str]]:
        """
        Look up patient in mutual file.
        Returns (insurance_code, insurance_name) or (None, None)
        """
        if not last:
            return (None, None)
        
        # Try exact match first
        pair = (last, first_token) if first_token else (last, "")
        if pair in self.exact_map:
            return self.exact_map[pair]
        
        # Try without first name
        if first_token and (last, "") in self.exact_map:
            return self.exact_map[(last, "")]
        
        # Try prefix matching on first names
        if first_token:
            candidates = self.last_to_firsts.get(last, [])
            best = None
            best_len = -1
            
            for cand_first, data in candidates:
                if self.names_prefix_match(first_token, cand_first):
                    common_len = len(os.path.commonprefix([first_token, cand_first]))
                    if common_len > best_len:
                        best_len = common_len
                        best = data
            
            if best:
                return best
        
        return (None, None)
    
    def format_date(self, val) -> str:
        """Format date as MM/DD/YYYY"""
        if pd.isna(val):
            return ""
        
        try:
            # Try to parse various date formats
            date_obj = pd.to_datetime(val)
            return date_obj.strftime("%m/%d/%Y")
        except:
            return str(val)
    
    def append_value(self, existing, new, sep=" | "):
        """Append new value to existing with separator, avoiding duplicates"""
        if pd.isna(new) or str(new).strip() == "":
            return existing
        
        new_s = str(new).strip()
        if pd.isna(existing) or str(existing).strip() == "":
            return new_s
        
        existing_s = str(existing).strip()
        parts = [p.strip() for p in existing_s.split(sep)]
        
        if new_s not in parts:
            return existing_s + sep + new_s
        return existing_s
    
    def process_status(self, status: str) -> str:
        """
        Process appointment status: remove 'Seen' but keep other statuses.
        """
        if pd.isna(status) or status == "":
            return ""
        
        status_str = str(status).strip()
        
        # If status is exactly "Seen", return empty string
        if status_str.lower() == "seen":
            return ""
        
        # Keep all other statuses (Pending, Canceled, etc.)
        return status_str
    
    def process_appointments(self, appointment_df, doctor_mapping: Dict[str, str]) -> Dict[str, pd.DataFrame]:
        """
        Process appointments and split by doctor.
        Returns dictionary: doctor_short -> DataFrame
        """
        doctor_dfs = {}
        
        for doctor_full, doctor_short in doctor_mapping.items():
            # Filter appointments for this doctor
            doctor_appointments = appointment_df[
                appointment_df['SeenBy'] == doctor_full
            ].copy()
            
            if len(doctor_appointments) == 0:
                continue
            
            # Create output DataFrame with required columns
            output_rows = []
            matched_count = 0
            
            for _, apt in doctor_appointments.iterrows():
                # Parse patient name (appointments use "First Last" format)
                patient_name = apt['Patient']
                last, first = self.parse_patient_name(patient_name, format_hint="first_last")
                
                # Format name for output (Last, First)
                if last and first:
                    # Capitalize properly for display
                    last_display = " ".join(w.capitalize() for w in last.split())
                    first_display = " ".join(w.capitalize() for w in first.split())
                    formatted_name = f"{last_display}, {first_display}"
                elif last:
                    formatted_name = last.capitalize()
                else:
                    formatted_name = str(patient_name)
                
                # Look up insurance and codes from mutual file
                codes, insurance = self.lookup_mutual(last, first)
                
                if codes or insurance:
                    matched_count += 1
                
                # Prepare row data
                row_data = [''] * 7  # Initialize with empty strings
                
                row_data[IDX_NAME] = formatted_name
                row_data[IDX_INSURANCE] = insurance if insurance else ""
                row_data[IDX_DATE] = self.format_date(apt['AppointmentTime'])
                row_data[IDX_STATUS] = self.process_status(apt['AppointmentStatus'])
                row_data[IDX_CODES] = codes if codes else ""
                
                output_rows.append(row_data)
            
            # Create DataFrame from rows
            doctor_df = pd.DataFrame(output_rows)
            doctor_dfs[doctor_short] = doctor_df
            
            # Update statistics
            self.stats['doctors_processed'][doctor_short] = {
                'total': len(doctor_appointments),
                'matched': matched_count
            }
        
        return doctor_dfs
    
    def generate_summary(self, doctor_mapping: Dict[str, str], period_str: str) -> str:
        """Generate processing summary as string"""
        lines = []
        lines.append("=" * 60)
        lines.append("DOCTOR VISIT PROCESSING SUMMARY")
        lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append("=" * 60 + "\n")
        
        lines.append(f"Input Files:\n")
        lines.append(f"  - Appointments: Uploaded CSV\n")
        lines.append(f"  - Mutual Data: Uploaded XLSX\n\n")
        
        lines.append(f"Total appointments processed: {self.stats['total_appointments']}\n")
        lines.append(f"Doctors processed: {len(self.stats['doctors_processed'])}\n\n")
        
        total_matched = sum(d['matched'] for d in self.stats['doctors_processed'].values())
        total_processed = sum(d['total'] for d in self.stats['doctors_processed'].values())
        
        if total_processed > 0:
            match_rate = 100 * total_matched / total_processed
            lines.append(f"Overall matching rate: {total_matched}/{total_processed} ({match_rate:.1f}%)\n\n")
        
        lines.append("Per-doctor statistics:\n")
        lines.append("-" * 50)
        lines.append(f"{'Doctor':<15} {'Total':<8} {'Matched':<8} {'Rate':<8}")
        lines.append("-" * 50)
        
        for doctor, stats in self.stats['doctors_processed'].items():
            match_rate = 100 * stats['matched'] / stats['total'] if stats['total'] > 0 else 0
            lines.append(f"{doctor:<15} {stats['total']:<8} {stats['matched']:<8} {match_rate:.1f}%")
        
        lines.append("-" * 50 + "\n")
        
        # Add list of generated files
        lines.append("Generated Files:\n")
        for doctor in sorted(self.stats['doctors_processed'].keys()):
            lines.append(f"  - {doctor}_visits_{period_str}.xlsx")
        lines.append(f"  - processing_summary_{period_str}.txt\n")
        
        return "\n".join(lines)
    
    def run(self, appointment_df: pd.DataFrame, mutual_df: pd.DataFrame, doctor_mapping: Dict[str, str], period_str: str):
        """Main processing function"""
        self.stats['total_appointments'] = len(appointment_df)
        
        # Select relevant columns from mutual_df
        mutual_df = mutual_df.iloc[:, [0, 1, 2]]
        
        self.build_mutual_index(mutual_df)
        
        doctor_dfs = self.process_appointments(appointment_df, doctor_mapping)
        
        # Generate summary text
        summary_text = self.generate_summary(doctor_mapping, period_str)
        
        # Create zip file in memory
        zip_output = io.BytesIO()
        with zipfile.ZipFile(zip_output, 'w') as zipf:
            for doctor_short, df in doctor_dfs.items():
                excel_buf = io.BytesIO()
                with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
                    df.to_excel(writer, header=False, index=False)
                excel_buf.seek(0)
                zipf.writestr(f"{doctor_short}_visits_{period_str}.xlsx", excel_buf.getvalue())
            
            # Add summary to zip
            zipf.writestr(f"processing_summary_{period_str}.txt", summary_text.encode())
        
        zip_output.seek(0)
        
        return zip_output, summary_text

# Streamlit App
st.title("Appointment Processor App")

st.markdown("""
This app processes appointment data from a CSV file and matches it with patient information from an Excel file.
It generates Excel files for each doctor and a summary report, bundled in a ZIP file.
""")

csv_file = st.file_uploader("Upload Appointment CSV", type=["csv"])
xlsx_file = st.file_uploader("Upload Mutual Patients XLSX", type=["xlsx", "xls"])
period_str = st.text_input("Period (e.g., 11_2025)", value="11_2025")
mutual_sheet = "Active"  # Fixed sheet name, can be made configurable if needed

if csv_file and xlsx_file:
    appointment_df = pd.read_csv(csv_file, dtype=str)
    mutual_df = pd.read_excel(xlsx_file, sheet_name=mutual_sheet, header=None, dtype=str, engine='openpyxl')
    
    unique_doctors = sorted(appointment_df['SeenBy'].unique())
    
    st.subheader("Configure Doctor Short Names")
    st.markdown("Provide short names for each doctor (used in filenames and summary). Defaults to the first word of the full name.")
    
    doctor_mapping = {}
    for doc in unique_doctors:
        default_short = doc.split()[0] if doc.split() else doc
        short = st.text_input(f"Short name for '{doc}'", value=default_short)
        doctor_mapping[doc] = short
    
    if st.button("Process Files"):
        processor = PatientProcessor()
        zip_bytes, summary_text = processor.run(appointment_df, mutual_df, doctor_mapping, period_str)
        
        st.success("Processing complete!")
        
        st.download_button(
            label="Download Results ZIP",
            data=zip_bytes,
            file_name=f"visits_{period_str}.zip",
            mime="application/zip"
        )
        
        st.subheader("Processing Summary")
        st.text_area("Summary", summary_text, height=400)
else:
    st.info("Please upload both the CSV and XLSX files to proceed.")

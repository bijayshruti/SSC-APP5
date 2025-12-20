# app.py
import streamlit as st
import pandas as pd
from datetime import datetime
import json
import logging
import os
from pathlib import Path
import io
import base64
import tempfile
from collections import defaultdict

# Constants
BASE_DIR = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "config.json"
DATA_FILE = BASE_DIR / "allocations_data.json"
REFERENCE_FILE = BASE_DIR / "allocation_references.json"
DELETED_RECORDS_FILE = BASE_DIR / "deleted_records.json"
BACKUP_DIR = BASE_DIR / "backups"

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename=BASE_DIR / 'app.log'
)

# Initialize session state
def init_session_state():
    if 'io_df' not in st.session_state:
        default_data = """NAME,AREA,CENTRE_CODE,MOBILE,EMAIL
John Doe,Kolkata,1001,9876543210,john@example.com
Jane Smith,Howrah,1002,9876543211,jane@example.com
Robert Johnson,Hooghly,1003,9876543212,robert@example.com
Emily Davis,Nadia,2001,9876543213,emily@example.com
Michael Wilson,North 24 Parganas,2002,9876543214,michael@example.com"""
        
        st.session_state.io_df = pd.read_csv(io.StringIO(default_data))
        st.session_state.io_df['CENTRE_CODE'] = st.session_state.io_df['CENTRE_CODE'].astype(str).str.zfill(4)
    
    if 'venue_df' not in st.session_state:
        st.session_state.venue_df = pd.DataFrame()
    
    if 'ey_df' not in st.session_state:
        st.session_state.ey_df = pd.DataFrame()
    
    if 'allocation' not in st.session_state:
        st.session_state.allocation = []
    
    if 'ey_allocation' not in st.session_state:
        st.session_state.ey_allocation = []
    
    if 'deleted_records' not in st.session_state:
        st.session_state.deleted_records = []
    
    if 'exam_data' not in st.session_state:
        st.session_state.exam_data = {}
    
    if 'current_exam_key' not in st.session_state:
        st.session_state.current_exam_key = ""
    
    if 'exam_name' not in st.session_state:
        st.session_state.exam_name = ""
    
    if 'exam_year' not in st.session_state:
        st.session_state.exam_year = ""
    
    if 'allocation_references' not in st.session_state:
        st.session_state.allocation_references = {}
    
    if 'remuneration_rates' not in st.session_state:
        st.session_state.remuneration_rates = {
            'multiple_shifts': 750,
            'single_shift': 450,
            'mock_test': 450,
            'ey_personnel': 5000
        }
    
    if 'ey_personnel_list' not in st.session_state:
        st.session_state.ey_personnel_list = []
    
    if 'selected_venue' not in st.session_state:
        st.session_state.selected_venue = ""
    
    if 'selected_role' not in st.session_state:
        st.session_state.selected_role = "Centre Coordinator"
    
    if 'selected_dates' not in st.session_state:
        st.session_state.selected_dates = {}
    
    if 'mock_test_mode' not in st.session_state:
        st.session_state.mock_test_mode = False
    
    if 'ey_allocation_mode' not in st.session_state:
        st.session_state.ey_allocation_mode = False
    
    if 'selected_ey_personnel' not in st.session_state:
        st.session_state.selected_ey_personnel = ""
    
    if 'selected_ey_venues' not in st.session_state:
        st.session_state.selected_ey_venues = []

# Load data from files
def load_data():
    try:
        # Load config
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
                if isinstance(config, dict):
                    if 'remuneration_rates' in config:
                        st.session_state.remuneration_rates.update(config['remuneration_rates'])
                    if 'ey_personnel_list' in config:
                        st.session_state.ey_personnel_list = config['ey_personnel_list']
        
        # Load exam data
        if DATA_FILE.exists():
            with open(DATA_FILE, 'r') as f:
                data = json.load(f)
                if isinstance(data, dict):
                    st.session_state.exam_data = data
        
        # Load references
        if REFERENCE_FILE.exists():
            with open(REFERENCE_FILE, 'r') as f:
                st.session_state.allocation_references = json.load(f)
        
        # Load deleted records
        if DELETED_RECORDS_FILE.exists():
            with open(DELETED_RECORDS_FILE, 'r') as f:
                st.session_state.deleted_records = json.load(f)
                
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")

# Save data to files
def save_data():
    try:
        # Save config
        config = {
            'remuneration_rates': st.session_state.remuneration_rates,
            'ey_personnel_list': st.session_state.ey_personnel_list
        }
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=4)
        
        # Save exam data
        if st.session_state.current_exam_key:
            st.session_state.exam_data[st.session_state.current_exam_key] = {
                'io_allocations': st.session_state.allocation,
                'ey_allocations': st.session_state.ey_allocation
            }
        
        with open(DATA_FILE, 'w') as f:
            json.dump(st.session_state.exam_data, f, indent=4, default=str)
        
        # Save references
        with open(REFERENCE_FILE, 'w') as f:
            json.dump(st.session_state.allocation_references, f, indent=4)
        
        # Save deleted records
        with open(DELETED_RECORDS_FILE, 'w') as f:
            json.dump(st.session_state.deleted_records, f, indent=4, default=str)
            
    except Exception as e:
        st.error(f"Error saving data: {str(e)}")

# Helper function for file downloads
def get_download_link(df, filename, text):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    return f'<a href="data:file/csv;base64,{b64}" download="{filename}">{text}</a>'

# Main app
def main():
    st.set_page_config(
        page_title="SSC (ER) Kolkata - Allocation System",
        page_icon="🏛️",
        layout="wide"
    )
    
    init_session_state()
    load_data()
    
    # Header
    st.title("🏛️ STAFF SELECTION COMMISSION (ER), KOLKATA")
    st.subheader("Centre Coordinator & Flying Squad Allocation System")
    
    # Main tabs
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📋 Exam Management", 
        "👥 Centre Coordinator Allocation", 
        "👁️ EY Personnel Allocation",
        "📊 Reports & Export",
        "⚙️ Settings & Tools"
    ])
    
    # Tab 1: Exam Management
    with tab1:
        st.header("Exam Information Management")
        
        col1, col2, col3 = st.columns([2, 1, 1])
        
        with col1:
            exam_options = sorted(st.session_state.exam_data.keys())
            selected_exam = st.selectbox(
                "Select Existing Exam",
                options=[""] + exam_options,
                key="exam_selector"
            )
            
            if selected_exam and selected_exam != st.session_state.current_exam_key:
                st.session_state.current_exam_key = selected_exam
                if selected_exam in st.session_state.exam_data:
                    exam_data = st.session_state.exam_data[selected_exam]
                    if isinstance(exam_data, dict):
                        st.session_state.allocation = exam_data.get('io_allocations', [])
                        st.session_state.ey_allocation = exam_data.get('ey_allocations', [])
                    else:
                        st.session_state.allocation = exam_data
                        st.session_state.ey_allocation = []
                    
                    if " - " in selected_exam:
                        name, year = selected_exam.split(" - ", 1)
                        st.session_state.exam_name = name
                        st.session_state.exam_year = year
                
                st.success(f"Loaded exam: {selected_exam}")
                st.rerun()
        
        with col2:
            st.session_state.exam_name = st.text_input(
                "New Exam Name",
                value=st.session_state.exam_name,
                key="new_exam_name"
            )
            
            current_year = datetime.now().year
            year_options = [str(y) for y in range(current_year-5, current_year+3)]
            st.session_state.exam_year = st.selectbox(
                "Exam Year",
                options=[""] + year_options,
                index=0 if not st.session_state.exam_year else year_options.index(st.session_state.exam_year),
                key="new_exam_year"
            )
        
        with col3:
            st.write("### Actions")
            
            if st.button("🚀 Create/Update Exam", use_container_width=True):
                if not st.session_state.exam_name or not st.session_state.exam_year:
                    st.error("Please enter both Exam Name and Year")
                else:
                    exam_key = f"{st.session_state.exam_name} - {st.session_state.exam_year}"
                    st.session_state.current_exam_key = exam_key
                    
                    if exam_key not in st.session_state.exam_data:
                        st.session_state.exam_data[exam_key] = {
                            'io_allocations': [],
                            'ey_allocations': []
                        }
                    
                    save_data()
                    st.success(f"Exam set: {exam_key}")
                    st.rerun()
            
            if st.session_state.current_exam_key and st.button("🗑️ Delete Exam", use_container_width=True):
                confirm = st.checkbox("I confirm I want to delete this exam and all its allocations")
                if confirm:
                    # Create backup
                    backup_file = create_backup(st.session_state.current_exam_key)
                    
                    del st.session_state.exam_data[st.session_state.current_exam_key]
                    st.session_state.allocation = []
                    st.session_state.ey_allocation = []
                    st.session_state.current_exam_key = ""
                    st.session_state.exam_name = ""
                    st.session_state.exam_year = ""
                    
                    save_data()
                    st.success(f"Exam deleted successfully!")
                    if backup_file:
                        st.info(f"Backup created: {backup_file.name}")
                    st.rerun()
        
        # Display current exam info
        if st.session_state.current_exam_key:
            st.info(f"**Current Exam:** {st.session_state.current_exam_key}")
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Centre Coordinator Allocations", len(st.session_state.allocation))
            with col2:
                st.metric("EY Personnel Allocations", len(st.session_state.ey_allocation))
        
        # Allocation References Section
        st.divider()
        st.subheader("📋 Allocation References")
        
        if st.session_state.current_exam_key:
            exam_key = st.session_state.current_exam_key
            
            if exam_key not in st.session_state.allocation_references:
                st.session_state.allocation_references[exam_key] = {}
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.write("**Centre Coordinator**")
                if 'Centre Coordinator' in st.session_state.allocation_references[exam_key]:
                    ref = st.session_state.allocation_references[exam_key]['Centre Coordinator']
                    st.write(f"Order No.: {ref.get('order_no', 'N/A')}")
                    st.write(f"Page No.: {ref.get('page_no', 'N/A')}")
                    
                    if st.button("✏️ Update", key="update_cc_ref"):
                        update_allocation_reference("Centre Coordinator")
                else:
                    st.write("No reference set")
                    if st.button("➕ Add Reference", key="add_cc_ref"):
                        add_allocation_reference("Centre Coordinator")
            
            with col2:
                st.write("**Flying Squad**")
                if 'Flying Squad' in st.session_state.allocation_references[exam_key]:
                    ref = st.session_state.allocation_references[exam_key]['Flying Squad']
                    st.write(f"Order No.: {ref.get('order_no', 'N/A')}")
                    st.write(f"Page No.: {ref.get('page_no', 'N/A')}")
                    
                    if st.button("✏️ Update", key="update_fs_ref"):
                        update_allocation_reference("Flying Squad")
                else:
                    st.write("No reference set")
                    if st.button("➕ Add Reference", key="add_fs_ref"):
                        add_allocation_reference("Flying Squad")
            
            with col3:
                st.write("**EY Personnel**")
                if 'EY Personnel' in st.session_state.allocation_references[exam_key]:
                    ref = st.session_state.allocation_references[exam_key]['EY Personnel']
                    st.write(f"Order No.: {ref.get('order_no', 'N/A')}")
                    st.write(f"Page No.: {ref.get('page_no', 'N/A')}")
                    
                    if st.button("✏️ Update", key="update_ey_ref"):
                        update_allocation_reference("EY Personnel")
                else:
                    st.write("No reference set")
                    if st.button("➕ Add Reference", key="add_ey_ref"):
                        add_allocation_reference("EY Personnel")
        
        # View All References
        st.divider()
        if st.button("👁️ View All References"):
            view_allocation_references()
        
        # View Deleted Records
        if st.button("🗑️ View Deleted Records"):
            view_deleted_records()
    
    # Tab 2: Centre Coordinator Allocation
    with tab2:
        st.header("Centre Coordinator Allocation")
        
        if not st.session_state.current_exam_key:
            st.warning("Please select or create an exam first from the Exam Management tab")
        else:
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.subheader("Step 1: Load Master Data")
                
                col1a, col1b = st.columns(2)
                with col1a:
                    io_file = st.file_uploader(
                        "Load Centre Coordinator Master",
                        type=["xlsx", "xls"],
                        key="io_master_upload"
                    )
                    if io_file:
                        try:
                            st.session_state.io_df = pd.read_excel(io_file)
                            st.session_state.io_df.columns = [str(col).strip().upper() for col in st.session_state.io_df.columns]
                            
                            required_cols = ["NAME", "AREA", "CENTRE_CODE"]
                            missing_cols = [col for col in required_cols if col not in st.session_state.io_df.columns]
                            
                            if missing_cols:
                                st.error(f"Missing required columns: {', '.join(missing_cols)}")
                            else:
                                if 'CENTRE_CODE' in st.session_state.io_df.columns:
                                    st.session_state.io_df['CENTRE_CODE'] = st.session_state.io_df['CENTRE_CODE'].astype(str).str.zfill(4)
                                st.success(f"Loaded {len(st.session_state.io_df)} Centre Coordinator records")
                        except Exception as e:
                            st.error(f"Error loading file: {str(e)}")
                
                with col1b:
                    venue_file = st.file_uploader(
                        "Load Venue List",
                        type=["xlsx", "xls"],
                        key="venue_upload"
                    )
                    if venue_file:
                        try:
                            st.session_state.venue_df = pd.read_excel(venue_file)
                            st.session_state.venue_df.columns = [str(col).strip().upper() for col in st.session_state.venue_df.columns]
                            
                            required_cols = ["VENUE", "DATE", "SHIFT", "CENTRE_CODE", "ADDRESS"]
                            missing_cols = [col for col in required_cols if col not in st.session_state.venue_df.columns]
                            
                            if missing_cols:
                                st.error(f"Missing required columns: {', '.join(missing_cols)}")
                            else:
                                st.session_state.venue_df['VENUE'] = st.session_state.venue_df['VENUE'].astype(str).str.strip()
                                st.session_state.venue_df['CENTRE_CODE'] = st.session_state.venue_df['CENTRE_CODE'].astype(str).str.zfill(4)
                                st.session_state.venue_df['DATE'] = pd.to_datetime(st.session_state.venue_df['DATE'], errors='coerce').dt.strftime('%d-%m-%Y')
                                st.success(f"Loaded {len(st.session_state.venue_df)} venue records")
                        except Exception as e:
                            st.error(f"Error loading file: {str(e)}")
                
                # Display loaded data info
                if not st.session_state.io_df.empty:
                    st.info(f"**Centre Coordinators Loaded:** {len(st.session_state.io_df)} records")
                if not st.session_state.venue_df.empty:
                    st.info(f"**Venues Loaded:** {len(st.session_state.venue_df)} records")
            
            with col2:
                st.subheader("Step 2: Configuration")
                
                st.session_state.mock_test_mode = st.checkbox(
                    "Mock Test Mode",
                    value=st.session_state.mock_test_mode,
                    key="mock_test_checkbox"
                )
                
                st.session_state.selected_role = st.selectbox(
                    "Select Role",
                    options=["Centre Coordinator", "Flying Squad"],
                    index=0,
                    key="role_selector"
                )
                
                # Remuneration Rates
                st.subheader("Remuneration Rates")
                col_rate1, col_rate2 = st.columns(2)
                with col_rate1:
                    st.session_state.remuneration_rates['multiple_shifts'] = st.number_input(
                        "Multiple Shifts (₹)",
                        min_value=0,
                        value=st.session_state.remuneration_rates['multiple_shifts'],
                        key="multi_shift_rate"
                    )
                    st.session_state.remuneration_rates['single_shift'] = st.number_input(
                        "Single Shift (₹)",
                        min_value=0,
                        value=st.session_state.remuneration_rates['single_shift'],
                        key="single_shift_rate"
                    )
                with col_rate2:
                    st.session_state.remuneration_rates['mock_test'] = st.number_input(
                        "Mock Test (₹)",
                        min_value=0,
                        value=st.session_state.remuneration_rates['mock_test'],
                        key="mock_test_rate"
                    )
                
                if st.button("💾 Save Rates", use_container_width=True):
                    save_data()
                    st.success("Rates saved successfully!")
            
            st.divider()
            
            # Step 3: Venue and Date Selection
            st.subheader("Step 3: Select Venue & Dates")
            
            if not st.session_state.venue_df.empty:
                venues = sorted(st.session_state.venue_df['VENUE'].dropna().unique())
                st.session_state.selected_venue = st.selectbox(
                    "Select Venue",
                    options=venues,
                    index=0 if not st.session_state.selected_venue else venues.index(st.session_state.selected_venue) if st.session_state.selected_venue in venues else 0,
                    key="venue_selector"
                )
                
                if st.session_state.selected_venue:
                    # Get available dates for selected venue
                    venue_dates_df = st.session_state.venue_df[
                        st.session_state.venue_df['VENUE'] == st.session_state.selected_venue
                    ].copy()
                    
                    if not venue_dates_df.empty:
                        # Group by date and get shifts
                        date_shifts = {}
                        for date in venue_dates_df['DATE'].unique():
                            shifts = venue_dates_df[venue_dates_df['DATE'] == date]['SHIFT'].unique()
                            date_shifts[date] = list(shifts)
                        
                        # Display date and shift selection
                        st.write("Select Dates and Shifts:")
                        
                        selected_dates = {}
                        for date, shifts in sorted(date_shifts.items()):
                            col_date, col_shifts = st.columns([1, 3])
                            with col_date:
                                select_date = st.checkbox(date, key=f"date_{date}")
                            with col_shifts:
                                if select_date:
                                    selected_shifts = []
                                    for shift in shifts:
                                        if st.checkbox(shift, key=f"shift_{date}_{shift}", value=True):
                                            selected_shifts.append(shift)
                                    if selected_shifts:
                                        selected_dates[date] = selected_shifts
                        
                        st.session_state.selected_dates = selected_dates
                        
                        # Step 4: IO Selection
                        st.divider()
                        st.subheader("Step 4: Select Centre Coordinator")
                        
                        if not st.session_state.io_df.empty:
                            # Filter IOs by venue centre code
                            venue_row = venue_dates_df.iloc[0]
                            centre_code = str(venue_row['CENTRE_CODE']).zfill(4)
                            
                            filtered_io = st.session_state.io_df[
                                st.session_state.io_df['CENTRE_CODE'].astype(str).str.zfill(4).str.startswith(centre_code[:4])
                            ]
                            
                            if filtered_io.empty:
                                filtered_io = st.session_state.io_df
                                st.warning(f"No IOs found with matching centre code. Showing all IOs.")
                            
                            # Search box
                            search_term = st.text_input("🔍 Search Centre Coordinator", "")
                            if search_term:
                                filtered_io = filtered_io[
                                    (filtered_io['NAME'].str.contains(search_term, case=False, na=False)) |
                                    (filtered_io['AREA'].str.contains(search_term, case=False, na=False))
                                ]
                            
                            # Display IO list with allocation status
                            io_list = []
                            for _, row in filtered_io.iterrows():
                                io_name = row['NAME']
                                area = row['AREA']
                                
                                # Check existing allocations
                                existing_allocations = [
                                    a for a in st.session_state.allocation 
                                    if a['IO Name'] == io_name and a.get('Exam') == st.session_state.current_exam_key
                                ]
                                
                                status = "🟢 Available"
                                if existing_allocations:
                                    current_venue_allocations = [
                                        a for a in existing_allocations 
                                        if a['Venue'] == st.session_state.selected_venue and a['Role'] == st.session_state.selected_role
                                    ]
                                    if current_venue_allocations:
                                        status = "🔴 Already allocated to this venue"
                                    else:
                                        status = "🟡 Allocated to other venues"
                                
                                io_list.append({
                                    "Name": io_name,
                                    "Area": area,
                                    "Status": status
                                })
                            
                            if io_list:
                                io_df_display = pd.DataFrame(io_list)
                                selected_io_index = st.dataframe(
                                    io_df_display,
                                    use_container_width=True,
                                    hide_index=True
                                )
                                
                                # IO selection
                                io_names = filtered_io['NAME'].tolist()
                                if io_names:
                                    selected_io = st.selectbox(
                                        "Select Centre Coordinator",
                                        options=io_names,
                                        key="io_selector"
                                    )
                                    
                                    # Allocation button
                                    if st.button("✅ Allocate Selected IO to Dates", use_container_width=True):
                                        if not selected_dates:
                                            st.error("Please select at least one date and shift")
                                        else:
                                            # Get allocation reference
                                            ref_data = get_allocation_reference(st.session_state.selected_role)
                                            if ref_data:
                                                # Perform allocation
                                                allocation_count = 0
                                                conflicts = []
                                                
                                                io_data = filtered_io[filtered_io['NAME'] == selected_io].iloc[0]
                                                area = io_data.get('AREA', '')
                                                
                                                for date, shifts in selected_dates.items():
                                                    for shift in shifts:
                                                        # Check for conflict
                                                        conflict = check_allocation_conflict(
                                                            selected_io, date, shift, 
                                                            st.session_state.selected_venue, 
                                                            st.session_state.selected_role, "IO"
                                                        )
                                                        
                                                        if conflict:
                                                            conflicts.append(conflict)
                                                            continue
                                                        
                                                        # Create allocation
                                                        allocation = {
                                                            'Sl. No.': len(st.session_state.allocation) + 1,
                                                            'Venue': st.session_state.selected_venue,
                                                            'Date': date,
                                                            'Shift': shift,
                                                            'IO Name': selected_io,
                                                            'Area': area,
                                                            'Role': st.session_state.selected_role,
                                                            'Mock Test': st.session_state.mock_test_mode,
                                                            'Exam': st.session_state.current_exam_key,
                                                            'Order No.': ref_data['order_no'],
                                                            'Page No.': ref_data['page_no'],
                                                            'Reference Remarks': ref_data.get('remarks', '')
                                                        }
                                                        st.session_state.allocation.append(allocation)
                                                        allocation_count += 1
                                                
                                                if conflicts:
                                                    st.error(f"Allocation conflicts: {', '.join(conflicts[:3])}")
                                                
                                                if allocation_count > 0:
                                                    save_data()
                                                    st.success(f"Allocated {selected_io} to {allocation_count} shifts!")
                                                    st.rerun()
                                            else:
                                                st.warning("Allocation cancelled - no reference provided")
                            else:
                                st.warning("No Centre Coordinators found matching the search criteria")
                        else:
                            st.warning("Please load Centre Coordinator master data first")
                    else:
                        st.warning("No date information found for selected venue")
                else:
                    st.warning("Please select a venue")
            else:
                st.warning("Please load venue data first")
            
            # Display current allocations
            st.divider()
            st.subheader("Current Allocations")
            
            if st.session_state.allocation:
                alloc_df = pd.DataFrame(st.session_state.allocation)
                st.dataframe(
                    alloc_df[['Sl. No.', 'Venue', 'Date', 'Shift', 'IO Name', 'Area', 'Role', 'Mock Test']],
                    use_container_width=True
                )
                
                col_del1, col_del2 = st.columns(2)
                with col_del1:
                    if st.button("🗑️ Delete Last Entry", use_container_width=True):
                        if st.session_state.allocation:
                            # Ask for deletion reference
                            del_ref = ask_for_deletion_reference(st.session_state.allocation[-1]['Role'], 1)
                            if del_ref:
                                # Add to deleted records
                                deleted_entry = st.session_state.allocation[-1].copy()
                                deleted_entry['Deletion Reason'] = del_ref['reason']
                                deleted_entry['Deletion Order No.'] = del_ref['order_no']
                                deleted_entry['Deletion Timestamp'] = datetime.now().isoformat()
                                deleted_entry['Type'] = 'IO'
                                st.session_state.deleted_records.append(deleted_entry)
                                
                                st.session_state.allocation.pop()
                                save_data()
                                st.success("Last entry deleted!")
                                st.rerun()
                
                with col_del2:
                    if st.button("🗑️ Bulk Delete", use_container_width=True):
                        open_bulk_delete_window()
            else:
                st.info("No allocations yet")
    
    # Tab 3: EY Personnel Allocation
    with tab3:
        st.header("EY Personnel Allocation")
        
        if not st.session_state.current_exam_key:
            st.warning("Please select or create an exam first from the Exam Management tab")
        else:
            st.session_state.ey_allocation_mode = st.checkbox(
                "Enable EY Personnel Allocation Mode",
                value=st.session_state.ey_allocation_mode,
                key="ey_mode_checkbox"
            )
            
            if st.session_state.ey_allocation_mode:
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    st.subheader("Step 1: Load EY Personnel Master")
                    
                    ey_file = st.file_uploader(
                        "Load EY Personnel Master",
                        type=["xlsx", "xls"],
                        key="ey_master_upload"
                    )
                    
                    if ey_file:
                        try:
                            st.session_state.ey_df = pd.read_excel(ey_file)
                            st.session_state.ey_df.columns = [str(col).strip().upper() for col in st.session_state.ey_df.columns]
                            
                            required_cols = ["NAME"]
                            missing_cols = [col for col in required_cols if col not in st.session_state.ey_df.columns]
                            
                            if missing_cols:
                                st.error(f"Missing required columns: {', '.join(missing_cols)}")
                            else:
                                optional_cols = ["MOBILE", "EMAIL", "ID_NUMBER", "DESIGNATION", "DEPARTMENT"]
                                for col in optional_cols:
                                    if col not in st.session_state.ey_df.columns:
                                        st.session_state.ey_df[col] = ""
                                
                                st.session_state.ey_df['NAME'] = st.session_state.ey_df['NAME'].astype(str).str.strip()
                                st.success(f"Loaded {len(st.session_state.ey_df)} EY Personnel records")
                        except Exception as e:
                            st.error(f"Error loading file: {str(e)}")
                    
                    # EY Rate
                    st.session_state.remuneration_rates['ey_personnel'] = st.number_input(
                        "EY Rate per Day (₹)",
                        min_value=0,
                        value=st.session_state.remuneration_rates['ey_personnel'],
                        key="ey_rate_input"
                    )
                    
                    if st.button("💾 Save EY Rate", use_container_width=True):
                        save_data()
                        st.success("EY rate saved!")
                
                with col2:
                    st.subheader("Step 2: Configuration")
                    
                    # Select venues for EY allocation
                    if not st.session_state.venue_df.empty:
                        venues = sorted(st.session_state.venue_df['VENUE'].dropna().unique())
                        st.session_state.selected_ey_venues = st.multiselect(
                            "Select Venues for EY Allocation",
                            options=venues,
                            default=venues[:min(3, len(venues))] if venues else []
                        )
                    
                    if st.button("📍 Select All Venues", use_container_width=True):
                        if not st.session_state.venue_df.empty:
                            venues = sorted(st.session_state.venue_df['VENUE'].dropna().unique())
                            st.session_state.selected_ey_venues = venues
                            st.rerun()
                
                st.divider()
                
                # Step 3: Select EY Personnel
                st.subheader("Step 3: Select EY Personnel")
                
                if not st.session_state.ey_df.empty:
                    # Search EY personnel
                    ey_search = st.text_input("🔍 Search EY Personnel", "")
                    
                    if ey_search:
                        filtered_ey = st.session_state.ey_df[
                            (st.session_state.ey_df['NAME'].str.contains(ey_search, case=False, na=False)) |
                            (st.session_state.ey_df['MOBILE'].str.contains(ey_search, case=False, na=False)) |
                            (st.session_state.ey_df['EMAIL'].str.contains(ey_search, case=False, na=False))
                        ]
                    else:
                        filtered_ey = st.session_state.ey_df
                    
                    if not filtered_ey.empty:
                        # Display EY personnel list
                        ey_list = []
                        for _, row in filtered_ey.iterrows():
                            display_text = f"{row['NAME']}"
                            if pd.notna(row.get('MOBILE')) and row['MOBILE']:
                                display_text += f" | Mobile: {row['MOBILE']}"
                            if pd.notna(row.get('EMAIL')) and row['EMAIL']:
                                display_text += f" | Email: {row['EMAIL']}"
                            ey_list.append(display_text)
                        
                        selected_ey_display = st.selectbox(
                            "Select EY Personnel",
                            options=ey_list,
                            key="ey_person_selector"
                        )
                        
                        if selected_ey_display:
                            ey_name = selected_ey_display.split("|")[0].strip()
                            
                            # Step 4: Select Dates
                            st.subheader("Step 4: Select Dates")
                            
                            if not st.session_state.venue_df.empty and st.session_state.selected_ey_venues:
                                # Get unique dates from selected venues
                                all_dates = set()
                                for venue in st.session_state.selected_ey_venues:
                                    venue_dates = st.session_state.venue_df[
                                        st.session_state.venue_df['VENUE'] == venue
                                    ]['DATE'].unique()
                                    all_dates.update(venue_dates)
                                
                                if all_dates:
                                    selected_ey_dates = st.multiselect(
                                        "Select Dates",
                                        options=sorted(all_dates),
                                        key="ey_date_selector"
                                    )
                                    
                                    # Get shifts for selected dates
                                    selected_shifts = {}
                                    for date in selected_ey_dates:
                                        shifts = st.multiselect(
                                            f"Shifts for {date}",
                                            options=["Morning", "Afternoon", "Evening"],
                                            default=["Morning", "Afternoon", "Evening"],
                                            key=f"ey_shifts_{date}"
                                        )
                                        if shifts:
                                            selected_shifts[date] = shifts
                                    
                                    # Allocation button
                                    if st.button("✅ Allocate EY Personnel", use_container_width=True):
                                        if not selected_shifts:
                                            st.error("Please select at least one date and shift")
                                        elif not st.session_state.selected_ey_venues:
                                            st.error("Please select at least one venue")
                                        else:
                                            # Get allocation reference
                                            ref_data = get_allocation_reference("EY Personnel")
                                            if ref_data:
                                                # Get EY details
                                                ey_details = {}
                                                ey_match = st.session_state.ey_df[
                                                    st.session_state.ey_df['NAME'].str.strip() == ey_name
                                                ]
                                                if not ey_match.empty:
                                                    ey_row = ey_match.iloc[0]
                                                    ey_details = {
                                                        'Mobile': ey_row.get('MOBILE', ''),
                                                        'Email': ey_row.get('EMAIL', ''),
                                                        'ID_Number': ey_row.get('ID_NUMBER', ''),
                                                        'Designation': ey_row.get('DESIGNATION', ''),
                                                        'Department': ey_row.get('DEPARTMENT', '')
                                                    }
                                                
                                                # Perform allocation
                                                allocation_count = 0
                                                conflicts = []
                                                
                                                for venue in st.session_state.selected_ey_venues:
                                                    for date, shifts in selected_shifts.items():
                                                        for shift in shifts:
                                                            # Check for conflict
                                                            conflict = check_allocation_conflict(
                                                                ey_name, date, shift, venue, "", "EY"
                                                            )
                                                            
                                                            if conflict:
                                                                conflicts.append(conflict)
                                                                continue
                                                            
                                                            # Create allocation
                                                            allocation = {
                                                                'Sl. No.': len(st.session_state.ey_allocation) + 1,
                                                                'Venue': venue,
                                                                'Date': date,
                                                                'Shift': shift,
                                                                'EY Personnel': ey_name,
                                                                'Mobile': ey_details.get('Mobile', ''),
                                                                'Email': ey_details.get('Email', ''),
                                                                'ID Number': ey_details.get('ID_Number', ''),
                                                                'Designation': ey_details.get('Designation', ''),
                                                                'Department': ey_details.get('Department', ''),
                                                                'Mock Test': False,  # EY allocations typically not for mock tests
                                                                'Exam': st.session_state.current_exam_key,
                                                                'Rate (₹)': st.session_state.remuneration_rates['ey_personnel'],
                                                                'Order No.': ref_data['order_no'],
                                                                'Page No.': ref_data['page_no'],
                                                                'Reference Remarks': ref_data.get('remarks', '')
                                                            }
                                                            st.session_state.ey_allocation.append(allocation)
                                                            allocation_count += 1
                                                
                                                if conflicts:
                                                    st.error(f"Allocation conflicts: {', '.join(conflicts[:3])}")
                                                
                                                if allocation_count > 0:
                                                    save_data()
                                                    st.success(f"Allocated {ey_name} to {allocation_count} shifts across {len(st.session_state.selected_ey_venues)} venues!")
                                                    st.rerun()
                                            else:
                                                st.warning("Allocation cancelled - no reference provided")
                                else:
                                    st.warning("No dates found for selected venues")
                            else:
                                st.warning("Please select venues first")
                    else:
                        st.warning("No EY personnel found matching search criteria")
                else:
                    st.warning("Please load EY Personnel master data first")
                
                # Display EY allocations
                st.divider()
                st.subheader("Current EY Allocations")
                
                if st.session_state.ey_allocation:
                    ey_alloc_df = pd.DataFrame(st.session_state.ey_allocation)
                    st.dataframe(
                        ey_alloc_df[['Sl. No.', 'Venue', 'Date', 'Shift', 'EY Personnel', 'Mobile', 'Email', 'Designation']],
                        use_container_width=True
                    )
                    
                    col_del1, col_del2 = st.columns(2)
                    with col_del1:
                        if st.button("🗑️ Delete Last EY Entry", use_container_width=True, key="del_last_ey"):
                            if st.session_state.ey_allocation:
                                # Ask for deletion reference
                                del_ref = ask_for_deletion_reference("EY Personnel", 1)
                                if del_ref:
                                    # Add to deleted records
                                    deleted_entry = st.session_state.ey_allocation[-1].copy()
                                    deleted_entry['Deletion Reason'] = del_ref['reason']
                                    deleted_entry['Deletion Order No.'] = del_ref['order_no']
                                    deleted_entry['Deletion Timestamp'] = datetime.now().isoformat()
                                    deleted_entry['Type'] = 'EY Personnel'
                                    st.session_state.deleted_records.append(deleted_entry)
                                    
                                    st.session_state.ey_allocation.pop()
                                    save_data()
                                    st.success("Last EY entry deleted!")
                                    st.rerun()
                    
                    with col_del2:
                        if st.button("🗑️ Bulk Delete EY", use_container_width=True):
                            open_ey_bulk_delete_window()
                else:
                    st.info("No EY allocations yet")
            else:
                st.info("Enable EY Personnel Allocation Mode to allocate EY personnel")
    
    # Tab 4: Reports & Export
    with tab4:
        st.header("Reports & Export")
        
        if not st.session_state.current_exam_key:
            st.warning("Please select or create an exam first")
        else:
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Export Options")
                
                if st.button("📊 Export Allocations Report", use_container_width=True):
                    export_allocations_report()
                
                if st.button("💰 Export Remuneration Report", use_container_width=True):
                    export_remuneration_report()
                
                if st.button("📋 Export Summary Report", use_container_width=True):
                    export_summary_report()
            
            with col2:
                st.subheader("Quick Reports")
                
                if st.button("👥 Centre Coordinator Summary", use_container_width=True):
                    show_io_summary()
                
                if st.button("👁️ EY Personnel Summary", use_container_width=True):
                    show_ey_summary()
                
                if st.button("📅 Date-wise Summary", use_container_width=True):
                    show_date_summary()
            
            st.divider()
            
            # Current Statistics
            st.subheader("Current Statistics")
            
            col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
            
            with col_stat1:
                total_io = len(st.session_state.allocation)
                st.metric("Centre Coordinator Allocations", total_io)
            
            with col_stat2:
                total_ey = len(st.session_state.ey_allocation)
                st.metric("EY Personnel Allocations", total_ey)
            
            with col_stat3:
                if st.session_state.allocation:
                    unique_ios = len(set(a['IO Name'] for a in st.session_state.allocation))
                    st.metric("Unique Centre Coordinators", unique_ios)
                else:
                    st.metric("Unique Centre Coordinators", 0)
            
            with col_stat4:
                if st.session_state.ey_allocation:
                    unique_ey = len(set(a['EY Personnel'] for a in st.session_state.ey_allocation))
                    st.metric("Unique EY Personnel", unique_ey)
                else:
                    st.metric("Unique EY Personnel", 0)
            
            # Recent Activity
            st.subheader("Recent Activity")
            
            recent_activity = []
            if st.session_state.allocation:
                for alloc in st.session_state.allocation[-5:]:
                    recent_activity.append({
                        "Type": "Centre Coordinator",
                        "Name": alloc['IO Name'],
                        "Venue": alloc['Venue'],
                        "Date": alloc['Date'],
                        "Shift": alloc['Shift'],
                        "Role": alloc['Role']
                    })
            
            if st.session_state.ey_allocation:
                for alloc in st.session_state.ey_allocation[-5:]:
                    recent_activity.append({
                        "Type": "EY Personnel",
                        "Name": alloc['EY Personnel'],
                        "Venue": alloc['Venue'],
                        "Date": alloc['Date'],
                        "Shift": alloc['Shift'],
                        "Role": "EY Supervisor"
                    })
            
            if recent_activity:
                recent_df = pd.DataFrame(recent_activity[-10:])  # Show last 10
                st.dataframe(recent_df, use_container_width=True)
            else:
                st.info("No recent activity")
    
    # Tab 5: Settings & Tools
    with tab5:
        st.header("Settings & Tools")
        
        tab_settings, tab_backup, tab_help = st.tabs(["⚙️ Settings", "💾 Backup", "❓ Help"])
        
        with tab_settings:
            st.subheader("Application Settings")
            
            # Data Management
            col_set1, col_set2 = st.columns(2)
            
            with col_set1:
                if st.button("🔄 Reset All Data", use_container_width=True):
                    if st.checkbox("I confirm I want to reset ALL data"):
                        st.session_state.allocation = []
                        st.session_state.ey_allocation = []
                        st.session_state.exam_data = {}
                        st.session_state.current_exam_key = ""
                        st.session_state.exam_name = ""
                        st.session_state.exam_year = ""
                        st.session_state.deleted_records = []
                        st.session_state.allocation_references = {}
                        save_data()
                        st.success("All data reset successfully!")
                        st.rerun()
            
            with col_set2:
                if st.button("🗑️ Clear Deleted Records", use_container_width=True):
                    if st.checkbox("I confirm I want to clear ALL deleted records"):
                        st.session_state.deleted_records = []
                        save_data()
                        st.success("Deleted records cleared!")
                        st.rerun()
            
            # System Information
            st.subheader("System Information")
            
            info_col1, info_col2 = st.columns(2)
            
            with info_col1:
                st.write("**Data Files:**")
                st.write(f"- Config: {'✅' if CONFIG_FILE.exists() else '❌'}")
                st.write(f"- Exam Data: {'✅' if DATA_FILE.exists() else '❌'}")
                st.write(f"- References: {'✅' if REFERENCE_FILE.exists() else '❌'}")
                st.write(f"- Deleted Records: {'✅' if DELETED_RECORDS_FILE.exists() else '❌'}")
            
            with info_col2:
                st.write("**Current Data:**")
                st.write(f"- Exams: {len(st.session_state.exam_data)}")
                st.write(f"- Total Allocations: {len(st.session_state.allocation) + len(st.session_state.ey_allocation)}")
                st.write(f"- Deleted Records: {len(st.session_state.deleted_records)}")
                st.write(f"- References: {sum(len(refs) for refs in st.session_state.allocation_references.values())}")
        
        with tab_backup:
            st.subheader("Backup & Restore")
            
            col_back1, col_back2 = st.columns(2)
            
            with col_back1:
                if st.button("💾 Create Backup", use_container_width=True):
                    backup_file = create_backup()
                    if backup_file:
                        st.success(f"Backup created: {backup_file.name}")
                    else:
                        st.error("Failed to create backup")
                
                # List existing backups
                BACKUP_DIR.mkdir(exist_ok=True)
                backup_files = list(BACKUP_DIR.glob("*.json"))
                
                if backup_files:
                    st.write("**Available Backups:**")
                    for i, backup_file in enumerate(sorted(backup_files, reverse=True)[:5]):
                        st.write(f"{i+1}. {backup_file.name} ({backup_file.stat().st_size} bytes)")
            
            with col_back2:
                if backup_files:
                    selected_backup = st.selectbox(
                        "Select Backup to Restore",
                        options=[f.name for f in sorted(backup_files, reverse=True)],
                        key="backup_selector"
                    )
                    
                    if st.button("🔄 Restore Backup", use_container_width=True):
                        if st.checkbox("I confirm I want to restore from backup (this will overwrite current data)"):
                            backup_path = BACKUP_DIR / selected_backup
                            if restore_from_backup(backup_path):
                                st.success("Backup restored successfully!")
                                st.rerun()
                            else:
                                st.error("Failed to restore backup")
        
        with tab_help:
            st.subheader("Help & Documentation")
            
            st.write("""
            ### User Guide
            
            **1. Exam Management**
            - Create a new exam or select an existing one
            - Set allocation references for each role
            - View and manage all references
            
            **2. Centre Coordinator Allocation**
            - Load Centre Coordinator and Venue master files
            - Select venue, dates, and shifts
            - Search and select Centre Coordinators
            - Allocate with proper references
            
            **3. EY Personnel Allocation**
            - Load EY Personnel master file
            - Select multiple venues and dates
            - Allocate EY personnel with references
            
            **4. Reports & Export**
            - Export allocation reports in Excel format
            - Generate remuneration reports
            - View summary statistics
            
            **5. Settings & Tools**
            - Backup and restore data
            - Reset application data
            - View system information
            
            ### Tips
            - Always set allocation references before allocating
            - Use the search functionality to find personnel quickly
            - Regularly backup your data
            - Check for allocation conflicts before finalizing
            """)
            
            st.divider()
            st.write("**Designed by Bijay Paswan**")
            st.caption("Version 1.0 | Staff Selection Commission (ER), Kolkata")

# Helper Functions
def create_backup(exam_key=None):
    """Create a backup of exam data"""
    try:
        BACKUP_DIR.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if exam_key:
            backup_file = BACKUP_DIR / f"backup_{exam_key.replace(' ', '_').replace('-', '_')}_{timestamp}.json"
        else:
            backup_file = BACKUP_DIR / f"full_backup_{timestamp}.json"
        
        with open(backup_file, 'w') as f:
            json.dump(st.session_state.exam_data, f, indent=4, default=str)
        
        logging.info(f"Created backup: {backup_file}")
        return backup_file
    except Exception as e:
        logging.error(f"Error creating backup: {str(e)}")
        return None

def restore_from_backup(backup_file):
    """Restore exam data from backup"""
    try:
        with open(backup_file, 'r') as f:
            restored_data = json.load(f)
        
        st.session_state.exam_data = restored_data
        
        # Clear current allocations
        st.session_state.allocation = []
        st.session_state.ey_allocation = []
        st.session_state.current_exam_key = ""
        st.session_state.exam_name = ""
        st.session_state.exam_year = ""
        
        save_data()
        logging.info(f"Restored data from backup: {backup_file}")
        return True
    except Exception as e:
        logging.error(f"Error restoring from backup: {str(e)}")
        return False

def get_allocation_reference(allocation_type):
    """Get allocation reference with dialog"""
    exam_key = st.session_state.current_exam_key
    if not exam_key:
        st.warning("Please select or create an exam first")
        return None
    
    if exam_key not in st.session_state.allocation_references:
        st.session_state.allocation_references[exam_key] = {}
    
    role_key = allocation_type
    
    # Check if reference exists
    if role_key in st.session_state.allocation_references[exam_key]:
        existing_ref = st.session_state.allocation_references[exam_key][role_key]
        
        # Show existing reference and ask user
        st.info(f"Existing reference found for {allocation_type}:")
        st.write(f"Order No.: {existing_ref.get('order_no', 'N/A')}")
        st.write(f"Page No.: {existing_ref.get('page_no', 'N/A')}")
        
        col_ref1, col_ref2 = st.columns(2)
        with col_ref1:
            use_existing = st.button(f"Use Existing Reference", key=f"use_existing_{role_key}")
        with col_ref2:
            create_new = st.button(f"Create New Reference", key=f"create_new_{role_key}")
        
        if use_existing:
            return existing_ref
        elif create_new:
            return ask_for_allocation_reference(allocation_type)
    else:
        return ask_for_allocation_reference(allocation_type)
    
    return None

def ask_for_allocation_reference(allocation_type):
    """Dialog to enter allocation reference"""
    with st.expander(f"Enter Reference for {allocation_type}", expanded=True):
        order_no = st.text_input("Order No.:", key=f"order_no_{allocation_type}")
        page_no = st.text_input("Page No.:", key=f"page_no_{allocation_type}")
        remarks = st.text_area("Remarks (Optional):", key=f"remarks_{allocation_type}")
        
        col_sub1, col_sub2 = st.columns(2)
        with col_sub1:
            if st.button("✅ Save Reference", key=f"save_ref_{allocation_type}"):
                if order_no and page_no:
                    exam_key = st.session_state.current_exam_key
                    if exam_key not in st.session_state.allocation_references:
                        st.session_state.allocation_references[exam_key] = {}
                    
                    st.session_state.allocation_references[exam_key][allocation_type] = {
                        'order_no': order_no,
                        'page_no': page_no,
                        'remarks': remarks,
                        'timestamp': datetime.now().isoformat(),
                        'allocation_type': allocation_type
                    }
                    
                    save_data()
                    st.success("Reference saved!")
                    return st.session_state.allocation_references[exam_key][allocation_type]
                else:
                    st.error("Please enter both Order No. and Page No.")
        with col_sub2:
            if st.button("❌ Cancel", key=f"cancel_ref_{allocation_type}"):
                return None
    
    return None

def ask_for_deletion_reference(allocation_type, entries_count):
    """Dialog to enter deletion reference"""
    with st.expander(f"Deletion Reference for {entries_count} {allocation_type} allocation(s)", expanded=True):
        order_no = st.text_input("Deletion Order No.:", key=f"del_order_no_{allocation_type}")
        reason = st.text_area("Deletion Reason:", key=f"del_reason_{allocation_type}")
        
        col_sub1, col_sub2 = st.columns(2)
        with col_sub1:
            if st.button("✅ Confirm Deletion", key=f"confirm_del_{allocation_type}"):
                if order_no and reason:
                    return {
                        'order_no': order_no,
                        'reason': reason,
                        'confirmed': True
                    }
                else:
                    st.error("Please enter both Order No. and Deletion Reason")
        with col_sub2:
            if st.button("❌ Cancel", key=f"cancel_del_{allocation_type}"):
                return None
    
    return None

def check_allocation_conflict(person_name, date, shift, venue, role, allocation_type):
    """Check for allocation conflicts"""
    if allocation_type == "IO":
        # Check for duplicate allocation
        duplicate = any(
            alloc['IO Name'] == person_name and 
            alloc['Date'] == date and 
            alloc['Shift'] == shift and 
            alloc['Venue'] == venue and 
            alloc['Role'] == role
            for alloc in st.session_state.allocation
        )
        if duplicate:
            return f"Duplicate allocation found! {person_name} is already allocated to {venue} on {date} ({shift}) as {role}."
        
        # For Centre Coordinator: Cannot be assigned to multiple venues on same date and shift
        if role == "Centre Coordinator":
            conflict = any(
                alloc['IO Name'] == person_name and 
                alloc['Date'] == date and 
                alloc['Shift'] == shift and 
                alloc['Venue'] != venue and
                alloc['Role'] == "Centre Coordinator"
                for alloc in st.session_state.allocation
            )
            if conflict:
                existing_venue = next(
                    alloc['Venue'] for alloc in st.session_state.allocation 
                    if alloc['IO Name'] == person_name and 
                       alloc['Date'] == date and 
                       alloc['Shift'] == shift and
                       alloc['Role'] == "Centre Coordinator"
                )
                return f"Centre Coordinator conflict! {person_name} is already allocated to {existing_venue} on {date} ({shift}). Cannot assign to {venue}."
    
    elif allocation_type == "EY":
        # Check for duplicate allocation
        duplicate = any(
            alloc['EY Personnel'] == person_name and 
            alloc['Date'] == date and 
            alloc['Shift'] == shift and 
            alloc['Venue'] == venue
            for alloc in st.session_state.ey_allocation
        )
        if duplicate:
            return f"Duplicate EY allocation found! {person_name} is already allocated to {venue} on {date} ({shift})."
        
        # EY Personnel: Cannot be assigned to multiple venues on same date and shift
        conflict = any(
            alloc['EY Personnel'] == person_name and 
            alloc['Date'] == date and 
            alloc['Shift'] == shift and 
            alloc['Venue'] != venue
            for alloc in st.session_state.ey_allocation
        )
        if conflict:
            existing_venue = next(
                alloc['Venue'] for alloc in st.session_state.ey_allocation 
                if alloc['EY Personnel'] == person_name and 
                   alloc['Date'] == date and 
                   alloc['Shift'] == shift
            )
            return f"EY Personnel conflict! {person_name} is already allocated to {existing_venue} on {date} ({shift}). Cannot assign to {venue}."
    
    return None

def open_bulk_delete_window():
    """Open bulk delete dialog"""
    if not st.session_state.allocation:
        st.warning("No allocations to delete")
        return
    
    with st.expander("Bulk Delete Allocations", expanded=True):
        # Search filter
        search_term = st.text_input("Search allocations:", key="bulk_delete_search")
        
        # Display allocations for selection
        alloc_list = []
        for i, alloc in enumerate(st.session_state.allocation):
            display_text = f"{i+1}. {alloc['IO Name']} | {alloc['Venue']} | {alloc['Date']} {alloc['Shift']} | {alloc['Role']}"
            if not search_term or search_term.lower() in display_text.lower():
                alloc_list.append({
                    "Select": False,
                    "Display": display_text,
                    "Index": i,
                    "Allocation": alloc
                })
        
        if alloc_list:
            # Create DataFrame for editing
            alloc_df = pd.DataFrame(alloc_list)
            
            # Use st.data_editor for selection
            edited_df = st.data_editor(
                alloc_df[['Select', 'Display']],
                use_container_width=True,
                hide_index=True
            )
            
            # Get selected indices
            selected_indices = [i for i, row in enumerate(edited_df['Select']) if row]
            
            if selected_indices:
                # Group by role
                roles_to_delete = {}
                for idx in selected_indices:
                    actual_idx = alloc_list[idx]['Index']
                    alloc = st.session_state.allocation[actual_idx]
                    role = alloc['Role']
                    if role not in roles_to_delete:
                        roles_to_delete[role] = []
                    roles_to_delete[role].append(actual_idx)
                
                if st.button(f"🗑️ Delete {len(selected_indices)} Selected Entries", use_container_width=True):
                    # Ask for deletion reference for each role
                    deletion_refs = {}
                    for role in roles_to_delete.keys():
                        del_ref = ask_for_deletion_reference(role, len(roles_to_delete[role]))
                        if not del_ref:
                            return  # User cancelled
                        deletion_refs[role] = del_ref
                    
                    if deletion_refs:
                        # Delete in reverse order
                        deleted_count = 0
                        deleted_entries = []
                        
                        for role, indices in roles_to_delete.items():
                            del_ref = deletion_refs[role]
                            for idx in sorted(indices, reverse=True):
                                if 0 <= idx < len(st.session_state.allocation):
                                    # Store deletion reference
                                    deleted_entry = st.session_state.allocation[idx].copy()
                                    deleted_entry['Deletion Reason'] = del_ref['reason']
                                    deleted_entry['Deletion Order No.'] = del_ref['order_no']
                                    deleted_entry['Deletion Timestamp'] = datetime.now().isoformat()
                                    deleted_entry['Type'] = 'IO'
                                    deleted_entries.append(deleted_entry)
                                    
                                    st.session_state.allocation.pop(idx)
                                    deleted_count += 1
                        
                        # Add to deleted records
                        st.session_state.deleted_records.extend(deleted_entries)
                        save_data()
                        st.success(f"Deleted {deleted_count} entries!")
                        st.rerun()

def open_ey_bulk_delete_window():
    """Open EY bulk delete dialog"""
    if not st.session_state.ey_allocation:
        st.warning("No EY allocations to delete")
        return
    
    with st.expander("Bulk Delete EY Allocations", expanded=True):
        # Search filter
        search_term = st.text_input("Search EY allocations:", key="ey_bulk_delete_search")
        
        # Display allocations for selection
        alloc_list = []
        for i, alloc in enumerate(st.session_state.ey_allocation):
            display_text = f"{i+1}. {alloc['EY Personnel']} | {alloc['Venue']} | {alloc['Date']} {alloc['Shift']}"
            if not search_term or search_term.lower() in display_text.lower():
                alloc_list.append({
                    "Select": False,
                    "Display": display_text,
                    "Index": i,
                    "Allocation": alloc
                })
        
        if alloc_list:
            # Create DataFrame for editing
            alloc_df = pd.DataFrame(alloc_list)
            
            # Use st.data_editor for selection
            edited_df = st.data_editor(
                alloc_df[['Select', 'Display']],
                use_container_width=True,
                hide_index=True
            )
            
            # Get selected indices
            selected_indices = [i for i, row in enumerate(edited_df['Select']) if row]
            
            if selected_indices:
                if st.button(f"🗑️ Delete {len(selected_indices)} Selected EY Entries", use_container_width=True):
                    # Ask for deletion reference
                    del_ref = ask_for_deletion_reference("EY Personnel", len(selected_indices))
                    
                    if del_ref:
                        # Delete in reverse order
                        deleted_count = 0
                        deleted_entries = []
                        
                        for listbox_idx in sorted(selected_indices, reverse=True):
                            actual_idx = alloc_list[listbox_idx]['Index']
                            if 0 <= actual_idx < len(st.session_state.ey_allocation):
                                # Store deletion reference
                                deleted_entry = st.session_state.ey_allocation[actual_idx].copy()
                                deleted_entry['Deletion Reason'] = del_ref['reason']
                                deleted_entry['Deletion Order No.'] = del_ref['order_no']
                                deleted_entry['Deletion Timestamp'] = datetime.now().isoformat()
                                deleted_entry['Type'] = 'EY Personnel'
                                deleted_entries.append(deleted_entry)
                                
                                st.session_state.ey_allocation.pop(actual_idx)
                                deleted_count += 1
                        
                        # Add to deleted records
                        st.session_state.deleted_records.extend(deleted_entries)
                        save_data()
                        st.success(f"Deleted {deleted_count} EY entries!")
                        st.rerun()

def view_allocation_references():
    """View all allocation references"""
    if not st.session_state.allocation_references:
        st.info("No allocation references found.")
        return
    
    with st.expander("All Allocation References", expanded=True):
        all_refs = []
        for exam_key, roles in st.session_state.allocation_references.items():
            for role, ref in roles.items():
                all_refs.append({
                    "Exam": exam_key,
                    "Role": role,
                    "Order No.": ref.get('order_no', ''),
                    "Page No.": ref.get('page_no', ''),
                    "Timestamp": ref.get('timestamp', ''),
                    "Remarks": ref.get('remarks', '')[:50] + "..." if len(ref.get('remarks', '')) > 50 else ref.get('remarks', '')
                })
        
        if all_refs:
            refs_df = pd.DataFrame(all_refs)
            st.dataframe(refs_df, use_container_width=True)
            
            # Delete options
            st.subheader("Delete References")
            
            col_del1, col_del2, col_del3 = st.columns(3)
            
            with col_del1:
                if st.button("🗑️ Delete Selected", use_container_width=True):
                    st.info("Select references from the table above (coming soon)")
            
            with col_del2:
                if st.button("🗑️ Delete by Exam", use_container_width=True):
                    exams = list(st.session_state.allocation_references.keys())
                    if exams:
                        selected_exam = st.selectbox("Select Exam to Delete:", exams)
                        if st.button("Confirm Delete Exam References"):
                            if selected_exam in st.session_state.allocation_references:
                                del st.session_state.allocation_references[selected_exam]
                                save_data()
                                st.success(f"Deleted all references for {selected_exam}")
                                st.rerun()
            
            with col_del3:
                if st.button("🗑️ Delete All", use_container_width=True):
                    if st.checkbox("I confirm I want to delete ALL references"):
                        st.session_state.allocation_references = {}
                        save_data()
                        st.success("All references deleted!")
                        st.rerun()

def view_deleted_records():
    """View deleted records"""
    if not st.session_state.deleted_records:
        st.info("No deleted records found.")
        return
    
    with st.expander("Deleted Records", expanded=True):
        deleted_list = []
        for record in st.session_state.deleted_records:
            if 'IO Name' in record:
                deleted_list.append({
                    "Type": "Centre Coordinator",
                    "Name": record['IO Name'],
                    "Venue": record['Venue'],
                    "Date": record['Date'],
                    "Shift": record['Shift'],
                    "Role": record.get('Role', ''),
                    "Deletion Order No.": record.get('Deletion Order No.', ''),
                    "Deletion Reason": record.get('Deletion Reason', '')[:50] + "..." if len(record.get('Deletion Reason', '')) > 50 else record.get('Deletion Reason', ''),
                    "Timestamp": record.get('Deletion Timestamp', '')
                })
            else:
                deleted_list.append({
                    "Type": "EY Personnel",
                    "Name": record['EY Personnel'],
                    "Venue": record['Venue'],
                    "Date": record['Date'],
                    "Shift": record['Shift'],
                    "Role": "EY Personnel",
                    "Deletion Order No.": record.get('Deletion Order No.', ''),
                    "Deletion Reason": record.get('Deletion Reason', '')[:50] + "..." if len(record.get('Deletion Reason', '')) > 50 else record.get('Deletion Reason', ''),
                    "Timestamp": record.get('Deletion Timestamp', '')
                })
        
        if deleted_list:
            deleted_df = pd.DataFrame(deleted_list)
            st.dataframe(deleted_df, use_container_width=True)
            
            # Delete options
            st.subheader("Delete Records Permanently")
            
            col_del1, col_del2 = st.columns(2)
            
            with col_del1:
                if st.button("🗑️ Delete Selected", use_container_width=True):
                    st.info("Select records from the table above (coming soon)")
            
            with col_del2:
                if st.button("🗑️ Delete All", use_container_width=True):
                    if st.checkbox("I confirm I want to permanently delete ALL deleted records"):
                        st.session_state.deleted_records = []
                        save_data()
                        st.success("All deleted records permanently deleted!")
                        st.rerun()

def export_allocations_report():
    """Export allocations report"""
    if not st.session_state.allocation and not st.session_state.ey_allocation:
        st.warning("No data to export.")
        return
    
    try:
        # Create Excel writer
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # IO Allocations
            if st.session_state.allocation:
                alloc_df = pd.DataFrame(st.session_state.allocation)
                alloc_df.to_excel(writer, index=False, sheet_name='IO Allocations')
            
            # EY Allocations
            if st.session_state.ey_allocation:
                ey_alloc_df = pd.DataFrame(st.session_state.ey_allocation)
                ey_alloc_df.to_excel(writer, index=False, sheet_name='EY Allocations')
            
            # Deleted Records
            if st.session_state.deleted_records:
                deleted_df = pd.DataFrame(st.session_state.deleted_records)
                deleted_df.to_excel(writer, index=False, sheet_name='Deleted Records')
        
        # Offer download
        st.download_button(
            label="📥 Download Allocation Report",
            data=output.getvalue(),
            file_name=f"Allocation_Report_{st.session_state.current_exam_key.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"Export failed: {str(e)}")

def export_remuneration_report():
    """Export remuneration report"""
    if not st.session_state.allocation and not st.session_state.ey_allocation:
        st.warning("No data to export.")
        return
    
    try:
        # Calculate IO remuneration
        io_remuneration = []
        if st.session_state.allocation:
            alloc_df = pd.DataFrame(st.session_state.allocation)
            for (io_name, date), group in alloc_df.groupby(['IO Name', 'Date']):
                shifts = group['Shift'].nunique()
                is_mock = any(group['Mock Test'])
                
                if is_mock:
                    amount = st.session_state.remuneration_rates['mock_test']
                    shift_type = "Mock Test"
                else:
                    if shifts > 1:
                        amount = st.session_state.remuneration_rates['multiple_shifts']
                        shift_type = "Multiple Shifts"
                    else:
                        amount = st.session_state.remuneration_rates['single_shift']
                        shift_type = "Single Shift"
                
                io_remuneration.append({
                    'IO Name': io_name,
                    'Date': date,
                    'Total Shifts': shifts,
                    'Shift Type': shift_type,
                    'Amount (₹)': amount
                })
        
        # Calculate EY remuneration
        ey_remuneration = []
        if st.session_state.ey_allocation:
            ey_df = pd.DataFrame(st.session_state.ey_allocation)
            for (ey_person, date), group in ey_df.groupby(['EY Personnel', 'Date']):
                amount = st.session_state.remuneration_rates['ey_personnel']
                ey_remuneration.append({
                    'EY Personnel': ey_person,
                    'Date': date,
                    'Rate Type': 'Per Day',
                    'Amount (₹)': amount
                })
        
        # Create Excel writer
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # IO Remuneration
            if io_remuneration:
                io_rem_df = pd.DataFrame(io_remuneration)
                io_rem_df.to_excel(writer, index=False, sheet_name='IO Remuneration')
                
                # IO Summary
                io_summary = []
                for io_name in set(item['IO Name'] for item in io_remuneration):
                    io_data = [item for item in io_remuneration if item['IO Name'] == io_name]
                    total_amount = sum(item['Amount (₹)'] for item in io_data)
                    total_days = len(io_data)
                    
                    io_summary.append({
                        'IO Name': io_name,
                        'Total Days': total_days,
                        'Total Amount (₹)': total_amount
                    })
                
                if io_summary:
                    io_summary_df = pd.DataFrame(io_summary)
                    io_summary_df.to_excel(writer, index=False, sheet_name='IO Summary')
            
            # EY Remuneration
            if ey_remuneration:
                ey_rem_df = pd.DataFrame(ey_remuneration)
                ey_rem_df.to_excel(writer, index=False, sheet_name='EY Remuneration')
                
                # EY Summary
                ey_summary = []
                for ey_person in set(item['EY Personnel'] for item in ey_remuneration):
                    ey_data = [item for item in ey_remuneration if item['EY Personnel'] == ey_person]
                    total_amount = sum(item['Amount (₹)'] for item in ey_data)
                    total_days = len(ey_data)
                    
                    ey_summary.append({
                        'EY Personnel': ey_person,
                        'Total Days': total_days,
                        'Total Amount (₹)': total_amount
                    })
                
                if ey_summary:
                    ey_summary_df = pd.DataFrame(ey_summary)
                    ey_summary_df.to_excel(writer, index=False, sheet_name='EY Summary')
            
            # Rates
            rates_data = [
                {'Category': 'Multiple Shifts', 'Amount (₹)': st.session_state.remuneration_rates['multiple_shifts'], 'Reference': 'Per allocation'},
                {'Category': 'Single Shift', 'Amount (₹)': st.session_state.remuneration_rates['single_shift'], 'Reference': 'Per allocation'},
                {'Category': 'Mock Test', 'Amount (₹)': st.session_state.remuneration_rates['mock_test'], 'Reference': 'Per allocation'},
                {'Category': 'EY Personnel', 'Amount (₹)': st.session_state.remuneration_rates['ey_personnel'], 'Reference': 'Per day'}
            ]
            rates_df = pd.DataFrame(rates_data)
            rates_df.to_excel(writer, index=False, sheet_name='Rates')
        
        # Offer download
        st.download_button(
            label="📥 Download Remuneration Report",
            data=output.getvalue(),
            file_name=f"Remuneration_Report_{st.session_state.current_exam_key.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"Export failed: {str(e)}")

def export_summary_report():
    """Export summary report"""
    if not st.session_state.allocation and not st.session_state.ey_allocation:
        st.warning("No data to export.")
        return
    
    try:
        # Create summary data
        summary_data = []
        
        # IO Summary
        if st.session_state.allocation:
            alloc_df = pd.DataFrame(st.session_state.allocation)
            io_summary = alloc_df.groupby('IO Name').agg({
                'Venue': 'nunique',
                'Date': 'nunique',
                'Shift': 'count'
            }).reset_index()
            io_summary.columns = ['IO Name', 'Unique Venues', 'Unique Dates', 'Total Shifts']
            
            for _, row in io_summary.iterrows():
                summary_data.append({
                    'Type': 'Centre Coordinator',
                    'Name': row['IO Name'],
                    'Unique Venues': row['Unique Venues'],
                    'Unique Dates': row['Unique Dates'],
                    'Total Shifts': row['Total Shifts']
                })
        
        # EY Summary
        if st.session_state.ey_allocation:
            ey_df = pd.DataFrame(st.session_state.ey_allocation)
            ey_summary = ey_df.groupby('EY Personnel').agg({
                'Venue': 'nunique',
                'Date': 'nunique',
                'Shift': 'count'
            }).reset_index()
            ey_summary.columns = ['EY Personnel', 'Unique Venues', 'Unique Dates', 'Total Shifts']
            
            for _, row in ey_summary.iterrows():
                summary_data.append({
                    'Type': 'EY Personnel',
                    'Name': row['EY Personnel'],
                    'Unique Venues': row['Unique Venues'],
                    'Unique Dates': row['Unique Dates'],
                    'Total Shifts': row['Total Shifts']
                })
        
        # Create Excel writer
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Summary
            if summary_data:
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, index=False, sheet_name='Summary')
            
            # Date Summary
            if st.session_state.allocation:
                alloc_df = pd.DataFrame(st.session_state.allocation)
                date_summary = alloc_df.groupby('Date').agg({
                    'Venue': 'nunique',
                    'IO Name': 'nunique',
                    'Shift': 'count'
                }).reset_index()
                date_summary.columns = ['Date', 'Unique Venues', 'Unique IOs', 'Total Shifts']
                date_summary.to_excel(writer, index=False, sheet_name='Date Summary')
        
        # Offer download
        st.download_button(
            label="📥 Download Summary Report",
            data=output.getvalue(),
            file_name=f"Summary_Report_{st.session_state.current_exam_key.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"Export failed: {str(e)}")

def show_io_summary():
    """Show IO summary"""
    if not st.session_state.allocation:
        st.info("No Centre Coordinator allocations yet.")
        return
    
    alloc_df = pd.DataFrame(st.session_state.allocation)
    
    # Group by IO Name
    io_summary = alloc_df.groupby('IO Name').agg({
        'Venue': lambda x: ', '.join(sorted(set(x))),
        'Date': lambda x: ', '.join(sorted(set(x))),
        'Shift': 'count',
        'Role': lambda x: ', '.join(sorted(set(x)))
    }).reset_index()
    
    io_summary.columns = ['IO Name', 'Venues', 'Dates', 'Total Shifts', 'Roles']
    
    st.dataframe(io_summary, use_container_width=True)
    
    # Statistics
    col_stat1, col_stat2, col_stat3 = st.columns(3)
    with col_stat1:
        st.metric("Total IOs", len(io_summary))
    with col_stat2:
        st.metric("Total Shifts", io_summary['Total Shifts'].sum())
    with col_stat3:
        unique_dates = set()
        for dates in io_summary['Dates']:
            unique_dates.update(dates.split(', '))
        st.metric("Unique Dates", len(unique_dates))

def show_ey_summary():
    """Show EY summary"""
    if not st.session_state.ey_allocation:
        st.info("No EY Personnel allocations yet.")
        return
    
    ey_df = pd.DataFrame(st.session_state.ey_allocation)
    
    # Group by EY Personnel
    ey_summary = ey_df.groupby('EY Personnel').agg({
        'Venue': lambda x: ', '.join(sorted(set(x))),
        'Date': lambda x: ', '.join(sorted(set(x))),
        'Shift': 'count'
    }).reset_index()
    
    ey_summary.columns = ['EY Personnel', 'Venues', 'Dates', 'Total Shifts']
    
    st.dataframe(ey_summary, use_container_width=True)
    
    # Statistics
    col_stat1, col_stat2, col_stat3 = st.columns(3)
    with col_stat1:
        st.metric("Total EY Personnel", len(ey_summary))
    with col_stat2:
        st.metric("Total Shifts", ey_summary['Total Shifts'].sum())
    with col_stat3:
        unique_dates = set()
        for dates in ey_summary['Dates']:
            unique_dates.update(dates.split(', '))
        st.metric("Unique Dates", len(unique_dates))

def show_date_summary():
    """Show date summary"""
    if not st.session_state.allocation:
        st.info("No allocations yet.")
        return
    
    alloc_df = pd.DataFrame(st.session_state.allocation)
    
    # Group by Date
    date_summary = alloc_df.groupby('Date').agg({
        'Venue': 'nunique',
        'IO Name': 'nunique',
        'Shift': 'count'
    }).reset_index()
    
    date_summary.columns = ['Date', 'Unique Venues', 'Unique IOs', 'Total Shifts']
    date_summary = date_summary.sort_values('Date')
    
    st.dataframe(date_summary, use_container_width=True)

def add_allocation_reference(allocation_type):
    """Add allocation reference"""
    exam_key = st.session_state.current_exam_key
    if not exam_key:
        st.warning("Please select or create an exam first")
        return
    
    with st.expander(f"Add Reference for {allocation_type}", expanded=True):
        order_no = st.text_input("Order No.:", key=f"add_order_{allocation_type}")
        page_no = st.text_input("Page No.:", key=f"add_page_{allocation_type}")
        remarks = st.text_area("Remarks (Optional):", key=f"add_remarks_{allocation_type}")
        
        if st.button("💾 Save Reference", use_container_width=True):
            if order_no and page_no:
                if exam_key not in st.session_state.allocation_references:
                    st.session_state.allocation_references[exam_key] = {}
                
                st.session_state.allocation_references[exam_key][allocation_type] = {
                    'order_no': order_no,
                    'page_no': page_no,
                    'remarks': remarks,
                    'timestamp': datetime.now().isoformat(),
                    'allocation_type': allocation_type
                }
                
                save_data()
                st.success("Reference added successfully!")
                st.rerun()
            else:
                st.error("Please enter both Order No. and Page No.")

def update_allocation_reference(allocation_type):
    """Update allocation reference"""
    exam_key = st.session_state.current_exam_key
    if not exam_key:
        st.warning("Please select or create an exam first")
        return
    
    if exam_key not in st.session_state.allocation_references or allocation_type not in st.session_state.allocation_references[exam_key]:
        st.warning(f"No existing reference found for {allocation_type}")
        return
    
    with st.expander(f"Update Reference for {allocation_type}", expanded=True):
        existing_ref = st.session_state.allocation_references[exam_key][allocation_type]
        
        order_no = st.text_input("Order No.:", value=existing_ref.get('order_no', ''), key=f"update_order_{allocation_type}")
        page_no = st.text_input("Page No.:", value=existing_ref.get('page_no', ''), key=f"update_page_{allocation_type}")
        remarks = st.text_area("Remarks (Optional):", value=existing_ref.get('remarks', ''), key=f"update_remarks_{allocation_type}")
        
        col_up1, col_up2 = st.columns(2)
        with col_up1:
            if st.button("💾 Update Reference", use_container_width=True):
                if order_no and page_no:
                    st.session_state.allocation_references[exam_key][allocation_type] = {
                        'order_no': order_no,
                        'page_no': page_no,
                        'remarks': remarks,
                        'timestamp': datetime.now().isoformat(),
                        'allocation_type': allocation_type
                    }
                    
                    save_data()
                    st.success("Reference updated successfully!")
                    st.rerun()
                else:
                    st.error("Please enter both Order No. and Page No.")
        
        with col_up2:
            if st.button("🗑️ Delete Reference", use_container_width=True):
                del st.session_state.allocation_references[exam_key][allocation_type]
                if not st.session_state.allocation_references[exam_key]:
                    del st.session_state.allocation_references[exam_key]
                
                save_data()
                st.success("Reference deleted successfully!")
                st.rerun()

if __name__ == "__main__":
    main()
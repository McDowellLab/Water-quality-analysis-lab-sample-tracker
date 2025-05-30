import os
import sys
import pyodbc
import pandas as pd
import customtkinter as ctk
from tkinter import ttk, filedialog, messagebox
import traceback
import tkinter as tk
import datetime
from tkcalendar import Calendar, DateEntry
import calendar

def get_file_path(filename):
    """Get the absolute path to a file based on whether the app is frozen or not."""
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(__file__)
    return os.path.join(application_path, filename)


def get_database_path():
    """Get the path to the Access database."""
    # Define potential database locations in order of preference
    database_filename = "wrrc database 110907.accdb"

    # Check if we're running as an executable or in development
    if getattr(sys, 'frozen', False):
        # Running as compiled exe
        exe_dir = os.path.dirname(sys.executable)

        # 1. First check if the database is in the same folder as the exe
        exe_path = os.path.join(exe_dir, database_filename)
        if os.path.exists(exe_path):
            print(f"Found database in exe directory: {exe_path}")
            return exe_path

        # 2. Then check one directory up from the exe
        parent_path = os.path.join(os.path.dirname(exe_dir), database_filename)
        if os.path.exists(parent_path):
            print(f"Found database in parent directory: {parent_path}")
            return parent_path

        # 3. Return the exe directory path even if the file isn't there yet
        # (user might copy it after starting the application)
        print(f"Database not found, defaulting to exe directory: {exe_dir}")
        return os.path.join(exe_dir, database_filename)
    else:
        # Running in development environment
        # Go up one level from sample_tracking
        base_path = os.path.dirname(os.path.dirname(__file__))
        db_path = os.path.join(base_path, database_filename)
        print(f"Development mode, using database path: {db_path}")
        return db_path

class SampleTrackerApp(ctk.CTk):

    # Now, modify the __init__ method to add the Edit tab
    def __init__(self):
        super().__init__()

        # Set appearance mode and default color theme
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.title("WRRC Sample Tracking Application")
        self.geometry("1200x800")
        print("Initializing application...")

        # Initialize state variables FIRST
        self.analysis_completed_var = ctk.BooleanVar(value=False)
        self.filter_by_date_var = ctk.BooleanVar(value=True)  # Default to filtering by date
        self.percent_date_filtered = 0
        self.years_limit = 1  # Changed from 8 to 1 year

        # Connect to Access database AFTER initializing variables
        self.db_path = get_database_path()
        self.password = "x"
        self.data = self._load_data_from_database()  # This should respect the date filter now

        # Then create UI elements
        self.create_tabview()
        self.create_search_tab()
        self.create_import_tab()
        self.create_edit_tab()
        self.create_calendar_tab()
        # Selected record for editing
        self.selected_record = None
        self.analysis_data = None

        # Start with the search tab showing
        self.tabview.set("Search")

    def _get_db_connection(self):
        """Create a connection to the Access database."""
        try:
            conn_str = (
                r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                f"DBQ={self.db_path};"
                f"PWD={self.password};"
                "Extended Properties='Excel 8.0;IMEX=1;'"
            )

            conn = pyodbc.connect(conn_str, autocommit=True)
            conn.setdecoding(pyodbc.SQL_CHAR, encoding='latin1')
            conn.setdecoding(pyodbc.SQL_WCHAR, encoding='latin1')
            conn.setencoding(encoding='latin1')

            return conn
        except Exception as e:
            print(f"Error connecting to database: {e}")
            messagebox.showerror("Database Error", f"Could not connect to the database: {str(e)}")
            return None

    def create_tabview(self):
        """Create the main tabview for the application."""
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)

        # Add tabs
        self.tabview.add("Search")
        self.tabview.add("Import")
        self.tabview.add("Edit")
        self.tabview.add("Calendar")  # Add the Calendar tab

    def _get_db_connection(self):
        """Create a connection to the Access database."""
        try:
            conn_str = (
                r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                f"DBQ={self.db_path};"
                f"PWD={self.password};"
                "Extended Properties='Excel 8.0;IMEX=1;'"
            )

            conn = pyodbc.connect(conn_str, autocommit=True)
            conn.setdecoding(pyodbc.SQL_CHAR, encoding='latin1')
            conn.setdecoding(pyodbc.SQL_WCHAR, encoding='latin1')
            conn.setencoding(encoding='latin1')

            return conn
        except Exception as e:
            print(f"Error connecting to database: {e}")
            messagebox.showerror("Database Error", f"Could not connect to the database: {str(e)}")
            return None

    def _load_data_from_database(self):
        """Load data from Access database instead of Excel."""
        try:
            # First, verify that the database file exists
            if not os.path.exists(self.db_path):
                print(f"ERROR: Database file not found at {self.db_path}")
                return pd.DataFrame()

            print(f"Database file confirmed at: {self.db_path}")

            # Create a connection string for the Access database with password
            conn_str = (
                r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                f"DBQ={self.db_path};"
                f"PWD={self.password};"
                "Extended Properties='Excel 8.0;IMEX=1;'"  # Added IMEX=1 to handle mixed data
            )

            print(f"Attempting to connect to database with connection string")

            # Connect to the database with explicit encoding settings
            conn = pyodbc.connect(conn_str, autocommit=True)

            # Adjust character encoding behavior
            conn.setdecoding(pyodbc.SQL_CHAR, encoding='latin1')
            conn.setdecoding(pyodbc.SQL_WCHAR, encoding='latin1')
            conn.setencoding(encoding='latin1')

            cursor = conn.cursor()

            # Check if we should limit the initial query by date
            years_limit = getattr(self, 'years_limit', 1)  # Default to 1 year if not set
            limit_by_date = True  # Always filter by date on initial load

            if limit_by_date:
                try:
                    # Calculate cutoff date (X years ago)
                    current_date = datetime.datetime.now()
                    cutoff_date = current_date - datetime.timedelta(days=years_limit * 365)
                    cutoff_date_str = cutoff_date.strftime('%Y-%m-%d')

                    print(f"Initial load: Limiting to samples newer than {cutoff_date_str} (past {years_limit} year)")

                    # Apply the filter directly in SQL for initial load
                    cursor.execute(
                        """
                        SELECT * FROM [WRRC sample info] 
                        WHERE ([Collection_Date] >= ? OR [Collection_Date] IS NULL)
                        """,
                        (cutoff_date_str,)
                    )
                except Exception as filter_err:
                    print(f"Error applying date filter in SQL: {filter_err}")
                    # Fall back to loading all records
                    cursor.execute("SELECT * FROM [WRRC sample info]")
            else:
                # Query all records
                cursor.execute("SELECT * FROM [WRRC sample info]")

            columns = [column[0] for column in cursor.description]

            # Create a safer way to fetch data without encoding issues
            rows = []
            while True:
                try:
                    row = cursor.fetchone()
                    if row is None:
                        break
                    # Convert each value in the row to string safely
                    safe_row = []
                    for val in row:
                        if val is None:
                            safe_row.append("")
                        else:
                            try:
                                safe_row.append(str(val))
                            except:
                                safe_row.append("")
                    rows.append(safe_row)
                except Exception as fetch_err:
                    print(f"Error fetching row: {fetch_err}")
                    continue

            # Create a DataFrame from the safe data
            df = pd.DataFrame(rows, columns=columns)

            cursor.close()
            conn.close()

            print(f"Successfully loaded {len(df)} rows from Access database (filtered by date)")
            return df

        except Exception as e:
            print(f"Error connecting to Access database: {e}")
            print(f"Detailed error info: {str(e)}")
            return pd.DataFrame()

            # Attempt to fall back to Excel as a last resort

            return pd.DataFrame()

    def search_by_sample(self):
        """Search rows by matching text in 'UNH#' or 'Sample_Name' columns."""
        search_term = self.sample_search_entry.get().strip()
        if not search_term:
            print("Sample search term is empty.")
            return

        print(f"Searching for samples matching: '{search_term}'")

        if self.data.empty:
            print("No data available to search")
            return

        # Check if the required columns exist
        sample_cols = ["UNH#", "Sample_Name"]
        available_cols = [col for col in sample_cols if col in self.data.columns]

        if not available_cols:
            print(f"Error: None of the required columns {sample_cols} found in the dataset.")
            return

        # Initialize an empty filtered dataframe
        filtered_data = pd.DataFrame(columns=self.data.columns)

        # Check for exact matches first
        for col in available_cols:
            exact_matches = self.data[self.data[col].astype(str).str.lower() == search_term.lower()]
            filtered_data = pd.concat([filtered_data, exact_matches])

        # If no exact matches, try partial matches
        if filtered_data.empty:
            for col in available_cols:
                partial_matches = self.data[self.data[col].astype(str).str.contains(search_term, case=False, na=False)]
                filtered_data = pd.concat([filtered_data, partial_matches])

        # Remove duplicates in case a row matched in multiple columns
        filtered_data = filtered_data.drop_duplicates()

        # Apply date filter if checkbox is checked
        filtered_data = self.apply_date_filter(filtered_data)

        if filtered_data.empty:
            print("No samples found for:", search_term)
        else:
            print(f"Found {len(filtered_data)} sample(s) for: {search_term}")
            print(f"First match: {filtered_data.iloc[0]['UNH#'] if 'UNH#' in filtered_data.columns else 'N/A'}, "
                  f"{filtered_data.iloc[0]['Sample_Name'] if 'Sample_Name' in filtered_data.columns else 'N/A'}")

        self.populate_treeview(filtered_data)

    def search_by_project(self):
        """
        Search rows by matching text across project-related columns.
        The search text is split into tokens, and a row is returned only if every token
        is found in the combined text of the project's fields.
        """
        search_term = self.project_search_entry.get().strip()
        if not search_term:
            print("Project search term is empty.")
            return

        print(f"Searching for project matching: '{search_term}'")

        if self.data.empty:
            print("No data available to search")
            return

        # Define project-related columns
        project_cols = ["Project", "Sub_Project", "Sub_ProjectA", "Sub_ProjectB"]
        available_cols = [col for col in project_cols if col in self.data.columns]

        if not available_cols:
            print("Error: No project-related columns found in the dataset.")
            return

        # Split search term into tokens
        tokens = search_term.split()

        def row_matches(row):
            combined = " ".join([str(row[col]) for col in available_cols if pd.notna(row[col])])
            combined_lower = combined.lower()
            return all(token.lower() in combined_lower for token in tokens)

        # Filter the data
        mask = self.data.apply(row_matches, axis=1)
        filtered_data = self.data[mask]

        # Apply date filter if checkbox is checked
        filtered_data = self.apply_date_filter(filtered_data)

        if filtered_data.empty:
            print("No project found for:", search_term)
        else:
            print("Found", len(filtered_data), "row(s) matching project search:", search_term)

        self.populate_treeview(filtered_data)

    def refresh_search(self):
        """Refresh search results based on the current filter settings."""
        # If there's an active search, re-run it
        if self.sample_search_entry.get().strip():
            self.search_by_sample()
        elif self.project_search_entry.get().strip():
            self.search_by_project()
        else:
            # Otherwise, show all records (respecting the filter)
            self.show_all()

    def apply_date_filter(self, df):
        """Apply date filter to restrict results to samples less than X years old."""
        if not self.filter_by_date_var.get():
            return df  # Return unfiltered if checkbox is unchecked

        try:
            # Get current date
            current_date = datetime.datetime.now()

            # Get the years limit from the class (default to 1 if not set)
            years_limit = getattr(self, 'years_limit', 1)

            # Calculate cutoff date (X years ago)
            cutoff_date = current_date - datetime.timedelta(days=years_limit * 365)
            cutoff_date_str = cutoff_date.strftime('%Y-%m-%d')

            print(f"Filtering samples newer than {cutoff_date_str}")

            # Check if 'Collection_Date' column exists
            if 'Collection_Date' not in df.columns:
                print("Warning: 'Collection_Date' column not found. Cannot apply date filter.")
                return df

            # Create a copy of the DataFrame to avoid modifying the original
            filtered_df = df.copy()

            # Debug: Print unique date formats in the dataset
            sample_dates = filtered_df['Collection_Date'].dropna().sample(min(10, len(filtered_df)))
            print(f"Sample date formats: {sample_dates.tolist()}")

            # Count valid and non-empty dates before conversion
            non_empty_dates = filtered_df['Collection_Date'].count()
            print(f"Non-empty dates before conversion: {non_empty_dates} out of {len(filtered_df)}")

            # Try to convert dates with more robust handling of formats
            try:
                # First attempt - try to parse with common formats
                filtered_df['Temp_Date'] = pd.to_datetime(filtered_df['Collection_Date'],
                                                          errors='coerce')  # Converts invalid dates to NaT

                # Count successful conversions
                valid_dates = filtered_df['Temp_Date'].count()
                print(f"Successfully converted dates: {valid_dates} out of {non_empty_dates}")

                # Check if we have a very low conversion rate, indicating potential format issues
                if valid_dates < non_empty_dates * 0.5 and non_empty_dates > 10:
                    print("Warning: Low date conversion rate. Attempting alternative formats...")

                    # Try additional date formats
                    for fmt in ['%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d', '%m-%d-%Y', '%d-%m-%Y']:
                        try:
                            filtered_df['Temp_Date_Alt'] = pd.to_datetime(filtered_df['Collection_Date'],
                                                                          errors='coerce',
                                                                          format=fmt)
                            # Count valid conversions with this format
                            valid_alt = filtered_df['Temp_Date_Alt'].count()
                            print(f"Format {fmt}: {valid_alt} valid conversions")

                            # If this format works better, use it
                            if valid_alt > valid_dates:
                                filtered_df['Temp_Date'] = filtered_df['Temp_Date_Alt']
                                valid_dates = valid_alt
                                print(f"Using alternative format {fmt} with {valid_dates} valid conversions")

                            filtered_df.drop('Temp_Date_Alt', axis=1, inplace=True)
                        except:
                            continue
            except Exception as date_err:
                print(f"Error during date conversion: {str(date_err)}")
                # Create an empty Temp_Date column to avoid errors in the filter
                filtered_df['Temp_Date'] = pd.NaT

            # Debug: Check range of dates after conversion
            if filtered_df['Temp_Date'].count() > 0:
                min_date = filtered_df['Temp_Date'].min()
                max_date = filtered_df['Temp_Date'].max()
                print(f"Date range in dataset: {min_date} to {max_date}")
                print(f"Cutoff date: {cutoff_date}")

            # Keep rows where:
            # 1. Date is greater than (newer than) the cutoff date, OR
            # 2. Date is missing/invalid
            # This avoids filtering out records without dates
            mask = (filtered_df['Temp_Date'] > cutoff_date) | (filtered_df['Temp_Date'].isnull())

            # Count how many records would be filtered
            would_filter = len(filtered_df) - len(filtered_df[mask])
            print(f"Would filter out {would_filter} records older than {cutoff_date_str}")

            # Apply filter
            filtered_df = filtered_df[mask]

            # Drop the temporary column
            filtered_df = filtered_df.drop('Temp_Date', axis=1)

            print(f"Date filter applied: {len(df)} rows reduced to {len(filtered_df)} rows")
            return filtered_df

        except Exception as e:
            print(f"Error applying date filter: {str(e)}")
            print(traceback.format_exc())
            return df  # Return original dataframe if there's an error

    def clear_search(self):
        """Clear both search fields and show all records."""
        self.sample_search_entry.delete(0, "end")
        self.project_search_entry.delete(0, "end")
        self.show_all()

    def create_tabview(self):
        """Create the main tabview for the application."""
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)

        # Add tabs
        self.tabview.add("Search")
        self.tabview.add("Import")
        self.tabview.add("Edit")
        self.tabview.add("Calendar")  # Add the Calendar tab

    def edit_selected_due_sample(self, event=None):
        """Handle double-click on a sample in the due samples list."""
        selected_items = self.due_samples_tree.selection()
        if not selected_items:
            return

        # Get the selected item
        item_id = selected_items[0]
        item_values = self.due_samples_tree.item(item_id, "values")

        # Find this sample in the database by UNH#
        unh_id = item_values[0]  # First column is UNH#

        # Query the database for the full record
        conn = self._get_db_connection()
        if not conn:
            return

        cursor = conn.cursor()

        try:
            # Get sample info
            query = "SELECT * FROM [WRRC sample info] WHERE [UNH#] = ?"
            cursor.execute(query, (unh_id,))

            # Get column names
            columns = [column[0] for column in cursor.description]

            # Get the row
            row = cursor.fetchone()

            if row:
                # Create a dictionary from column names and values
                record_dict = {}
                for i, col in enumerate(columns):
                    record_dict[col] = row[i] if i < len(row) else ""

                # Store the selected record
                self.selected_record = record_dict

                # Load analysis data for this record
                self.load_analysis_data(unh_id)

                # Switch to the Edit tab
                self.tabview.set("Edit")

                # Populate the edit form
                self.populate_edit_form()
            else:
                messagebox.showwarning("Sample Not Found", f"No sample record found for UNH# {unh_id}")

        except Exception as e:
            print(f"Error loading sample for editing: {str(e)}")
            traceback.print_exc()
            messagebox.showerror("Error", f"Could not load sample for editing: {str(e)}")
        finally:
            cursor.close()
            conn.close()
    # Add the edit_selected_record method
    def edit_selected_record(self):
        """Edit the selected record from the search results."""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("No Selection", "Please select a record to edit.")
            return

        # Get the selected item
        item_id = selected_items[0]

        # Get the values from the selected item
        item_values = self.tree.item(item_id, "values")

        # Get column names
        columns = self.tree.cget("columns")

        # Create a dictionary from column names and values
        record_dict = {}
        for i, col in enumerate(columns):
            record_dict[col] = item_values[i] if i < len(item_values) else ""

        # Store the selected record
        self.selected_record = record_dict

        # Load analysis data for this record
        unh_id = record_dict.get("UNH#", "")
        if unh_id:
            self.load_analysis_data(unh_id)
        else:
            self.analysis_data = None

        # Switch to the Edit tab
        self.tabview.set("Edit")

        # Populate the edit form
        self.populate_edit_form()

    def create_calendar_tab(self):
        """Create a calendar view tab for visualizing due dates."""
        calendar_tab = self.tabview.tab("Calendar")

        # Create control frame
        control_frame = ctk.CTkFrame(calendar_tab)
        control_frame.pack(fill="x", padx=10, pady=10)

        # Month and Year selection
        today = datetime.date.today()
        month_var = ctk.StringVar(value=calendar.month_name[today.month])
        year_var = ctk.IntVar(value=today.year)

        month_label = ctk.CTkLabel(control_frame, text="Month:")
        month_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        month_options = [calendar.month_name[i] for i in range(1, 13)]
        month_dropdown = ctk.CTkOptionMenu(
            control_frame,
            values=month_options,
            variable=month_var,
            command=lambda x: self.update_calendar_view()
        )
        month_dropdown.grid(row=0, column=1, padx=10, pady=10)

        year_label = ctk.CTkLabel(control_frame, text="Year:")
        year_label.grid(row=0, column=2, padx=10, pady=10, sticky="w")

        # Create years list for dropdown (5 years back, 5 years forward)
        current_year = today.year
        years = list(range(current_year - 5, current_year + 6))  # 5 years before, 5 years after
        year_dropdown = ctk.CTkOptionMenu(
            control_frame,
            values=[str(y) for y in years],
            variable=year_var,
            command=lambda x: self.update_calendar_view()
        )
        year_dropdown.grid(row=0, column=3, padx=10, pady=10)

        # Refresh button
        refresh_btn = ctk.CTkButton(
            control_frame,
            text="Refresh Calendar",
            command=self.update_calendar_view
        )
        refresh_btn.grid(row=0, column=4, padx=20, pady=10)

        # Store the month and year variables
        self.calendar_month_var = month_var
        self.calendar_year_var = year_var

        # Create the calendar frame
        self.calendar_frame = ctk.CTkFrame(calendar_tab)
        self.calendar_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Create samples due frame (shows list of samples due)
        self.samples_due_frame = ctk.CTkFrame(calendar_tab)
        self.samples_due_frame.pack(fill="both", expand=True, padx=10, pady=10)

        samples_due_label = ctk.CTkLabel(
            self.samples_due_frame,
            text="Samples Due on Selected Date",
            font=("Helvetica", 16, "bold")
        )
        samples_due_label.pack(pady=10)

        # Create a treeview to show samples due
        columns = ["UNH#", "Sample_Name", "Collection_Date", "Due_Date", "Project"]
        self.due_samples_tree = ttk.Treeview(
            self.samples_due_frame,
            columns=columns,
            show="headings"
        )

        # Configure columns
        for col in columns:
            self.due_samples_tree.heading(col, text=col)
            self.due_samples_tree.column(col, width=150, minwidth=50)

        self.due_samples_tree.pack(fill="both", expand=True, padx=5, pady=5)

        # Add scrollbars for the treeview
        y_scrollbar = ttk.Scrollbar(
            self.samples_due_frame,
            orient="vertical",
            command=self.due_samples_tree.yview
        )
        y_scrollbar.pack(side="right", fill="y")

        x_scrollbar = ttk.Scrollbar(
            self.samples_due_frame,
            orient="horizontal",
            command=self.due_samples_tree.xview
        )
        x_scrollbar.pack(side="bottom", fill="x")

        self.due_samples_tree.configure(
            yscrollcommand=y_scrollbar.set,
            xscrollcommand=x_scrollbar.set
        )

        # Add double-click event to go to edit
        self.due_samples_tree.bind("<Double-1>", self.edit_selected_due_sample)
        # Initialize the calendar view
        self.initialize_calendar_view()

    def initialize_calendar_view(self):
        """Initialize the calendar view with current month's due dates."""
        # Clear any existing widgets in the frame
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()

        # Get current month and year
        month_name = self.calendar_month_var.get()
        month_num = list(calendar.month_name).index(month_name)
        year = int(self.calendar_year_var.get())

        # Create title
        title_label = ctk.CTkLabel(
            self.calendar_frame,
            text=f"{month_name} {year}",
            font=("Helvetica", 18, "bold")
        )
        title_label.pack(pady=10)

        # Create calendar grid
        days_frame = ctk.CTkFrame(self.calendar_frame)
        days_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Add day headers
        days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        for i, day in enumerate(days):
            day_label = ctk.CTkLabel(
                days_frame,
                text=day,
                font=("Helvetica", 12, "bold")
            )
            day_label.grid(row=0, column=i, padx=5, pady=5, sticky="nsew")

        # Get calendar for the month
        cal = calendar.monthcalendar(year, month_num)

        # Get samples due this month
        due_samples = self.get_samples_due_in_month(year, month_num)

        # Dictionary to store samples by day
        samples_by_day = {}
        for sample in due_samples:
            due_date = sample.get('Due_Date')
            if due_date:
                try:
                    # Parse the due date
                    if isinstance(due_date, str):
                        due_date_obj = datetime.datetime.strptime(due_date, '%Y-%m-%d').date()
                    else:
                        due_date_obj = due_date

                    # Group samples by day
                    day = due_date_obj.day
                    if day not in samples_by_day:
                        samples_by_day[day] = []

                    # Add UNH# to the list for this day
                    samples_by_day[day].append(sample)
                except Exception as e:
                    print(f"Error parsing due date: {due_date} - {str(e)}")

        # Create day cells
        for week_idx, week in enumerate(cal):
            for day_idx, day in enumerate(week):
                if day == 0:
                    # Empty cell for days not in this month
                    empty_frame = ctk.CTkFrame(days_frame, fg_color="transparent")
                    empty_frame.grid(row=week_idx + 1, column=day_idx, padx=2, pady=2, sticky="nsew")
                else:
                    # Create a frame for the day
                    day_frame = ctk.CTkFrame(days_frame)
                    day_frame.grid(row=week_idx + 1, column=day_idx, padx=2, pady=2, sticky="nsew")

                    # Add day number
                    day_num_label = ctk.CTkLabel(
                        day_frame,
                        text=str(day),
                        font=("Helvetica", 12)
                    )
                    day_num_label.pack(pady=(5, 0))

                    # Check if there are samples due on this day
                    if day in samples_by_day:
                        day_samples = samples_by_day[day]
                        count = len(day_samples)

                        # Create a mini-list of UNH IDs (limited to first 3 if many)
                        if count > 0:
                            if count <= 3:
                                # Show all UNH IDs if 3 or fewer
                                for sample in day_samples:
                                    # When creating the UNH# labels in the calendar view:
                                    unh_id = sample.get('UNH#', '')
                                    sample_label = ctk.CTkLabel(
                                        day_frame,
                                        text=f"UNH# {unh_id}",
                                        font=("Helvetica", 11),
                                        text_color="red"
                                    )
                                    sample_label.pack(pady=0)
                                    # Bind the label to directly edit this specific UNH#
                                    sample_label.bind("<Button-1>",
                                                      lambda event, uid=unh_id: self.edit_sample_by_unh(event, uid))
                            else:
                                # If more than 3, show count with first UNH ID
                                first_unh = day_samples[0].get('UNH#', '')
                                sample_label = ctk.CTkLabel(
                                    day_frame,
                                    text=f"{count} due: {first_unh}...",
                                    font=("Helvetica", 11),
                                    text_color="red"
                                )
                                sample_label.pack(pady=0)
                                sample_label.bind("<Button-1>",
                                                  lambda event, d=day: self.show_samples_due_on_day(d))

                    # Make the day clickable
                    day_frame.bind("<Button-1>",
                                   lambda event, d=day: self.show_samples_due_on_day(d))
                    day_num_label.bind("<Button-1>",
                                       lambda event, d=day: self.show_samples_due_on_day(d))

        # Configure grid to expand properly
        for i in range(7):  # 7 columns
            days_frame.columnconfigure(i, weight=1)

        for i in range(7):  # Up to 6 weeks plus header
            days_frame.rowconfigure(i, weight=1)

    def update_filter_checkbox_label(self, percent_filtered):
        """Update the checkbox label with the percentage of records that would be filtered."""
        if hasattr(self, 'filter_by_date_checkbox'):
            years_limit = getattr(self, 'years_limit', 1)  # Default to 1 year if not set

            if percent_filtered < 5:
                # Less than 5% would be filtered - let the user know
                self.filter_by_date_checkbox.configure(
                    text=f"Limit results to samples less than {years_limit} year old (only affects {percent_filtered:.1f}% of records)"
                )
            else:
                # Significant filtering - show the percentage
                self.filter_by_date_checkbox.configure(
                    text=f"Limit results to samples less than {years_limit} year old (filters out {percent_filtered:.1f}% of records)"
                )

    def update_calendar_view(self):
        """Update the calendar view when month or year changes."""
        self.initialize_calendar_view()
        # Clear the samples due treeview
        for item in self.due_samples_tree.get_children():
            self.due_samples_tree.delete(item)

    def get_samples_due_in_month(self, year, month):
        """Get all samples with due dates in the specified month and year."""
        try:
            conn = self._get_db_connection()
            if not conn:
                return []

            cursor = conn.cursor()

            # Create date range for the month
            start_date = f"{year}-{month:02d}-01"

            # Calculate the last day of the month
            if month == 12:
                next_month = 1
                next_year = year + 1
            else:
                next_month = month + 1
                next_year = year

            end_date = f"{next_year}-{next_month:02d}-01"

            # Log what we're doing for debugging
            print(f"Querying samples due between {start_date} and {end_date}")

            # Only query the specific fields we need rather than using SELECT *
            # This joins the tables only for samples with due dates in the specified month
            query = """
            SELECT a.[UNH#], s.Sample_Name, s.Collection_Date, a.Due_Date, s.Project
            FROM [WRRC sample analysis requested] AS a 
            LEFT JOIN [WRRC sample info] AS s 
            ON a.[UNH#] = s.[UNH#]
            WHERE a.Due_Date >= ? AND a.Due_Date < ?
            ORDER BY a.Due_Date
            """

            cursor.execute(query, (start_date, end_date))

            # Fetch all results
            samples = []
            for row in cursor.fetchall():
                sample = {
                    'UNH#': row[0] if row[0] else '',
                    'Sample_Name': row[1] if row[1] else '',
                    'Collection_Date': row[2] if row[2] else '',
                    'Due_Date': row[3] if row[3] else '',
                    'Project': row[4] if row[4] else ''
                }
                samples.append(sample)

            cursor.close()
            conn.close()

            print(f"Found {len(samples)} samples due in {month}/{year}")
            return samples

        except Exception as e:
            print(f"Error getting samples due in month: {str(e)}")
            traceback.print_exc()
            return []

    def edit_selected_due_sample(self, event=None):
        """Handle double-click on a sample in the due samples list."""
        # Get the treeview that was clicked
        tree = event.widget

        # Get the selected item
        selected_items = tree.selection()
        if not selected_items:
            print("No item selected")
            return

        # Get the selected item
        item_id = selected_items[0]
        item_values = tree.item(item_id, "values")

        # Find this sample in the database by UNH#
        unh_id = item_values[0]  # First column is UNH#
        print(f"Opening UNH# {unh_id} for editing")

        # Query the database for the full record
        conn = self._get_db_connection()
        if not conn:
            return

        cursor = conn.cursor()

        try:
            # Get sample info
            query = "SELECT * FROM [WRRC sample info] WHERE [UNH#] = ?"
            cursor.execute(query, (unh_id,))

            # Get column names
            columns = [column[0] for column in cursor.description]

            # Get the row
            row = cursor.fetchone()

            if row:
                # Create a dictionary from column names and values
                record_dict = {}
                for i, col in enumerate(columns):
                    record_dict[col] = row[i] if i < len(row) else ""

                # Store the selected record
                self.selected_record = record_dict

                # Load analysis data for this record
                self.load_analysis_data(unh_id)

                # Switch to the Edit tab
                self.tabview.set("Edit")

                # Populate the edit form
                self.populate_edit_form()
            else:
                messagebox.showwarning("Sample Not Found", f"No sample record found for UNH# {unh_id}")

        except Exception as e:
            print(f"Error loading sample for editing: {str(e)}")
            traceback.print_exc()
            messagebox.showerror("Error", f"Could not load sample for editing: {str(e)}")
        finally:
            cursor.close()
            conn.close()

    def edit_sample_by_unh(self, event, unh_id):
        """Open the edit tab for a specific UNH#."""
        print(f"Opening UNH# {unh_id} for editing from calendar")

        # Query the database for the full record
        conn = self._get_db_connection()
        if not conn:
            return

        cursor = conn.cursor()

        try:
            # Get sample info
            query = "SELECT * FROM [WRRC sample info] WHERE [UNH#] = ?"
            cursor.execute(query, (unh_id,))

            # Get column names
            columns = [column[0] for column in cursor.description]

            # Get the row
            row = cursor.fetchone()

            if row:
                # Create a dictionary from column names and values
                record_dict = {}
                for i, col in enumerate(columns):
                    record_dict[col] = row[i] if i < len(row) else ""

                # Store the selected record
                self.selected_record = record_dict

                # Load analysis data for this record
                self.load_analysis_data(unh_id)

                # Switch to the Edit tab
                self.tabview.set("Edit")

                # Populate the edit form
                self.populate_edit_form()
            else:
                messagebox.showwarning("Sample Not Found", f"No sample record found for UNH# {unh_id}")

        except Exception as e:
            print(f"Error loading sample for editing: {str(e)}")
            traceback.print_exc()
            messagebox.showerror("Error", f"Could not load sample for editing: {str(e)}")
        finally:
            cursor.close()
            conn.close()

    def show_samples_due_on_day(self, day):
        """Show samples due on the selected day in the month."""
        try:
            # Get current month and year
            month_name = self.calendar_month_var.get()
            month_num = list(calendar.month_name).index(month_name)
            year = int(self.calendar_year_var.get())

            # Create the date string
            selected_date = f"{year}-{month_num:02d}-{day:02d}"

            # Clear existing items in treeview
            for item in self.due_samples_tree.get_children():
                self.due_samples_tree.delete(item)

            # Get the connection
            conn = self._get_db_connection()
            if not conn:
                return

            cursor = conn.cursor()

            # Query for samples due on this specific day - using proper Access SQL syntax
            query = """
            SELECT a.[UNH#], s.Sample_Name, s.Collection_Date, a.Due_Date, s.Project
            FROM [WRRC sample analysis requested] AS a 
            INNER JOIN [WRRC sample info] AS s 
            ON a.[UNH#] = s.[UNH#]
            WHERE a.Due_Date = ?
            ORDER BY s.Project, s.Sample_Name
            """

            cursor.execute(query, (selected_date,))

            # Add samples to the treeview
            for row in cursor.fetchall():
                values = [
                    row[0] if row[0] else "",  # UNH#
                    row[1] if row[1] else "",  # Sample_Name
                    row[2] if row[2] else "",  # Collection_Date
                    row[3] if row[3] else "",  # Due_Date
                    row[4] if row[4] else ""  # Project
                ]
                self.due_samples_tree.insert("", "end", values=values)

            cursor.close()
            conn.close()

        except Exception as e:
            print(f"Error showing samples due on day: {str(e)}")
            traceback.print_exc()

    def create_edit_tab(self):
        """Create the edit tab contents."""
        edit_tab = self.tabview.tab("Edit")

        # Create a main frame for the edit tab
        main_frame = ctk.CTkFrame(edit_tab)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Create buttons for saving and canceling at the top
        button_frame = ctk.CTkFrame(main_frame)
        button_frame.pack(fill="x", padx=10, pady=10)

        save_button = ctk.CTkButton(
            button_frame,
            text="Save Changes",
            command=self.save_edited_record
        )
        save_button.pack(side="left", padx=10)

        cancel_button = ctk.CTkButton(
            button_frame,
            text="Cancel",
            command=lambda: self.tabview.set("Search")
        )
        cancel_button.pack(side="left", padx=10)

        # Status label
        self.edit_status_var = ctk.StringVar()
        self.edit_status_var.set("Select a record from the Search tab to edit")

        status_label = ctk.CTkLabel(
            button_frame,
            textvariable=self.edit_status_var,
            font=("Helvetica", 16),
            text_color="blue"
        )
        status_label.pack(side="left", padx=20)

        # Add "Saved" confirmation label with green text (initially hidden)
        self.saved_label = ctk.CTkLabel(
            button_frame,
            text="Saved.",
            font=("Helvetica", 16, "bold"),
            text_color="#00aa00"  # Bright green
        )

        # Create a frame for the sample info
        sample_info_frame = ctk.CTkFrame(main_frame)
        sample_info_frame.pack(fill="x", padx=10, pady=10)

        # Sample Info Title
        sample_info_title = ctk.CTkLabel(
            sample_info_frame,
            text="Sample Information",
            font=("Helvetica", 16, "bold")
        )
        sample_info_title.pack(pady=(5, 10))

        # Create a scrollable frame for the sample info fields
        sample_scroll_frame = ctk.CTkScrollableFrame(sample_info_frame, height=250)
        sample_scroll_frame.pack(fill="x", padx=10, pady=10)

        # Define common sample info fields
        self.sample_info_fields = [
            "UNH#", "Sample_Name", "Collection_Date", "Collection_Time",
            "Project", "Sub_Project", "Sub_ProjectA", "Sub_ProjectB",
            "Sample_Type", "Field_Notes", "pH", "Cond", "Spec_Cond",
            "DO_Conc", "DO%", "Temperature", "Salinity"
        ]

        # Create entry widgets for each field
        self.sample_info_entries = {}

        for i, field in enumerate(self.sample_info_fields):
            row = i // 2
            col = i % 2 * 2

            # Label
            label = ctk.CTkLabel(sample_scroll_frame, text=f"{field}:")
            label.grid(row=row, column=col, padx=(10, 5), pady=5, sticky="e")

            # Entry
            entry = ctk.CTkEntry(sample_scroll_frame, width=200)
            entry.grid(row=row, column=col + 1, padx=(0, 10), pady=5, sticky="w")

            self.sample_info_entries[field] = entry

        # Create a frame for the analysis info
        analysis_frame = ctk.CTkFrame(main_frame)
        analysis_frame.pack(fill="x", padx=10, pady=10)

        # Analysis Info Title
        analysis_title = ctk.CTkLabel(
            analysis_frame,
            text="Analysis Information",
            font=("Helvetica", 16, "bold")
        )
        analysis_title.pack(pady=(5, 10))

        # Create a scrollable frame for the analysis fields
        analysis_scroll_frame = ctk.CTkScrollableFrame(analysis_frame, height=250)
        analysis_scroll_frame.pack(fill="x", padx=10, pady=10)

        # Define common analysis fields
        self.analysis_fields = [
            "Containers", "Filtered", "Preservation", "Filter_Volume",
            "DOC", "TDN", "Anions", "Cations", "NO3AndNO2", "NO2", "NH4",
            "PO4OrSRP", "SiO2", "TN", "TP", "TDP", "TSS", "PCAndPN",
            "Chl_a", "EEMs", "Gases_GC", "Additional", "Due_Date"  # Added Due_Date
        ]

        # Mapping between analysis fields and related tables
        self.data_table_mapping = {
            "DOC": "NPOC",
            "TDN": "TDN",
            "Anions": "Anion",
            "Cations": "Cation",
            "NO3AndNO2": "NO3_Cd",
            "NH4": "NH4",
            "PO4OrSRP": "PO4",
            "SiO2": "SiO2",
            "TP": "TP"
        }

        # Create a dictionary to hold the data existence labels
        self.data_exists_labels = {}

        # Create a dedicated frame for the Due Date at the top of the analysis section
        # In the create_edit_tab method, modify the Due Date section:

        # Create a dedicated frame for the Due Date at the top of the analysis section
        # In create_edit_tab method:

        # First, create a dedicated frame just for the completion status at the very top
        # First, create a dedicated frame just for the completion status at the very top
        completion_frame = ctk.CTkFrame(analysis_scroll_frame, fg_color="#f0f0f0", corner_radius=8)
        completion_frame.grid(row=0, column=0, columnspan=6, padx=10, pady=(5, 15), sticky="ew")

        # Left side - Checkbox
        checkbox_frame = ctk.CTkFrame(completion_frame, fg_color="transparent")
        checkbox_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        # Completion status header
        completion_header = ctk.CTkLabel(
            checkbox_frame,
            text="Analysis Status:",
            font=("Helvetica", 14, "bold"),
            text_color="#333333"
        )
        completion_header.pack(side="left", padx=5, pady=10)

        # Create the checkbox with a clear visual style
        self.analysis_completed_var = ctk.BooleanVar(value=False)
        self.completed_checkbox = ctk.CTkCheckBox(
            checkbox_frame,
            text="Mark Analysis as Completed",
            variable=self.analysis_completed_var,
            command=self.toggle_due_date_state,
            font=("Helvetica", 12, "bold"),
            text_color="#227722",  # Green text
            fg_color="#2c974b",  # Green checkbox
            hover_color="#2da44e",  # Slightly lighter green on hover
            checkbox_width=24,  # Larger checkbox
            checkbox_height=24  # Larger checkbox
        )
        self.completed_checkbox.pack(side="left", padx=5, pady=10)

        # Right side - Due date
        due_date_frame = ctk.CTkFrame(completion_frame, fg_color="transparent")
        due_date_frame.pack(side="right", fill="both", padx=10, pady=10)

        # Due Date Label with emphasis
        due_date_label = ctk.CTkLabel(
            due_date_frame,
            text="Analysis Due Date:",
            font=("Helvetica", 12, "bold"),
            text_color="#c22a1f"  # Red color for emphasis
        )
        due_date_label.pack(side="left", padx=5)

        # Date picker using tkcalendar's DateEntry
        due_date_entry = DateEntry(
            due_date_frame,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='yyyy-mm-dd'
        )
        due_date_entry.pack(side="left", padx=5)
        due_date_entry.bind("<<DateEntrySelected>>", self.fix_calendar_popup)

        # Add the due date entry to our entries dictionary
        self.analysis_entries = {}
        self.analysis_entries["Due_Date"] = due_date_entry
        due_date_entry.bind("<<DateEntrySelected>>", self.fix_calendar_popup)

        # Start other fields at row 2 now
        row_offset = 2  # Start after completion status and due date rows

        # Create entry widgets for the remaining analysis fields
        row_offset = 1  # Start after the due date row

        for i, field in enumerate(self.analysis_fields):
            # Skip Due_Date as we've already added it
            if field == "Due_Date":
                continue

            row = (i // 3) + row_offset
            col = i % 3 * 2

            # Frame to hold the label and data exists indicator
            field_frame = ctk.CTkFrame(analysis_scroll_frame, fg_color="transparent")
            field_frame.grid(row=row, column=col, padx=(10, 5), pady=5, sticky="e")

            # Label
            label = ctk.CTkLabel(field_frame, text=f"{field}:")
            label.pack(side="left", padx=0)

            # Data exists label - will be shown/hidden during form population
            if field in self.data_table_mapping:
                data_exists_label = ctk.CTkLabel(
                    field_frame,
                    text="Data exists",
                    text_color="#2ecc71",  # Green color
                    font=("Helvetica", 16)
                )
                self.data_exists_labels[field] = data_exists_label
                # Initially hidden, will be shown when data exists
                # data_exists_label.pack(side="left", padx=(3, 0))

            # Entry
            entry = ctk.CTkEntry(analysis_scroll_frame, width=150)
            entry.grid(row=row, column=col + 1, padx=(0, 10), pady=5, sticky="w")

            self.analysis_entries[field] = entry

        # Create buttons for saving and canceling at the bottom
        bottom_button_frame = ctk.CTkFrame(main_frame)
        bottom_button_frame.pack(fill="x", padx=10, pady=10)

        save_button = ctk.CTkButton(
            bottom_button_frame,
            text="Save Changes",
            command=self.save_edited_record
        )
        save_button.pack(side="left", padx=10)

        cancel_button = ctk.CTkButton(
            bottom_button_frame,
            text="Cancel",
            command=lambda: self.tabview.set("Search")
        )
        cancel_button.pack(side="left", padx=10)

    def toggle_due_date_state(self):
        """Toggle the due date entry state based on the completed checkbox."""
        try:
            due_date_entry = self.analysis_entries.get("Due_Date")
            if not due_date_entry:
                return

            if self.analysis_completed_var.get():
                # Analysis is completed, disable the date entry
                due_date_entry.configure(state="disabled")
                print("Due date disabled - analysis completed")
            else:
                # Analysis is not completed, enable the date entry
                due_date_entry.configure(state="normal")
                print("Due date enabled - analysis not completed")
        except Exception as e:
            print(f"Error toggling due date state: {str(e)}")
    # Add this method to your class:
    def fix_calendar_popup(self, event=None):
        """Ensure the calendar popup has enough space for navigation buttons."""
        try:
            # This runs after the calendar is displayed
            # Get the toplevel window of the calendar popup
            for toplevel in self.winfo_toplevel().winfo_children():
                if isinstance(toplevel, tk.Toplevel) and hasattr(toplevel, 'calendar'):
                    # Found the calendar toplevel
                    calendar = toplevel.calendar

                    # Make sure the month navigation buttons are visible
                    prev_button = calendar._prev_month_button
                    next_button = calendar._next_month_button

                    if prev_button and next_button:
                        # Increase button size
                        prev_button.configure(width=6, font=('Helvetica', 16, 'bold'))
                        next_button.configure(width=6, font=('Helvetica', 16, 'bold'))

                        # Ensure the calendar has enough space
                        toplevel.update_idletasks()
                        current_width = toplevel.winfo_width()
                        if current_width < 300:  # If it's too narrow
                            toplevel.geometry(f"300x{toplevel.winfo_height()}")

                    break
        except Exception as e:
            print(f"Error fixing calendar popup: {str(e)}")

    def check_related_data(self, unh_id):
        """
        Check if related data exists for a sample in measurement tables.
        Returns a dictionary with table names as keys and boolean values indicating data existence.
        """
        # Mapping of analysis fields to actual table names
        table_mapping = {
            "NPOC": "WRRC NPOC Data",
            "NO3_Cd": "WRRC NO3_Cd Data",
            "Cation": "WRRC Cation Data",
            "Anion": "WRRC Anion Data",
            "PO4": "WRRC PO4 Data",
            "SiO2": "WRRC SiO2 data",  # Note lowercase 'data'
            "TDN": "WRRC TDN Data",
            "TP": "WRRC TP Data",
            "NH4": "WRRC NH4 Data",
            "DIC": "WRRC DIC Data"
        }

        # Initialize all to false
        related_data = {key: False for key in table_mapping.keys()}

        try:
            conn = self._get_db_connection()
            if not conn:
                return related_data

            cursor = conn.cursor()

            # Check each related table
            for key, full_table_name in table_mapping.items():
                try:
                    # Use a parameterized query to check for data
                    query = f"SELECT COUNT(*) FROM [{full_table_name}] WHERE [UNH#] = ?"
                    cursor.execute(query, (unh_id,))
                    count = cursor.fetchone()[0]

                    # If count > 0, data exists
                    related_data[key] = count > 0
                    print(f"Table {full_table_name}: {count} records found for UNH# {unh_id}")

                except Exception as table_error:
                    print(f"Error checking table {full_table_name}: {str(table_error)}")
                    # Keep default False value

            cursor.close()
            conn.close()

        except Exception as e:
            print(f"Error checking related data: {str(e)}")
            traceback.print_exc()

        return related_data

    # Add the load_analysis_data method
    def load_analysis_data(self, unh_id):
        """Load analysis data for the given UNH ID."""
        try:
            conn = self._get_db_connection()
            if not conn:
                return

            cursor = conn.cursor()

            # Query the analysis data
            query = f"SELECT * FROM [WRRC sample analysis requested] WHERE [UNH#] = ?"
            cursor.execute(query, (unh_id,))

            # Get column names
            columns = [column[0] for column in cursor.description]

            # Get the first row
            row = cursor.fetchone()

            if row:
                # Create a dictionary from column names and values
                analysis_dict = {}
                for i, col in enumerate(columns):
                    analysis_dict[col] = row[i] if i < len(row) else ""

                self.analysis_data = analysis_dict
            else:
                self.analysis_data = None

            cursor.close()
            conn.close()

        except Exception as e:
            print(f"Error loading analysis data: {str(e)}")
            self.analysis_data = None

    # Add the populate_edit_form method
    def populate_edit_form(self):
        """Populate the edit form with the selected record data."""
        if not self.selected_record:
            self.edit_status_var.set("No record selected for editing")
            return

        # Clear all entries first
        for field, entry in self.sample_info_entries.items():
            entry.delete(0, "end")

        for field, entry in self.analysis_entries.items():
            # Special handling for DateEntry widget for Due_Date
            if field == "Due_Date" and hasattr(entry, 'set_date'):
                # Reset to today's date as default
                today = datetime.date.today()
                entry.set_date(today)
            else:
                # Standard text entry
                try:
                    entry.delete(0, "end")
                except Exception as e:
                    print(f"Error clearing field {field}: {str(e)}")

        # Populate sample info entries
        for field, entry in self.sample_info_entries.items():
            if field in self.selected_record:
                value = self.selected_record[field]
                entry.insert(0, str(value) if value is not None else "")

        # Get the UNH ID for checking related data
        unh_id = self.selected_record.get("UNH#", "")

        # Check for related data if we have a UNH ID
        if unh_id:
            related_data = self.check_related_data(unh_id)

            # Show/hide data exists labels based on results
            for field, label in self.data_exists_labels.items():
                table_name = self.data_table_mapping.get(field)
                if table_name and related_data.get(table_name, False):
                    # Data exists - show the label
                    label.pack(side="left", padx=(3, 0))
                else:
                    # No data - hide the label
                    label.pack_forget()

        # Populate analysis entries if we have analysis data
        if self.analysis_data:
            for field, entry in self.analysis_entries.items():
                if field in self.analysis_data:
                    value = self.analysis_data[field]

                    # Special handling for Due_Date which uses DateEntry
                    if field == "Due_Date" and hasattr(entry, 'set_date'):
                        try:
                            if value and str(value).strip():
                                # There is a due date, so analysis is not completed
                                if hasattr(self, 'analysis_completed_var'):
                                    self.analysis_completed_var.set(False)
                                if hasattr(self, 'completed_checkbox') and hasattr(self.completed_checkbox, 'deselect'):
                                    self.completed_checkbox.deselect()
                                entry.configure(state="normal")

                                try:
                                    # Try to parse the date
                                    if isinstance(value, datetime.date):
                                        due_date = value
                                    else:
                                        # Try different date formats
                                        for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y']:
                                            try:
                                                due_date = datetime.datetime.strptime(str(value), fmt).date()
                                                break
                                            except:
                                                continue
                                        else:
                                            # If no format worked, default to today
                                            due_date = datetime.date.today()

                                    # Set the date in the widget
                                    entry.set_date(due_date)
                                except Exception as e:
                                    print(f"Error setting date for Due_Date: {str(e)}")
                                    # Default to today if there's an error
                                    entry.set_date(datetime.date.today())
                            else:
                                # No due date - analysis is completed
                                if hasattr(self, 'analysis_completed_var'):
                                    self.analysis_completed_var.set(True)
                                if hasattr(self, 'completed_checkbox') and hasattr(self.completed_checkbox, 'select'):
                                    self.completed_checkbox.select()
                                entry.configure(state="disabled")
                        except Exception as e:
                            print(f"Error handling due date: {str(e)}")
                            # Use safe defaults
                            if hasattr(self, 'analysis_completed_var'):
                                self.analysis_completed_var.set(False)
                            if hasattr(entry, 'set_date'):
                                entry.set_date(datetime.date.today())
                                entry.configure(state="normal")
                    else:
                        # Regular entry widget
                        entry.insert(0, str(value) if value is not None else "")

        # Handle the case where a due date isn't present in the analysis data
        if self.analysis_data is None or "Due_Date" not in self.analysis_data or not self.analysis_data["Due_Date"]:
            try:
                # No analysis data or no due date - mark as completed
                if hasattr(self, 'analysis_completed_var'):
                    self.analysis_completed_var.set(True)
                if hasattr(self, 'completed_checkbox') and hasattr(self.completed_checkbox, 'select'):
                    self.completed_checkbox.select()

                # Disable the date entry widget
                due_date_entry = self.analysis_entries.get("Due_Date")
                if due_date_entry and hasattr(due_date_entry, 'configure'):
                    due_date_entry.configure(state="disabled")
            except Exception as e:
                print(f"Error setting defaults for missing due date: {str(e)}")

        # Update status
        unh_id = self.selected_record.get("UNH#", "")
        sample_name = self.selected_record.get("Sample_Name", "")
        self.edit_status_var.set(f"Editing record: UNH# {unh_id}, Sample Name: {sample_name}")
    # Add the save_edited_record method
    def save_edited_record(self):
        """Save the edited record back to the database."""
        if not self.selected_record:
            messagebox.showwarning("No Record", "No record is selected for editing.")
            return

        try:
            # Get connection to database
            conn = self._get_db_connection()
            if not conn:
                self.edit_status_var.set("Error: Could not connect to the database")
                return

            cursor = conn.cursor()

            # Begin transaction
            if conn.autocommit:
                conn.autocommit = False

            # Update sample info table
            success_sample = self._update_sample_info_record(cursor)

            # Update analysis table
            success_analysis = self._update_analysis_record(cursor)

            if success_sample or success_analysis:
                # Commit transaction
                conn.commit()
                self.edit_status_var.set("Record updated successfully")

                # Show the "Saved" message
                self.saved_label.pack(side="right", padx=20)

                # Schedule the label to disappear after 3 seconds
                self.after(3000, lambda: self.saved_label.pack_forget())

                # Refresh the data
                self.data = self._load_data_from_database()
                self.populate_treeview(self.data)

                # Switch back to search tab after a brief delay to show the "Saved" message
                # self.after(1500, lambda: self.tabview.set("Search"))
            else:
                conn.rollback()
                self.edit_status_var.set("No changes were made")

            cursor.close()
            conn.close()

        except Exception as e:
            print(f"Error saving edited record: {str(e)}")
            print(traceback.format_exc())
            self.edit_status_var.set(f"Error: {str(e)}")
            messagebox.showerror("Save Error", f"Error saving changes: {str(e)}")

            # Try to rollback if there's a connection
            try:
                if conn and not conn.autocommit:
                    conn.rollback()
                    cursor.close()
                    conn.close()
            except:
                pass

    def _update_sample_info_record(self, cursor):
        """Update a record in the WRRC sample info table."""
        try:
            # Get the primary key value
            unh_id = self.selected_record.get("UNH#", "")
            if not unh_id:
                return False

            # Build SET clause and parameters for the UPDATE statement
            set_clauses = []
            params = []

            for field, entry in self.sample_info_entries.items():
                # Skip UNH# as it's our key
                if field == "UNH#":
                    continue

                # Get the current value and the new value
                current_value = self.selected_record.get(field, "")
                new_value = entry.get().strip()

                # If there's a difference, add to the update
                if str(current_value) != new_value:
                    # Handle special fields
                    if field == "DO%":
                        set_clauses.append("[DO%] = ?")
                    else:
                        set_clauses.append(f"[{field}] = ?")

                    # Handle empty strings as NULL for certain fields
                    if not new_value and field in ['Collection_Date', 'Collection_Time', 'pH', 'Cond', 'Spec_Cond',
                                                   'DO_Conc', 'DO%', 'Temperature', 'Salinity']:
                        params.append(None)
                    else:
                        params.append(new_value)

            # If no changes, return early
            if not set_clauses:
                print("No changes to update in sample info")
                return False

            # Build the query
            query = f"UPDATE [WRRC sample info] SET {', '.join(set_clauses)} WHERE [UNH#] = ?"
            params.append(unh_id)

            print(f"Sample update query: {query}")
            print(f"Parameters: {params}")

            # Execute the query
            cursor.execute(query, params)

            return True

        except Exception as e:
            print(f"Error updating sample info: {str(e)}")
            raise

    def _update_analysis_record(self, cursor):
        """Update a record in the WRRC sample analysis requested table."""
        try:
            # Get the primary key value
            unh_id = self.selected_record.get("UNH#", "")
            if not unh_id:
                return False

            # Check if we have analysis data
            if not self.analysis_data:
                # Check if we have any new values to insert
                has_new_values = False
                for field, entry in self.analysis_entries.items():
                    if field == "Due_Date":
                        # If analysis is not completed, check if date entry has value
                        if not self.analysis_completed_var.get() and hasattr(entry, 'get_date'):
                            has_new_values = True
                            break
                    elif entry.get().strip():
                        has_new_values = True
                        break

                if has_new_values:
                    # Need to INSERT a new record
                    return self._insert_new_analysis_record(cursor, unh_id)
                else:
                    return False

            # We have existing analysis data, so update it
            set_clauses = []
            params = []

            for field, entry in self.analysis_entries.items():
                # Get the current value from the analysis data
                current_value = self.analysis_data.get(field, "")

                # Special handling for Due_Date field
                if field == "Due_Date":
                    if self.analysis_completed_var.get():
                        # Analysis is completed, set Due_Date to NULL
                        # Only add to update if current value is not already NULL
                        if current_value is not None and current_value != "":
                            print(f"Setting Due_Date to NULL for UNH# {unh_id} (analysis completed)")
                            set_clauses.append(f"[{field}] = ?")
                            params.append(None)  # This will set it to NULL in the database
                    elif hasattr(entry, 'get_date'):
                        # Analysis not completed, get date from DateEntry
                        try:
                            new_value = entry.get_date().strftime('%Y-%m-%d')
                            # If different from current, add to update
                            if str(current_value) != new_value:
                                print(f"Updating Due_Date to {new_value} for UNH# {unh_id}")
                                set_clauses.append(f"[{field}] = ?")
                                params.append(new_value)
                        except Exception as e:
                            print(f"Error getting date from DateEntry: {str(e)}")
                else:
                    # Regular fields
                    new_value = entry.get().strip()

                    # If there's a difference, add to the update
                    if str(current_value) != new_value:
                        set_clauses.append(f"[{field}] = ?")

                        # Handle empty strings as NULL for appropriate fields
                        if not new_value:
                            params.append(None)
                        else:
                            params.append(new_value)

            # If no changes, return early
            if not set_clauses:
                print("No changes to update in analysis data")
                return False

            # Build the query
            query = f"UPDATE [WRRC sample analysis requested] SET {', '.join(set_clauses)} WHERE [UNH#] = ?"
            params.append(unh_id)

            print(f"Analysis update query: {query}")
            print(f"Parameters: {params}")

            # Execute the query
            cursor.execute(query, params)
            print(f"Successfully updated analysis data for UNH# {unh_id}")

            return True

        except Exception as e:
            print(f"Error updating analysis info: {str(e)}")
            raise

    def _insert_new_analysis_record(self, cursor, unh_id):
        """Insert a new record in the WRRC sample analysis requested table."""
        try:
            # Build columns and values for the INSERT statement
            columns = ["[UNH#]"]
            values = [unh_id]

            for field, entry in self.analysis_entries.items():
                # Handle Due_Date field which uses DateEntry
                if field == "Due_Date":
                    if not self.analysis_completed_var.get() and hasattr(entry, 'get_date'):
                        due_date = entry.get_date().strftime('%Y-%m-%d')
                        if due_date:
                            columns.append(f"[{field}]")
                            values.append(due_date)
                    # If analysis is completed, don't add Due_Date (it will be NULL by default)
                else:
                    # Regular entry fields
                    value = entry.get().strip()
                    if value:
                        columns.append(f"[{field}]")
                        values.append(value)

            # If only UNH#, no need to insert
            if len(columns) <= 1:
                return False

            # Build the query
            query = f"INSERT INTO [WRRC sample analysis requested] ({', '.join(columns)}) VALUES ({', '.join(['?'] * len(values))})"

            print(f"Analysis insert query: {query}")
            print(f"Parameters: {values}")

            # Execute the query
            cursor.execute(query, values)

            return True

        except Exception as e:
            print(f"Error inserting analysis info: {str(e)}")
            raise

    def _insert_new_analysis_record(self, cursor, unh_id):
        """Insert a new record in the WRRC sample analysis requested table."""
        try:
            # Build columns and values for the INSERT statement
            columns = ["[UNH#]"]
            values = [unh_id]

            for field, entry in self.analysis_entries.items():
                value = entry.get().strip()
                if value:
                    columns.append(f"[{field}]")
                    values.append(value)

            # If only UNH#, no need to insert
            if len(columns) <= 1:
                return False

            # Build the query
            query = f"INSERT INTO [WRRC sample analysis requested] ({', '.join(columns)}) VALUES ({', '.join(['?'] * len(values))})"

            print(f"Analysis insert query: {query}")
            print(f"Parameters: {values}")

            # Execute the query
            cursor.execute(query, values)

            return True

        except Exception as e:
            print(f"Error inserting analysis info: {str(e)}")
            raise

    # First, add the edit button to the search tab
    def create_search_tab(self):
        """Create the search tab contents."""
        search_tab = self.tabview.tab("Search")

        # Create search frame
        search_frame = ctk.CTkFrame(search_tab)
        search_frame.pack(pady=10, fill="x")

        # Sample Search Widgets
        sample_label = ctk.CTkLabel(search_frame, text="Search by Sample (UNH# or Sample_Name):")
        sample_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.sample_search_entry = ctk.CTkEntry(search_frame, width=300)
        self.sample_search_entry.grid(row=0, column=1, padx=10, pady=10)

        sample_search_button = ctk.CTkButton(
            search_frame,
            text="Search Sample",
            command=self.search_by_sample
        )
        sample_search_button.grid(row=0, column=2, padx=10, pady=10)

        # Project Search Widgets
        project_label = ctk.CTkLabel(search_frame, text="Search by Project (Project, Sub_Project, etc.):")
        project_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.project_search_entry = ctk.CTkEntry(search_frame, width=300)
        self.project_search_entry.grid(row=1, column=1, padx=10, pady=10)

        project_search_button = ctk.CTkButton(
            search_frame,
            text="Search Project",
            command=self.search_by_project
        )
        project_search_button.grid(row=1, column=2, padx=10, pady=10)

        # Date filter checkbox
        filter_frame = ctk.CTkFrame(search_frame)
        filter_frame.grid(row=2, column=0, columnspan=3, pady=5, sticky="w")

        self.filter_by_date_var = ctk.BooleanVar(value=True)  # Default to checked
        self.filter_by_date_checkbox = ctk.CTkCheckBox(
            filter_frame,
            text=f"Limit results to samples less than {self.years_limit} year old (improves search speed)",
            variable=self.filter_by_date_var,
            command=self.refresh_search  # Refresh search results when toggled
        )
        self.filter_by_date_checkbox.pack(side="left", padx=10, pady=5)

        # Clear Search and Show All Buttons
        button_frame = ctk.CTkFrame(search_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=10)

        clear_search_button = ctk.CTkButton(
            button_frame,
            text="Clear Search",
            command=self.clear_search
        )
        clear_search_button.pack(side="left", padx=10)

        show_all_button = ctk.CTkButton(
            button_frame,
            text="Show All",
            command=self.show_all
        )
        show_all_button.pack(side="left", padx=10)

        # Add Edit button
        edit_button = ctk.CTkButton(
            button_frame,
            text="Edit Selected",
            command=self.edit_selected_record
        )
        edit_button.pack(side="left", padx=10)

        # Treeview for results
        treeview_frame = ctk.CTkFrame(search_tab)
        treeview_frame.pack(fill="both", expand=True, pady=10)

        # CustomTkinter doesn't have its own treeview, so we'll use the ttk one
        self.tree = self.create_styled_treeview(treeview_frame)

        # Add scrollbars
        y_scrollbar = ctk.CTkScrollbar(treeview_frame, command=self.tree.yview)
        y_scrollbar.pack(side="right", fill="y")

        x_scrollbar = ctk.CTkScrollbar(treeview_frame, orientation="horizontal", command=self.tree.xview)
        x_scrollbar.pack(side="bottom", fill="x")

        self.tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)

        # Populate the tree with data
        self.populate_treeview(self.data)

        # Bind double-click event to edit function
        self.tree.bind("<Double-1>", lambda event: self.edit_selected_record())

    def create_styled_treeview(self, parent):
        """Create a ttk.Treeview with styling to match CustomTkinter."""
        import tkinter as tk
        from tkinter import ttk

        # Create a style for the treeview
        style = ttk.Style()

        # Configure the treeview style
        current_theme = ctk.get_appearance_mode().lower()

        if current_theme == "dark":
            # Dark theme settings
            style.configure(
                "Custom.Treeview",
                background="#2a2d2e",
                foreground="white",
                fieldbackground="#2a2d2e",
                rowheight=40,
                font=('Helvetica', 12)
            )
            style.configure(
                "Custom.Treeview.Heading",
                background="#1f6aa5",
                foreground="white",
                font=('Helvetica', 13, 'bold')
            )
        else:
            # Light theme settings
            style.configure(
                "Custom.Treeview",
                background="white",
                foreground="black",
                fieldbackground="white",
                rowheight=40,
                font=('Helvetica', 12)
            )
            style.configure(
                "Custom.Treeview.Heading",
                background="#1f6aa5",
                foreground="white",
                font=('Helvetica', 13, 'bold')
            )

        # Create the treeview with the columns from the data
        columns = self.data.columns.tolist() if not self.data.empty else []
        tree = ttk.Treeview(
            parent,
            columns=columns,
            show='headings',
            style="Custom.Treeview"
        )

        # Configure the columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150, minwidth=50)

        return tree

    def populate_treeview(self, df):
        """Populate the treeview with data from the DataFrame."""
        # Clear the current content of the treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        if df.empty:
            print("No data to populate treeview")
            return

        # Insert rows into the treeview
        for _, row in df.iterrows():
            # Convert any non-string values to strings
            values = [str(val) if not isinstance(val, str) and val is not None else "" if val is None else val for val
                      in row]
            self.tree.insert("", "end", values=values)

        print(f"Treeview populated with {len(df)} rows.")

    def show_all(self):
        """Display all records, but respect the date filter if enabled."""
        all_data = self.data.copy()

        # Apply date filter if checkbox is checked
        filtered_data = self.apply_date_filter(all_data)

        records_count = len(filtered_data)
        total_count = len(all_data)

        if self.filter_by_date_var.get():
            print(f"Displaying {records_count} records out of {total_count} total (filtered by date).")
        else:
            print(f"Displaying all {records_count} records from the Access database.")

        self.populate_treeview(filtered_data)

    def create_import_tab(self):
        """Create the import tab contents with support for both submission and log book formats."""
        import_tab = self.tabview.tab("Import")

        # Create frames for the import tab
        instruction_frame = ctk.CTkFrame(import_tab)
        instruction_frame.pack(fill="x", padx=10, pady=10)

        # Instructions
        instruction_label = ctk.CTkLabel(
            instruction_frame,
            text="Import Sample Data",
            font=("Helvetica", 16, "bold")
        )
        instruction_label.pack(pady=5)

        description_label = ctk.CTkLabel(
            instruction_frame,
            text="Select an Excel file format to import. You can import from either a sample submission form or a log book file.",
            wraplength=800
        )
        description_label.pack(pady=5)

        # ================ CUSTOM TABS IMPLEMENTATION ================
        # Create a frame to hold our custom tab system
        tab_system_frame = ctk.CTkFrame(import_tab)
        tab_system_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Create a frame for tab buttons
        tab_buttons_frame = ctk.CTkFrame(tab_system_frame, fg_color="transparent")
        tab_buttons_frame.pack(fill="x", padx=0, pady=0)

        # Create a frame for tab content
        self.tab_content_frame = ctk.CTkFrame(tab_system_frame)
        self.tab_content_frame.pack(fill="both", expand=True, padx=0, pady=(0, 10))

        # Create the content frames for each tab
        self.submission_content_frame = ctk.CTkFrame(self.tab_content_frame)
        self.logbook_content_frame = ctk.CTkFrame(self.tab_content_frame)

        # Variables to keep track of the active tab
        self.active_tab = ctk.StringVar(value="submission")

        # Function to switch between tabs
        def switch_tab(tab_name):
            # Hide all tab content frames
            self.submission_content_frame.pack_forget()
            self.logbook_content_frame.pack_forget()

            # Update button colors based on active tab
            if tab_name == "submission":
                submission_tab_button.configure(fg_color="#27ae60", hover_color="#2ecc71")  # Active green
                logbook_tab_button.configure(fg_color="#1e8449", hover_color="#27ae60")  # Inactive green
                self.submission_content_frame.pack(fill="both", expand=True, padx=0, pady=0)
                self.active_tab.set("submission")
            else:  # logbook
                submission_tab_button.configure(fg_color="#1e8449", hover_color="#27ae60")  # Inactive green
                logbook_tab_button.configure(fg_color="#27ae60", hover_color="#2ecc71")  # Active green
                self.logbook_content_frame.pack(fill="both", expand=True, padx=0, pady=0)
                self.active_tab.set("logbook")

        # Create the tab buttons
        submission_tab_button = ctk.CTkButton(
            tab_buttons_frame,
            text="Sample Submission Format",
            font=("Helvetica", 16),
            fg_color="#27ae60",  # Active green
            hover_color="#2ecc71",
            text_color="white",
            corner_radius=6,
            command=lambda: switch_tab("submission")
        )
        submission_tab_button.pack(side="left", padx=(0, 5), pady=0, fill="x", expand=True)

        logbook_tab_button = ctk.CTkButton(
            tab_buttons_frame,
            text="Log Book Format",
            font=("Helvetica", 16),
            fg_color="#1e8449",  # Inactive green
            hover_color="#27ae60",
            text_color="white",
            corner_radius=6,
            command=lambda: switch_tab("logbook")
        )
        logbook_tab_button.pack(side="left", padx=0, pady=0, fill="x", expand=True)

        # Initially show the submission tab
        self.submission_content_frame.pack(fill="both", expand=True, padx=0, pady=0)

        # ================ SUBMISSION TAB CONTENT ================
        # Add project input field for Sample Submission format only
        submission_project_frame = ctk.CTkFrame(self.submission_content_frame)
        submission_project_frame.pack(fill="x", padx=10, pady=10)

        project_label = ctk.CTkLabel(submission_project_frame, text="Project Name:")
        project_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.project_entry = ctk.CTkEntry(submission_project_frame, width=300)
        self.project_entry.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        submission_file_frame = ctk.CTkFrame(self.submission_content_frame)
        submission_file_frame.pack(fill="x", padx=10, pady=10)

        # File selection for submission format
        sub_file_label = ctk.CTkLabel(submission_file_frame, text="Select Submission Excel File:")
        sub_file_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.submission_file_path_var = ctk.StringVar()
        sub_file_entry = ctk.CTkEntry(submission_file_frame, textvariable=self.submission_file_path_var, width=500)
        sub_file_entry.grid(row=0, column=1, padx=10, pady=10)

        sub_browse_button = ctk.CTkButton(
            submission_file_frame,
            text="Browse",
            command=lambda: self.browse_excel_file("submission")
        )
        sub_browse_button.grid(row=0, column=2, padx=10, pady=10)

        # Preview and import buttons for submission format
        sub_button_frame = ctk.CTkFrame(submission_file_frame)
        sub_button_frame.grid(row=1, column=0, columnspan=3, pady=10)

        sub_preview_button = ctk.CTkButton(
            sub_button_frame,
            text="Preview Data",
            command=self.preview_excel_data
        )
        sub_preview_button.pack(side="left", padx=10)

        sub_import_button = ctk.CTkButton(
            sub_button_frame,
            text="Import Data",
            command=self.import_excel_data
        )
        sub_import_button.pack(side="left", padx=10)

        # ================ LOG BOOK TAB CONTENT ================
        logbook_file_frame = ctk.CTkFrame(self.logbook_content_frame)
        logbook_file_frame.pack(fill="x", padx=10, pady=10)

        # File selection for log book format
        log_file_label = ctk.CTkLabel(logbook_file_frame, text="Select Log Book Excel File:")
        log_file_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.logbook_file_path_var = ctk.StringVar()
        log_file_entry = ctk.CTkEntry(logbook_file_frame, textvariable=self.logbook_file_path_var, width=500)
        log_file_entry.grid(row=0, column=1, padx=10, pady=10)

        log_browse_button = ctk.CTkButton(
            logbook_file_frame,
            text="Browse",
            command=lambda: self.browse_excel_file("logbook")
        )
        log_browse_button.grid(row=0, column=2, padx=10, pady=10)

        # Preview and import buttons for log book format
        log_button_frame = ctk.CTkFrame(logbook_file_frame)
        log_button_frame.grid(row=1, column=0, columnspan=3, pady=10)

        log_preview_button = ctk.CTkButton(
            log_button_frame,
            text="Preview Data",
            command=self.preview_logbook_data
        )
        log_preview_button.pack(side="left", padx=10)

        log_import_button = ctk.CTkButton(
            log_button_frame,
            text="Import Data",
            command=self.import_logbook_data
        )
        log_import_button.pack(side="left", padx=10)

        # ================ PREVIEW AREA ================
        # Common preview area using notebook with tabs
        preview_frame = ctk.CTkFrame(import_tab)
        preview_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Create a basic notebook for the preview (doesn't need fancy styling)
        preview_notebook = ttk.Notebook(preview_frame)
        preview_notebook.pack(fill="both", expand=True, padx=5, pady=5)

        # Project Info tab
        self.project_preview_frame = ttk.Frame(preview_notebook)
        preview_notebook.add(self.project_preview_frame, text="Project Information")

        # Sample Info tab
        self.sample_preview_frame = ttk.Frame(preview_notebook)
        preview_notebook.add(self.sample_preview_frame, text="Sample Information")

        # Create treeviews for previews
        self.project_tree = ttk.Treeview(self.project_preview_frame)
        self.project_tree.pack(fill="both", expand=True)

        # Add scrollbars for project tree
        project_y_scrollbar = ttk.Scrollbar(self.project_preview_frame, orient="vertical",
                                            command=self.project_tree.yview)
        project_y_scrollbar.pack(side="right", fill="y")

        project_x_scrollbar = ttk.Scrollbar(self.project_preview_frame, orient="horizontal",
                                            command=self.project_tree.xview)
        project_x_scrollbar.pack(side="bottom", fill="x")

        self.project_tree.configure(yscrollcommand=project_y_scrollbar.set, xscrollcommand=project_x_scrollbar.set)

        # Sample tree
        self.sample_tree = ttk.Treeview(self.sample_preview_frame)
        self.sample_tree.pack(fill="both", expand=True)

        # Add scrollbars for sample tree
        sample_y_scrollbar = ttk.Scrollbar(self.sample_preview_frame, orient="vertical", command=self.sample_tree.yview)
        sample_y_scrollbar.pack(side="right", fill="y")

        sample_x_scrollbar = ttk.Scrollbar(self.sample_preview_frame, orient="horizontal",
                                           command=self.sample_tree.xview)
        sample_x_scrollbar.pack(side="bottom", fill="x")

        self.sample_tree.configure(yscrollcommand=sample_y_scrollbar.set, xscrollcommand=sample_x_scrollbar.set)

        # Status label for import
        self.import_status_var = ctk.StringVar()
        self.import_status_var.set("No file selected")

        status_label = ctk.CTkLabel(
            preview_frame,
            textvariable=self.import_status_var,
            font=("Helvetica", 12),
            text_color="blue"
        )
        status_label.pack(pady=5)

    def browse_excel_file(self, file_type="submission"):
        """Open a file dialog to select an Excel file for the specified import type."""
        if file_type == "submission":
            title = "Select Sample Submission Excel File"
        else:  # logbook
            title = "Select Log Book Excel File"

        file_path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel Files", "*.xls *.xlsx")]
        )

        if file_path:
            if file_type == "submission":
                self.submission_file_path_var.set(file_path)
            else:  # logbook
                self.logbook_file_path_var.set(file_path)

            self.import_status_var.set(f"{file_type.capitalize()} file selected: {os.path.basename(file_path)}")
            print(f"Selected {file_type} file: {file_path}")

    def preview_logbook_data(self):
        """Preview the data from the selected Log Book Excel file."""
        file_path = self.logbook_file_path_var.get()
        if not file_path:
            messagebox.showwarning("No File Selected", "Please select a Log Book Excel file first.")
            return

        try:
            # Load the Excel file
            self.import_status_var.set("Loading Log Book file for preview...")

            # Try to read the Excel file
            log_data = self.read_logbook_excel(file_path)

            if log_data is None or log_data.empty:
                self.import_status_var.set("Error: Could not read the Log Book Excel file or it's empty.")
                return

            # Extract sample data for validation
            samples = self.extract_logbook_data(log_data)

            # Update the status
            self.import_status_var.set(f"Preview ready. Found {len(samples)} samples in Log Book format.")

            # Populate the preview treeviews
            self.populate_logbook_preview(log_data, samples)

        except Exception as e:
            error_message = f"Error previewing Log Book file: {str(e)}"
            print(error_message)
            print(traceback.format_exc())
            messagebox.showerror("Preview Error", error_message)
            self.import_status_var.set("Error previewing file. See console for details.")

    def read_logbook_excel(self, file_path):
        """Read the Log Book Excel file and return a DataFrame."""
        try:
            # Read the Excel file
            log_data = pd.read_excel(file_path)

            # Clean up the dataframe
            log_data = log_data.fillna("")

            # Print some info about the data
            print(f"Log Book DataFrame has {len(log_data)} rows and {len(log_data.columns)} columns")
            print(f"Column names: {log_data.columns.tolist()}")

            return log_data

        except Exception as e:
            error_message = f"Error reading Log Book Excel file: {str(e)}"
            print(error_message)
            print(traceback.format_exc())
            messagebox.showerror("Excel Import Error", error_message)
            return None

    def extract_logbook_data(self, log_data):
        """
        Extract sample information from the Log Book DataFrame.
        Returns a list of dictionaries, each containing a sample's information.
        Skip rows where only UNH# is populated.
        """
        samples = []

        # Check if DataFrame is empty
        if log_data.empty:
            return samples

        # Print column names for debugging
        print("Actual column names in Log Book DataFrame:")
        for col in log_data.columns:
            print(f"  {col}")

        # Define fields we're interested in
        field_mappings = {
            'UNH#': 'unh_id',
            'Sample_Name': 'sample_name',
            'Collection_Date': 'collection_date',
            'Collection_Time': 'collection_time',
            'Project': 'project',
            'Sub_Project': 'sub_project',
            'Sub_ProjectA': 'sub_projecta',
            'Sub_ProjectB': 'sub_projectb',
            'BatchID': 'batch_id',
            'Frozen_Received': 'frozen_received',
            'Refrigerated_Received': 'refrigerated_received',
            'Field_Notes': 'field_notes',
            'Lab_Notes': 'lab_notes',
            'Sample_Type': 'sample_type',
            'Logger': 'logger',
            'pH': 'ph',
            'Cond': 'cond',
            'Spec_Cond': 'spec_cond',
            'DO_Conc': 'do_conc',
            'DO%': 'do_percent',
            'Temperature': 'temperature',
            'Turbidity': 'turbidity',
            'Salinity': 'salinity',
            'DTWT': 'dtwt',
            'Volume': 'volume',
            'Dilution': 'dilution',
            'Start Date/Time': 'start_datetime',
            'Atm_Pressure_mb': 'atm_pressure',
            'ORP_mV': 'orp_mv'
        }

        # List of possible analysis columns
        analysis_names = [
            'DOC', 'TDN', 'Anions', 'Cations', 'NO3+NO2', 'NO2', 'NH4',
            'PO4/SRP', 'SiO2', 'TN', 'TP', 'TDP', 'TSS', 'PC/PN',
            'Chl a', 'EEMs', 'Gases - GC', 'ICPOES', 'Additional'
        ]

        # Find columns that match our field mappings
        column_mapping = {}
        for col in log_data.columns:
            col_str = str(col).strip()
            # Check for exact matches first
            if col_str in field_mappings:
                column_mapping[col] = field_mappings[col_str]
                continue
            # Then check for partial matches
            for key, value in field_mappings.items():
                if key.lower() in col_str.lower() or col_str.lower() in key.lower():
                    column_mapping[col] = value
                    break

        # Find analysis columns
        analysis_columns = {}
        for col in log_data.columns:
            col_str = str(col).strip()
            for analysis in analysis_names:
                if col_str == analysis or col_str.lower() == analysis.lower():
                    analysis_columns[col] = analysis
                    break

        print("Column mapping for Log Book:")
        for col, field in column_mapping.items():
            print(f"  {col} -> {field}")

        print("Analysis columns in Log Book:", list(analysis_columns.keys()))

        # Process each row of data
        for idx, row in log_data.iterrows():
            # Skip rows that are completely empty
            if row.isnull().all():
                continue

            # Initialize a sample dictionary
            sample = {}

            # Extract UNH# first to check if it's the only field populated
            unh_id = None
            for col, field_name in column_mapping.items():
                if field_name == 'unh_id':
                    value = row[col]
                    if pd.notna(value) and str(value).strip():
                        unh_id = str(value).strip()
                        sample['unh_id'] = unh_id
                    break

            # Skip rows where only UNH# is populated
            if unh_id and len(sample) == 1:
                has_other_data = False
                for col, field_name in column_mapping.items():
                    if field_name != 'unh_id':
                        value = row[col]
                        if pd.notna(value) and str(value).strip():
                            has_other_data = True
                            break

                if not has_other_data:
                    print(f"Skipping row with only UNH#: {unh_id}")
                    continue

            # Extract the rest of the values
            for col, field_name in column_mapping.items():
                if field_name != 'unh_id':  # Already processed UNH# above
                    value = row[col]
                    if pd.notna(value) and str(value).strip():
                        # Convert to string but handle special types
                        if isinstance(value, (datetime.datetime, datetime.date)):
                            sample[field_name] = value.strftime('%Y-%m-%d')
                        elif isinstance(value, datetime.time):
                            sample[field_name] = value.strftime('%H:%M:%S')
                        else:
                            sample[field_name] = str(value).strip()

            # Extract analysis requirements
            sample['analyses'] = {}
            for col, analysis_name in analysis_columns.items():
                value = row[col]
                is_required = False
                if pd.notna(value):
                    value_str = str(value).upper().strip()
                    if value_str == 'X' or value_str == 'TRUE' or value_str == '1' or value_str == 'Y':
                        is_required = True
                sample['analyses'][analysis_name] = is_required

            # Only add samples that have at least a sample name or UNH ID
            if ('sample_name' in sample and sample['sample_name']) or ('unh_id' in sample and sample['unh_id']):
                samples.append(sample)

        print(f"Extracted {len(samples)} valid samples from Log Book")
        if samples:
            print("First sample from Log Book:")
            for key, value in samples[0].items():
                if key != 'analyses':
                    print(f"  {key}: {value}")
            print("  Analyses requested:")
            for analysis, requested in samples[0]['analyses'].items():
                if requested:
                    print(f"    {analysis}")

        return samples

    def populate_logbook_preview(self, log_data, samples):
        """Populate the preview treeviews with data from the Log Book Excel file."""
        # Clear the current content of both treeviews
        for item in self.project_tree.get_children():
            self.project_tree.delete(item)

        for item in self.sample_tree.get_children():
            self.sample_tree.delete(item)

        # For Log Book format, we don't have separate project info
        # So we'll create a simple project info display with basic stats
        project_info = [
            {"Field": "File Type", "Value": "Log Book Format"},
            {"Field": "Samples Found", "Value": str(len(samples))},
            {"Field": "Project", "Value": self.project_entry.get() or "Not specified"}
        ]

        # Find unique projects in the data
        unique_projects = set()
        for sample in samples:
            if 'project' in sample and sample['project']:
                unique_projects.add(sample['project'])

        if unique_projects:
            project_info.append({"Field": "Projects in File", "Value": ", ".join(unique_projects)})

        # Configure project tree columns
        project_columns = ["Field", "Value"]
        self.project_tree["columns"] = project_columns
        self.project_tree["show"] = "headings"

        for col in project_columns:
            self.project_tree.heading(col, text=col)
            self.project_tree.column(col, width=150, minwidth=50)

        # Add project info to the treeview
        for info in project_info:
            self.project_tree.insert("", "end", values=[info["Field"], info["Value"]])

        # Configure and populate sample treeview
        # Check if we have extracted samples
        if samples:
            # Create a preview DataFrame with properly extracted data
            preview_data = []
            for sample in samples:
                sample_row = {}
                # Add basic fields
                for key, value in sample.items():
                    if key != 'analyses':
                        sample_row[key] = value

                # Add analysis fields
                for analysis, required in sample.get('analyses', {}).items():
                    sample_row[analysis] = 'X' if required else ''

                preview_data.append(sample_row)

            # Convert to DataFrame for easier display
            if preview_data:
                preview_df = pd.DataFrame(preview_data)

                # Get preview columns
                preview_columns = preview_df.columns.tolist()

                # Configure columns for sample tree based on extracted data
                self.sample_tree["columns"] = preview_columns
                self.sample_tree["show"] = "headings"

                # Configure each column
                for col in preview_columns:
                    self.sample_tree.heading(col, text=str(col))
                    self.sample_tree.column(col, width=100, minwidth=50)

                # Add extracted data to sample tree
                for _, row in preview_df.iterrows():
                    values = [str(val) if pd.notna(val) else "" for val in row]
                    self.sample_tree.insert("", "end", values=values)

                print(f"Sample preview populated with {len(preview_df)} rows and {len(preview_columns)} columns")
                return

        # If we couldn't extract samples properly, show the raw data
        print("Using raw Log Book data for preview")
        sample_columns = log_data.columns.tolist()

        self.sample_tree["columns"] = sample_columns
        self.sample_tree["show"] = "headings"

        for col in sample_columns:
            self.sample_tree.heading(col, text=str(col))
            self.sample_tree.column(col, width=100, minwidth=50)

        for _, row in log_data.iterrows():
            values = [str(val) if pd.notna(val) else "" for val in row]
            self.sample_tree.insert("", "end", values=values)

        print(f"Raw Log Book preview populated with {len(log_data)} rows and {len(sample_columns)} columns")

    def import_logbook_data(self):
        """Import the data from the selected Log Book Excel file into the Access database."""
        file_path = self.logbook_file_path_var.get()
        if not file_path:
            messagebox.showwarning("No File Selected", "Please select a Log Book Excel file first.")
            return

        try:
            # Read the Excel file
            self.import_status_var.set("Loading Log Book file for import...")
            log_data = self.read_logbook_excel(file_path)

            if log_data is None or log_data.empty:
                self.import_status_var.set("Error: Could not read the Log Book file or it's empty.")
                return

            # Confirm import
            confirm = messagebox.askyesno(
                "Confirm Import",
                f"Are you sure you want to import samples from this Log Book file?"
            )

            if not confirm:
                self.import_status_var.set("Import cancelled by user.")
                return

            # Perform the import
            self.import_status_var.set("Importing data from Log Book...")
            success = self.perform_logbook_import(log_data)

            if success:
                self.import_status_var.set(f"Successfully imported samples from Log Book.")
                # Reload the main data table
                self.data = self._load_data_from_database()
                self.show_all()
            else:
                self.import_status_var.set("Error during Log Book import. See console for details.")

        except Exception as e:
            error_message = f"Error importing Log Book file: {str(e)}"
            print(error_message)
            print(traceback.format_exc())
            messagebox.showerror("Import Error", error_message)
            self.import_status_var.set("Error during import. See console for details.")

    def perform_logbook_import(self, log_data):
        """Perform the actual import of Log Book data into the Access database."""
        conn = self._get_db_connection()
        if not conn:
            return False

        cursor = conn.cursor()
        imported_count = 0

        try:
            # Begin transaction
            if conn.autocommit:
                conn.autocommit = False

            # Extract sample data
            samples = self.extract_logbook_data(log_data)

            if not samples:
                messagebox.showwarning("No Samples", "No valid samples found in the Log Book file.")
                return False

            print(f"Found {len(samples)} samples to import from Log Book")

            # Process each sample
            success_count = 0
            skipped_count = 0
            for sample in samples:
                print(f"Processing Log Book sample: {sample.get('sample_name', 'Unknown')}")

                # Check if UNH# already exists
                unh_id = sample.get('unh_id', '')
                if unh_id:
                    # Check if this UNH# already exists
                    if self._check_unh_exists(cursor, unh_id):
                        print(f"Skipping existing UNH# {unh_id}")
                        skipped_count += 1
                        continue

                # Get project info directly from the sample
                project_info = {
                    'user_project_name': sample.get('project', 'Default Project'),
                    'project_name': sample.get('project', ''),
                    'sub_project': sample.get('sub_project', ''),
                    'sub_projecta': sample.get('sub_projecta', '')
                }

                success = self._insert_logbook_sample(cursor, project_info, sample)

                if success:
                    # Insert into WRRC sample analysis requested
                    self._insert_logbook_analysis(cursor, sample)
                    success_count += 1

            # Commit the transaction
            conn.commit()
            print(f"Successfully imported {success_count} samples, skipped {skipped_count} existing samples.")

            if success_count > 0 or skipped_count > 0:
                messagebox.showinfo("Import Result",
                                    f"Import completed:\n- {success_count} samples imported\n- {skipped_count} samples skipped (already exist)")
                return True
            else:
                messagebox.showwarning("Import Warning", "No samples were imported. Check the console for details.")
                return False

        except Exception as e:
            # Rollback in case of error
            conn.rollback()
            error_message = f"Error during Log Book import: {str(e)}"
            print(error_message)
            print(traceback.format_exc())
            messagebox.showerror("Import Error", error_message)
            return False

        finally:
            cursor.close()
            conn.close()

    def _check_unh_exists(self, cursor, unh_id):
        """Check if a UNH ID already exists in the database."""
        try:
            query = "SELECT COUNT(*) FROM [WRRC sample info] WHERE [UNH#] = ?"
            cursor.execute(query, (unh_id,))
            count = cursor.fetchone()[0]
            return count > 0
        except Exception as e:
            print(f"Error checking if UNH# exists: {str(e)}")
            return False

    def _insert_logbook_sample(self, cursor, project_info, sample):
        """Insert a sample from Log Book into the WRRC sample info table."""
        try:
            # Extract sample information
            unh_id = sample.get('unh_id', '')
            sample_name = sample.get('sample_name', '')
            sample_type = sample.get('sample_type', '')
            field_notes = sample.get('field_notes', '')

            # Handle date and time formatting
            collection_date = sample.get('collection_date', '')
            collection_time = sample.get('collection_time', '')

            # Convert date/time to proper format if needed
            if collection_date:
                try:
                    # If it's already a datetime object
                    if isinstance(collection_date, (datetime.datetime, datetime.date)):
                        collection_date = collection_date.strftime('%Y-%m-%d')
                    # If it's a string, try to parse it
                    elif isinstance(collection_date, str):
                        try:
                            # Try multiple date formats
                            date_formats = [
                                '%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y',
                                '%m-%d-%Y', '%d-%m-%Y', '%m.%d.%Y', '%d.%m.%Y'
                            ]

                            parsed_date = None
                            for fmt in date_formats:
                                try:
                                    parsed_date = datetime.datetime.strptime(collection_date, fmt)
                                    break
                                except:
                                    continue

                            if parsed_date:
                                collection_date = parsed_date.strftime('%Y-%m-%d')
                            else:
                                print(f"Warning: Could not parse date format for {collection_date}")
                        except Exception as date_err:
                            print(f"Date parsing error: {date_err}")
                except Exception as date_conv_err:
                    print(f"Date conversion error: {date_conv_err}")

            # Format collection time if needed
            if collection_time:
                try:
                    # Handle different time formats
                    if isinstance(collection_time, datetime.time):
                        collection_time = collection_time.strftime('%H:%M:%S')
                    elif isinstance(collection_time, str) and collection_time.strip():
                        # Try to standardize time format
                        time_formats = ['%H:%M:%S', '%I:%M:%S %p', '%I:%M %p', '%H:%M']
                        for fmt in time_formats:
                            try:
                                parsed_time = datetime.datetime.strptime(collection_time, fmt).time()
                                collection_time = parsed_time.strftime('%H:%M:%S')
                                break
                            except:
                                continue
                    else:
                        collection_time = None
                except Exception as time_err:
                    print(f"Time parsing error: {time_err}")
                    collection_time = None
            else:
                collection_time = None

            # Use project info from the sample itself
            project = project_info.get('user_project_name', '')
            if not project:
                project = "Default Project"
                print(f"Warning: Using default project name because no project specified in Log Book")

            # Get project-related fields from sample
            sub_project = project_info.get('sub_project', '')
            sub_projecta = project_info.get('sub_projecta', '')
            sub_projectb = sample.get('sub_projectb', '')

            # Get additional measurements if available
            ph = sample.get('ph', '')
            conductivity = sample.get('cond', '')
            spec_cond = sample.get('spec_cond', '')
            do_conc = sample.get('do_conc', '')
            do_percent = sample.get('do_percent', '')
            temperature = sample.get('temperature', '')
            salinity = sample.get('salinity', '')

            # Validate numeric fields
            for field_name in ['ph', 'cond', 'spec_cond', 'do_conc', 'do_percent', 'temperature', 'salinity']:
                value = sample.get(field_name)
                if value is not None:
                    if value == '' or (isinstance(value, str) and value.lower() in ['s', 'na', 'n/a']):
                        sample[field_name] = None
                    elif isinstance(value, str):
                        try:
                            sample[field_name] = float(value)
                        except ValueError:
                            print(f"Warning: Invalid {field_name} value '{value}' - setting to NULL")
                            sample[field_name] = None

            print(f"Log Book sample info: UNH ID={unh_id}, Name={sample_name}, Date={collection_date}")
            print(f"Using project from Log Book: '{project}'")

            # First check if the sample already exists (should be redundant with earlier check but safer)
            check_query = """
            SELECT COUNT(*) FROM [WRRC sample info] 
            WHERE [UNH#] = ?
            """
            cursor.execute(check_query, (unh_id,))
            count = cursor.fetchone()[0]
            if count > 0:
                print(f"Sample with UNH# {unh_id} already exists in database. Skipping.")
                return False

            # Build a field mapping from variables to database column names
            field_mapping = {
                'UNH#': unh_id if unh_id else None,
                'Sample_Name': sample_name if sample_name else "Unknown Sample",
                'Collection_Date': collection_date if collection_date else None,
                'Project': project if project else "Default Project",
                'Sub_Project': sub_project if sub_project else None,
                'Sub_ProjectA': sub_projecta if sub_projecta else None,
                'Sub_ProjectB': sub_projectb if sub_projectb else None,
                'Sample_Type': sample_type if sample_type else None,
                'Field_Notes': field_notes if field_notes else None,
                'pH': sample.get('ph') if sample.get('ph') is not None else None,
                'Cond': sample.get('cond') if sample.get('cond') is not None else None,
                'Spec_Cond': sample.get('spec_cond') if sample.get('spec_cond') is not None else None,
                'DO_Conc': sample.get('do_conc') if sample.get('do_conc') is not None else None,
                'DO%': sample.get('do_percent') if sample.get('do_percent') is not None else None,
                'Temperature': sample.get('temperature') if sample.get('temperature') is not None else None,
                'Salinity': sample.get('salinity') if sample.get('salinity') is not None else None
            }

            # Only add Collection_Time if it's not empty
            if collection_time:
                field_mapping['Collection_Time'] = collection_time

            # Filter out None values
            fields = {k: v for k, v in field_mapping.items() if v is not None}

            # Build columns and parameters for SQL query
            columns = list(fields.keys())

            # Handle special characters in column names
            formatted_columns = []
            for col in columns:
                if col == 'DO%':
                    formatted_columns.append('[DO%]')
                else:
                    formatted_columns.append(f'[{col}]')

            placeholders = ['?'] * len(columns)
            values = list(fields.values())

            # Construct and execute the query
            query = f"INSERT INTO [WRRC sample info] ({', '.join(formatted_columns)}) VALUES ({', '.join(placeholders)})"

            print(f"Log Book sample insert query: {query}")
            print(f"Parameters: {values}")

            cursor.execute(query, values)

            print(f"Inserted Log Book sample info for: {sample_name} (UNH# {unh_id})")
            return True

        except Exception as e:
            print(f"Error inserting Log Book sample info: {str(e)}")
            print(traceback.format_exc())
            raise

    def _insert_logbook_analysis(self, cursor, sample):
        """Insert analysis data from Log Book into the WRRC sample analysis requested table."""
        try:
            # Get UNH ID
            unh_id = sample.get('unh_id', '')

            # Make sure we have the minimum required data
            if not unh_id:
                print("Cannot insert analysis request: Missing UNH ID")
                return False

            # Get analysis requirements
            analyses = sample.get('analyses', {})

            # Get additional fields for analysis
            containers = sample.get('containers', '')
            filtered = sample.get('filtered', '')
            preservation = sample.get('preservation', '')
            filter_volume = sample.get('filter_volume', '')

            # Create mapping between Excel analysis names and database column names
            analysis_mapping = {
                'DOC': 'DOC',
                'TDN': 'TDN',
                'Anions': 'Anions',
                'Cations': 'Cations',
                'NO3+NO2': 'NO3AndNO2',
                'NO2': 'NO2',
                'NH4': 'NH4',
                'PO4/SRP': 'PO4OrSRP',
                'SiO2': 'SiO2',
                'TN': 'TN',
                'TP': 'TP',
                'TDP': 'TDP',
                'TSS': 'TSS',
                'PC/PN': 'PCAndPN',
                'Chl a': 'Chl_a',
                'EEMs': 'EEMs',
                'Gases - GC': 'Gases_GC',
                'ICPOES': 'ICPOES',
                'Additional': 'Additional'
            }

            # Log the analyses that are marked as required
            required_analyses = [analysis for analysis, is_required in analyses.items() if is_required]
            print(f"Required analyses for UNH# {unh_id}: {required_analyses}")

            # Prepare columns and values
            columns = ["[UNH#]"]
            values = [str(unh_id)]

            # Add additional fields if they exist
            if containers:
                columns.append("[Containers]")
                values.append(str(containers))

            if filtered:
                columns.append("[Filtered]")
                values.append(str(filtered))

            if preservation:
                columns.append("[Preservation]")
                values.append(str(preservation))

            if filter_volume:
                columns.append("[Filter_Volume]")
                values.append(str(filter_volume))

            # Add analysis fields
            for analysis, is_required in analyses.items():
                if is_required:
                    db_column = analysis_mapping.get(analysis)
                    if db_column:
                        columns.append(f"[{db_column}]")
                        values.append("required")

            # If we have any data to insert
            if len(columns) > 1:  # At least UNH# plus one more field
                # Build the SQL query
                query = f"INSERT INTO [WRRC sample analysis requested] ({', '.join(columns)}) VALUES ({', '.join(['?'] * len(values))})"

                try:
                    cursor.execute(query, values)
                    print(f"Inserted analysis request for UNH# {unh_id} with {len(columns) - 1} fields")
                    return True
                except Exception as e:
                    print(f"Error inserting analysis request for UNH# {unh_id}: {str(e)}")
                    raise
            else:
                print(f"No analysis fields to insert for UNH# {unh_id}")
                return False

        except Exception as e:
            print(f"Error inserting Log Book analysis data: {str(e)}")
            raise

    # Update these existing methods to work with the new dual-import system

    def preview_excel_data(self):
        """Preview the data from the selected Sample Submission Excel file."""
        file_path = self.submission_file_path_var.get()
        if not file_path:
            messagebox.showwarning("No File Selected", "Please select a Sample Submission Excel file first.")
            return

        try:
            # Load the Excel file
            self.import_status_var.set("Loading file for preview...")

            # Try to read the Excel file
            project_df, sample_df = self.read_sample_submission_excel(file_path)

            if project_df is None or sample_df is None:
                self.import_status_var.set("Error: Could not read the expected sheets from the Excel file.")
                return

            # Extract project info
            self.current_project_info = self.extract_project_info(project_df)

            # Extract sample data for validation
            samples = self.extract_sample_data(sample_df)

            # Update the status
            self.import_status_var.set(f"Preview ready. Found {len(samples)} samples.")

            # Display Excel project name in status but don't override textbox
            excel_project = self.current_project_info.get('project_name', 'Not specified')
            self.import_status_var.set(f"Preview ready. Found {len(samples)} samples. Excel Project: {excel_project}")

            # Populate the preview treeviews
            self.populate_preview_treeviews(project_df, sample_df)

        except Exception as e:
            error_message = f"Error previewing Excel file: {str(e)}"
            print(error_message)
            print(traceback.format_exc())
            messagebox.showerror("Preview Error", error_message)
            self.import_status_var.set("Error previewing file. See console for details.")

    def import_excel_data(self):
        """Import the data from the selected Sample Submission Excel file into the Access database."""
        file_path = self.submission_file_path_var.get()
        if not file_path:
            messagebox.showwarning("No File Selected", "Please select a Sample Submission Excel file first.")
            return

        try:
            # Read the Excel file
            self.import_status_var.set("Loading file for import...")
            project_df, sample_df = self.read_sample_submission_excel(file_path)

            if project_df is None or sample_df is None:
                self.import_status_var.set("Error: Could not read the expected sheets from the Excel file.")
                return

            # Confirm import
            confirm = messagebox.askyesno(
                "Confirm Import",
                f"Are you sure you want to import {len(sample_df)} samples from this file?"
            )

            if not confirm:
                self.import_status_var.set("Import cancelled by user.")
                return

            # Perform the import
            self.import_status_var.set("Importing data...")
            success = self.perform_import(project_df, sample_df)

            if success:
                self.import_status_var.set(f"Successfully imported {len(sample_df)} samples.")
                # Reload the main data table
                self.data = self._load_data_from_database()
                self.show_all()
            else:
                self.import_status_var.set("Error during import. See console for details.")

        except Exception as e:
            error_message = f"Error importing Excel file: {str(e)}"
            print(error_message)
            print(traceback.format_exc())
            messagebox.showerror("Import Error", error_message)
            self.import_status_var.set("Error during import. See console for details.")

    def perform_import(self, project_df, sample_df):
        """Perform the actual import of data into the Access database."""
        # Check if project name is provided
        project_name = self.project_entry.get().strip()
        if not project_name:
            messagebox.showerror("Missing Project", "Please enter a Project name.")
            return False

        conn = self._get_db_connection()
        if not conn:
            return False

        cursor = conn.cursor()
        imported_count = 0
        skipped_count = 0

        try:
            # Begin transaction
            if conn.autocommit:
                conn.autocommit = False

            # Extract project info from the Excel file
            excel_project_info = self.extract_project_info(project_df)

            # Create a separate dictionary for database import
            # This preserves the user's entry for Project
            project_info = {
                # Copy Excel project info
                **excel_project_info,
                # Override with user-entered project name
                'user_project_name': project_name
            }

            # Log what we're using
            print(f"Using user-entered Project name: '{project_name}'")
            if 'project_name' in excel_project_info:
                print(f"Using Excel-derived Sub_Project name: '{excel_project_info.get('project_name', '')}'")

            # Extract sample data
            samples = self.extract_sample_data(sample_df)

            if not samples:
                messagebox.showwarning("No Samples", "No valid samples found in the Excel file.")
                return False

            print("Project information:", project_info)
            print(f"Found {len(samples)} samples to import")

            # Process each sample
            success_count = 0
            for sample in samples:
                print(f"Processing sample: {sample.get('sample_name', 'Unknown')}")

                # Check if sample has UNH# and if it already exists
                unh_id = sample.get('unh_id', '')
                if unh_id and self._check_unh_exists(cursor, unh_id):
                    print(f"Skipping existing UNH# {unh_id}")
                    skipped_count += 1
                    continue

                # Insert into WRRC sample info
                success = self._insert_sample_info(cursor, project_info, sample)

                if success:
                    # Insert into WRRC sample analysis requested
                    self._insert_sample_analysis_requested(cursor, sample)
                    success_count += 1

            # Commit the transaction
            conn.commit()
            print(f"Successfully imported {success_count} samples, skipped {skipped_count} existing samples.")

            if success_count > 0 or skipped_count > 0:
                messagebox.showinfo("Import Result",
                                    f"Import completed:\n- {success_count} samples imported\n- {skipped_count} samples skipped (already exist)")
                return True
            else:
                messagebox.showwarning("Import Warning", "No samples were imported. Check the console for details.")
                return False

        except Exception as e:
            # Rollback in case of error
            conn.rollback()
            error_message = f"Error during import: {str(e)}"
            print(error_message)
            print(traceback.format_exc())
            messagebox.showerror("Import Error", error_message)
            return False

        finally:
            cursor.close()
            conn.close()

    def _insert_sample_info(self, cursor, project_info, sample):
        """Insert a record into the WRRC sample info table."""
        try:
            # Import datetime up front
            import datetime

            # Extract sample information
            unh_id = sample.get('unh_id', '')
            sample_name = sample.get('sample_name', '')
            sample_type = sample.get('sample_type', '')
            field_notes = sample.get('field_notes', '')

            # Handle date and time formatting to prevent data type mismatches
            collection_date = sample.get('collection_date', '')
            collection_time = sample.get('collection_time', '')

            # Convert date/time to proper format if needed
            if collection_date:
                try:
                    # If it's already a datetime object
                    if isinstance(collection_date, (datetime.datetime, datetime.date)):
                        collection_date = collection_date.strftime('%Y-%m-%d')
                    # If it's a string, try to parse it
                    elif isinstance(collection_date, str):
                        try:
                            # Try multiple date formats
                            date_formats = [
                                '%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y',
                                '%m-%d-%Y', '%d-%m-%Y', '%m.%d.%Y', '%d.%m.%Y'
                            ]

                            parsed_date = None
                            for fmt in date_formats:
                                try:
                                    parsed_date = datetime.datetime.strptime(collection_date, fmt)
                                    break
                                except:
                                    continue

                            if parsed_date:
                                # Convert to Access-compatible date format (yyyy-mm-dd)
                                collection_date = parsed_date.strftime('%Y-%m-%d')
                            else:
                                # If parsing fails, keep as is but print warning
                                print(f"Warning: Could not parse date format for {collection_date}")
                        except Exception as date_err:
                            print(f"Date parsing error: {date_err}")
                            # Keep original value if conversion fails
                except Exception as date_conv_err:
                    print(f"Date conversion error: {date_conv_err}")
                    # Keep original value if conversion fails

            # Format collection time if needed
            if collection_time:
                try:
                    # Handle different time formats
                    if isinstance(collection_time, datetime.time):
                        collection_time = collection_time.strftime('%H:%M:%S')
                    elif isinstance(collection_time, str) and collection_time.strip():
                        # Try to standardize time format
                        time_formats = ['%H:%M:%S', '%I:%M:%S %p', '%I:%M %p', '%H:%M']
                        for fmt in time_formats:
                            try:
                                parsed_time = datetime.datetime.strptime(collection_time, fmt).time()
                                collection_time = parsed_time.strftime('%H:%M:%S')
                                break
                            except:
                                continue
                    else:
                        collection_time = None  # Empty or non-string time value
                except Exception as time_err:
                    print(f"Time parsing error: {time_err}")
                    collection_time = None
            else:
                collection_time = None

            # IMPORTANT: Use the user-entered project name for the Project field
            # Get it directly from project_info['user_project_name']
            project = project_info.get('user_project_name', '')

            # Make sure we have a valid project name or use a default
            if not project or project.strip() == '':
                project = "Default Project"
                print(f"Warning: Using default project name because user entry is empty")
            else:
                print(f"Using user-entered project name: '{project}'")

            # Use the Excel project name for Sub_Project
            sub_project = project_info.get('project_name', '')
            print(f"Using Excel project name for Sub_Project: '{sub_project}'")

            # If no sub_project found, use a default value to avoid empty string error
            if not sub_project or sub_project.strip() == '':
                sub_project = "Default Sub Project"

            # Get contact name for Sub_ProjectA
            proj_manager = project_info.get('contact_name', '')

            # Make sure proj_manager is not empty
            if not proj_manager or proj_manager.strip() == '':
                proj_manager = "Unknown"  # Set a default value

            # Get additional measurements if available
            ph = sample.get('ph', '')
            conductivity = sample.get('cond', '')
            spec_cond = sample.get('spec_cond', '')
            do_conc = sample.get('do_conc', '')
            do_percent = sample.get('do_percent', '')
            temperature = sample.get('temperature', '')
            salinity = sample.get('salinity', '')

            # Validate numeric fields to prevent type errors
            if salinity is not None:
                if salinity == '' or (isinstance(salinity, str) and salinity.lower() == 's'):
                    salinity = None  # Handle special case
                elif isinstance(salinity, str):
                    try:
                        salinity = float(salinity)
                    except ValueError:
                        print(f"Warning: Invalid salinity value '{salinity}' - setting to NULL")
                        salinity = None

            print(f"Sample info: UNH ID={unh_id}, Name={sample_name}, Date={collection_date}, Time={collection_time}")
            print(f"Project info: Project={project}, Sub_Project={sub_project}, Manager={proj_manager}")

            # First check if the sample already exists to avoid duplicate key error
            try:
                check_query = """
                SELECT COUNT(*) FROM [WRRC sample info] 
                WHERE Sample_Name = ? AND Collection_Date = ?
                """
                cursor.execute(check_query, (
                    str(sample_name) if sample_name else "Unknown Sample",
                    str(collection_date) if collection_date else None
                ))

                count = cursor.fetchone()[0]
                if count > 0:
                    print(
                        f"Warning: Sample {sample_name} with date {collection_date} already exists in database. Skipping.")
                    return False
            except Exception as check_err:
                print(f"Error checking for existing sample: {str(check_err)}")
                # Continue with insert attempt

            # Build a field mapping from our variables to the actual database column names
            field_mapping = {
                'UNH#': unh_id if unh_id else None,
                'Sample_Name': sample_name if sample_name else "Unknown Sample",
                'Collection_Date': collection_date if collection_date else None,
                'Project': project if project else "Default Project",
                'Sub_Project': sub_project if sub_project else "Default Sub Project",
                'Sub_ProjectA': proj_manager if proj_manager else "Unknown",
                'Sample_Type': sample_type if sample_type else None,
                'Field_Notes': field_notes if field_notes else None,
                'pH': ph if ph else None,
                'Cond': conductivity if conductivity else None,
                'Spec_Cond': spec_cond if spec_cond else None,
                'DO_Conc': do_conc if do_conc else None,
                'DO%': do_percent if do_percent else None,
                'Temperature': temperature if temperature else None,
                'Salinity': salinity if salinity is not None else None
            }

            # Only add Collection_Time if it's not empty
            if collection_time:
                field_mapping['Collection_Time'] = collection_time

            # Filter out None values to avoid issues
            fields = {k: v for k, v in field_mapping.items() if v is not None}

            # Build columns and parameters for SQL query
            columns = list(fields.keys())

            # Fix special characters in column names for SQL
            # In Access SQL, brackets are used to escape special characters and reserved words
            formatted_columns = []
            for col in columns:
                if col == 'DO%':
                    # Special handling for column names with special characters
                    formatted_columns.append('[DO%]')
                else:
                    formatted_columns.append(f'[{col}]')

            placeholders = ['?'] * len(columns)
            values = list(fields.values())

            # Construct the query with proper column escaping
            query = f"INSERT INTO [WRRC sample info] ({', '.join(formatted_columns)}) VALUES ({', '.join(placeholders)})"

            # Log the query and parameters for debugging
            print(f"Query: {query}")
            print(f"Parameters: {values}")

            # Execute the query
            cursor.execute(query, values)

            print(f"Inserted sample info for: {sample_name}")
            return True

        except Exception as e:
            print(f"Error inserting sample info: {str(e)}")
            raise

    def read_sample_submission_excel(self, file_path):
        """Read the sample submission Excel file and return two DataFrames for Project and Sample info."""
        try:
            # Check if both sheets exist in the Excel file
            xls = pd.ExcelFile(file_path)
            required_sheets = ["Project Information", "Sample Information"]

            if not all(sheet in xls.sheet_names for sheet in required_sheets):
                missing_sheets = [sheet for sheet in required_sheets if sheet not in xls.sheet_names]
                print(f"Missing sheets in Excel file: {missing_sheets}")
                messagebox.showerror(
                    "Invalid Excel File",
                    f"The Excel file is missing the following required sheets: {', '.join(missing_sheets)}"
                )
                return None, None

            # Read the Project Information sheet
            project_df = pd.read_excel(file_path, sheet_name="Project Information")

            # For Sample Information, we need to handle the multi-row header structure
            # First read without headers to see the structure
            raw_df = pd.read_excel(file_path, sheet_name="Sample Information", header=None)

            print("Raw Sample DataFrame first few rows (to debug header structure):")
            print(raw_df.head(5))

            # Now read the file with the correct header row
            # This uses row index 1 (second row) as the header
            sample_df = pd.read_excel(file_path, sheet_name="Sample Information", header=1)

            # Clean up the dataframes
            project_df = project_df.fillna("")
            sample_df = sample_df.fillna("")

            print("Sample DataFrame columns with correct header:")
            print(sample_df.columns.tolist())

            # Print first row to debug
            print("Sample DataFrame first row (data):")
            if not sample_df.empty:
                print(sample_df.iloc[0])

            return project_df, sample_df

        except Exception as e:
            error_message = f"Error reading Excel file: {str(e)}"
            print(error_message)
            print(traceback.format_exc())
            messagebox.showerror("Excel Import Error", error_message)
            return None, None

    def extract_sample_data(self, sample_df):
        """
        Extract sample information from the DataFrame.
        Returns a list of dictionaries, each containing a sample's information.
        """
        samples = []

        # Check if DataFrame is empty
        if sample_df.empty:
            return samples

        print(f"Sample DataFrame has {len(sample_df)} rows and {len(sample_df.columns)} columns")

        # Print column names to understand structure
        print("Actual column names in DataFrame:")
        for col in sample_df.columns:
            print(f"  {col}")

        # Define fields we're looking for
        field_mappings = {
            'UNH ID': 'unh_id',
            'Sample_Name': 'sample_name',
            'Collection_Date': 'collection_date',
            'Collection_Time': 'collection_time',
            'Sample_Type': 'sample_type',
            'Field_Notes': 'field_notes',
            'pH': 'ph',
            'Cond µS/cm': 'cond',
            'Spec_Cond µS/cm': 'spec_cond',
            'DO_Conc mg/L': 'do_conc',
            'DO%': 'do_percent',
            'Temperature degrees C': 'temperature',
            'Salinity (ppt)': 'salinity',
            'Number of containers': 'containers',
            'Filtered/unfiltered?': 'filtered',
            'Preservation': 'preservation',
            'Filter - Volume Filtered mL': 'filter_volume'
        }

        # List of possible analysis columns
        analysis_names = [
            'DOC', 'TDN', 'Anions', 'Cations', 'NO3+NO2', 'NO2', 'NH4',
            'PO4/SRP', 'SiO2', 'TN', 'TP', 'TDP', 'TSS', 'PC/PN',
            'Chl a', 'EEMs', 'Gases - GC', 'Additional'
        ]

        # Find columns that match our field mappings
        column_mapping = {}
        for col in sample_df.columns:
            col_str = str(col).strip()
            # Check for exact matches first
            if col_str in field_mappings:
                column_mapping[col] = field_mappings[col_str]
                continue
            # Then check for partial matches
            for key, value in field_mappings.items():
                if key.lower() in col_str.lower() or col_str.lower() in key.lower():
                    column_mapping[col] = value
                    break

        # Find analysis columns
        analysis_columns = {}
        for col in sample_df.columns:
            col_str = str(col).strip()
            for analysis in analysis_names:
                if col_str == analysis or col_str.lower() == analysis.lower():
                    analysis_columns[col] = analysis
                    break

        print("Column mapping:")
        for col, field in column_mapping.items():
            print(f"  {col} -> {field}")

        print("Analysis columns:", list(analysis_columns.keys()))

        # Process each row of data
        for idx, row in sample_df.iterrows():
            # Skip rows that are completely empty
            if row.isnull().all():
                continue

            # Skip header rows if they accidentally got included
            first_value = row.iloc[0] if len(row) > 0 else None
            if isinstance(first_value, str) and (
                    'UNH' in first_value or 'Sample' in first_value or 'ID' in first_value):
                print(f"Skipping header-like row: {first_value}")
                continue

            sample = {}

            # Extract values using our mapping
            for col, field_name in column_mapping.items():
                value = row[col]
                if pd.notna(value):
                    # Convert to string but handle special types
                    if isinstance(value, (datetime.datetime, datetime.date)):
                        sample[field_name] = value.strftime('%Y-%m-%d')
                    elif isinstance(value, datetime.time):
                        sample[field_name] = value.strftime('%H:%M:%S')
                    else:
                        sample[field_name] = str(value).strip()

            # Extract analysis requirements
            sample['analyses'] = {}
            for col, analysis_name in analysis_columns.items():
                value = row[col]
                is_required = False
                if pd.notna(value):
                    value_str = str(value).upper().strip()
                    if value_str == 'X' or value_str == 'TRUE' or value_str == '1' or value_str == 'Y':
                        is_required = True
                sample['analyses'][analysis_name] = is_required

            # Only add samples that have at least a sample name or UNH ID
            if ('sample_name' in sample and sample['sample_name']) or ('unh_id' in sample and sample['unh_id']):
                samples.append(sample)

        print(f"Extracted {len(samples)} valid samples")
        if samples:
            print("First sample:")
            for key, value in samples[0].items():
                if key != 'analyses':
                    print(f"  {key}: {value}")
            print("  Analyses requested:")
            for analysis, requested in samples[0]['analyses'].items():
                if requested:
                    print(f"    {analysis}")

        return samples

    def extract_project_info(self, project_df):
        """
        Extract project information by finding field names in the first column
        and their values in the second column.
        """
        project_info = {}

        # Check if the DataFrame is not empty
        if project_df.empty:
            print("Project DataFrame is empty")
            return project_info

        # Print DataFrame shape and columns for debugging
        print(f"Project DataFrame shape: {project_df.shape}")
        print(f"Project DataFrame columns: {project_df.columns.tolist()}")
        print("First few rows:")
        print(project_df.head().to_string())

        # Extract field names and values by searching for labels in the first column
        # and their corresponding values in the second column
        field_mapping = {
            "Contact Name": "contact_name",
            "Contact Address": "contact_address",
            "Contact Email": "contact_email",
            "Project Name": "project_name",
            "Project Location/Area": "project_location",
            "Brief Project Description": "project_description",
            "Date samples shipped": "shipment_date"
        }

        # Find the column names (they might vary)
        first_col = project_df.columns[0] if len(project_df.columns) > 0 else None
        second_col = project_df.columns[1] if len(project_df.columns) > 1 else None

        if first_col is None or second_col is None:
            print("Error: DataFrame doesn't have enough columns")
            return project_info

        # Iterate through each row to find the fields
        for i in range(len(project_df)):
            field_label = project_df.iloc[i][first_col]
            field_value = project_df.iloc[i][second_col] if i < len(project_df) else ""

            # Skip if the field label is not a string
            if not isinstance(field_label, str):
                continue

            # Clean up the field label by removing : and whitespace
            clean_label = field_label.strip().rstrip(':')

            # Print each label for debugging
            print(f"Checking label: '{field_label}' -> cleaned to -> '{clean_label}'")

            # Match with our field mapping
            for key, mapped_name in field_mapping.items():
                if clean_label.lower() == key.lower():
                    project_info[mapped_name] = str(field_value).strip()
                    print(f"✅ Matched '{clean_label}' -> '{mapped_name}' with value: '{field_value}'")
                    break

        # Print what we found for debugging
        print("Extracted project info:")
        for k, v in project_info.items():
            print(f"  {k}: {v}")

        return project_info

    def _insert_sample_analysis_requested(self, cursor, sample):
        """Insert a record into the WRRC sample analysis requested table."""
        try:
            # Get basic sample information
            unh_id = sample.get('unh_id', '')  # We'll now use UNH# instead of Sample_Name

            # Make sure we have the minimum required data
            if not unh_id:
                print("Cannot insert analysis request: Missing UNH ID")
                return False

            # Get analysis requirements
            analyses = sample.get('analyses', {})

            # Get additional fields
            containers = sample.get('containers', '')
            filtered = sample.get('filtered', '')
            preservation = sample.get('preservation', '')
            filter_volume = sample.get('filter_volume', '')
            field_notes = sample.get('field_notes', '')
            sample_type = sample.get('sample_type', '')
            collection_date = sample.get('collection_date', '')
            collection_time = sample.get('collection_time', '')

            # Create mapping between Excel analysis names and database column names
            analysis_mapping = {
                'DOC': 'DOC',
                'TDN': 'TDN',
                'Anions': 'Anions',
                'Cations': 'Cations',
                'NO3+NO2': 'NO3AndNO2',
                'NO2': 'NO2',
                'NH4': 'NH4',
                'PO4/SRP': 'PO4OrSRP',
                'SiO2': 'SiO2',
                'TN': 'TN',
                'TP': 'TP',
                'TDP': 'TDP',
                'TSS': 'TSS',
                'PC/PN': 'PCAndPN',
                'Chl a': 'Chl_a',
                'EEMs': 'EEMs',
                'Gases - GC': 'Gases_GC',
                'Additional': 'Additional'
            }

            # Log the analyses that are marked as required
            required_analyses = [analysis for analysis, is_required in analyses.items() if is_required]
            print(f"Required analyses for UNH# {unh_id}: {required_analyses}")

            # Prepare columns and values
            columns = ["[UNH#]"]
            values = [str(unh_id)]

            # Add additional fields if they exist
            if containers:
                columns.append("[Containers]")
                values.append(str(containers))

            if filtered:
                columns.append("[Filtered]")
                values.append(str(filtered))

            if preservation:
                columns.append("[Preservation]")
                values.append(str(preservation))

            if filter_volume:
                columns.append("[Filter_Volume]")
                values.append(str(filter_volume))

            if field_notes:
                columns.append("[Field_Notes]")
                values.append(str(field_notes))

            if sample_type:
                columns.append("[Sample_Type]")
                values.append(str(sample_type))

            if collection_date:
                columns.append("[Collection_Date]")
                values.append(str(collection_date))

            if collection_time:
                columns.append("[Collection_Time]")
                values.append(str(collection_time))

            # Add analysis fields
            for analysis, is_required in analyses.items():
                if is_required:
                    db_column = analysis_mapping.get(analysis)
                    if db_column:
                        columns.append(f"[{db_column}]")
                        values.append("required")

            # If we have any data to insert
            if len(columns) > 1:  # At least UNH# plus one more field
                # Build the SQL query
                query = f"INSERT INTO [WRRC sample analysis requested] ({', '.join(columns)}) VALUES ({', '.join(['?'] * len(values))})"

                try:
                    cursor.execute(query, values)
                    print(f"Inserted analysis request for UNH# {unh_id} with {len(columns) - 1} fields")
                    return True
                except Exception as e:
                    print(f"Error inserting analysis request for UNH# {unh_id}: {str(e)}")
                    raise
            else:
                print(f"No analysis fields to insert for UNH# {unh_id}")
                return False

        except Exception as e:
            print(f"Error inserting sample analysis requested: {str(e)}")
            raise

    def populate_preview_treeviews(self, project_df, sample_df):
        """Populate the preview treeviews with data from the Excel file."""
        # Configure and populate project treeview
        for item in self.project_tree.get_children():
            self.project_tree.delete(item)

        # Get column names for project dataframe
        project_columns = project_df.columns.tolist()

        # Configure columns for project tree
        self.project_tree["columns"] = project_columns

        # Create the column configuration - including the first column
        self.project_tree["show"] = "headings"  # This was likely the issue - default is "tree headings"

        for col in project_columns:
            self.project_tree.heading(col, text=str(col))
            self.project_tree.column(col, width=150, minwidth=50)

        # Add data to project tree
        for _, row in project_df.iterrows():
            values = [str(val) if pd.notna(val) else "" for val in row]
            self.project_tree.insert("", "end", values=values)

        print(f"Project preview populated with {len(project_df)} rows and {len(project_columns)} columns")
        print(f"Project columns: {project_columns}")

        # Configure and populate sample treeview
        for item in self.sample_tree.get_children():
            self.sample_tree.delete(item)

        # Get column names for sample dataframe
        sample_columns = sample_df.columns.tolist()

        # Configure columns for sample tree
        self.sample_tree["columns"] = sample_columns

        # Create the column configuration - including the first column
        self.sample_tree["show"] = "headings"  # This was likely the issue - default is "tree headings"

        for col in sample_columns:
            self.sample_tree.heading(col, text=str(col))
            self.sample_tree.column(col, width=100, minwidth=50)

        # Extract samples for preview
        samples = self.extract_sample_data(sample_df)

        # Check if we have extracted samples
        if samples:
            print(f"Extracted {len(samples)} samples for preview")

            # Create a special preview DataFrame for the treeview with properly extracted data
            preview_data = []
            for sample in samples:
                sample_row = {}
                # Add basic fields
                for key, value in sample.items():
                    if key != 'analyses':
                        sample_row[key] = value

                # Add analysis fields
                for analysis, required in sample.get('analyses', {}).items():
                    sample_row[analysis] = 'X' if required else ''

                preview_data.append(sample_row)

            # Convert to DataFrame for easier display
            if preview_data:
                preview_df = pd.DataFrame(preview_data)

                # Get preview columns
                preview_columns = preview_df.columns.tolist()

                # Configure columns for sample tree based on extracted data
                self.sample_tree["columns"] = preview_columns
                self.sample_tree["show"] = "headings"

                # Configure each column
                for col in preview_columns:
                    self.sample_tree.heading(col, text=str(col))
                    self.sample_tree.column(col, width=100, minwidth=50)

                # Add extracted data to sample tree
                for _, row in preview_df.iterrows():
                    values = [str(val) if pd.notna(val) else "" for val in row]
                    self.sample_tree.insert("", "end", values=values)

                print(f"Preview populated with {len(preview_df)} rows and {len(preview_columns)} columns")
                print(f"Preview columns: {preview_columns}")
                return

        # If we couldn't extract samples properly, show the raw data
        print("Using raw sample data for preview")
        for _, row in sample_df.iterrows():
            values = [str(val) if pd.notna(val) else "" for val in row]
            self.sample_tree.insert("", "end", values=values)

        print(f"Raw sample preview populated with {len(sample_df)} rows and {len(sample_columns)} columns")

if __name__ == "__main__":
    print("Starting the Sample Tracker App using Access database")
    app = SampleTrackerApp()
    app.mainloop()

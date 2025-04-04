import json
import tkinter as tk
from tkinter import ttk, messagebox
from jira import JIRA
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
import time
from openpyxl import load_workbook
import logging

# Set up logging
logging.basicConfig(filename='export_log.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Function to save credentials to a file
def save_credentials():
    credentials = {
        "jira_server": url_entry.get(),
        "jira_user": username_entry.get(),
        "jira_api_token": auth_token_entry.get(),
        "jql_query": jql_entry.get("1.0", tk.END).strip()
    }
    with open("credentials.json", "w") as f:
        json.dump(credentials, f)
    logging.info("Credentials saved successfully.")

# Function to load credentials from a file
def load_credentials():
    try:
        with open("credentials.json", "r") as f:
            credentials = json.load(f)
        url_entry.insert(0, credentials.get("jira_server", ""))
        username_entry.insert(0, credentials.get("jira_user", ""))
        auth_token_entry.insert(0, credentials.get("jira_api_token", ""))
        jql_entry.insert("1.0", credentials.get("jql_query", ""))
        logging.info("Credentials loaded successfully.")
    except FileNotFoundError:
        logging.warning("Credentials file not found.")

# Function to categorize dates into quarters
def categorize_quarters(date_str):
    date = datetime.strptime(date_str, '%m-%d-%Y')
    month = date.month
    if month in [11, 12, 1]:
        return 'Q1'
    elif month in [2, 3, 4]:
        return 'Q2'
    elif month in [5, 6, 7]:
        return 'Q3'
    else:
        return 'Q4'

# Function to categorize Bug Resolution values
def categorize_bug_resolution(value, status):
    fixed_code_change = ["Fixed: by Earlier Fix", "Fixed: Code/Build Change", "Fixed: Code/Build Change due to Design/Spec Change", "Fixed: Code Change", "Fixed: Code Change - Partner API Impact", "Fixed: Configuration Change", "Fixed: Database Change", "Fixed: Design Change", "Fixed: Design Changed", "Fixed: Documentation Change", "Fixed: Engine Change", "Fixed: Firmware Change", "Fixed: Hardware Change", "Fixed: Infrastructure Change", "Unknown", "Code Change Out of Scope", "Deployed", "New Requirement", "Other: See Comment", "Tool Change", "Transfer: HP Internal", "Unknown", "Service unavailability"]
    nad = ["Not a Defect: Other", "Not a Defect: As Designed", "Not a Defect: App/OS Error", "Not a Defect: Feature Not Ready", "Not a Defect: Design Limitation"]
    cannot_reproduce = ["Cannot Reproduce", "Root Cause Unknown: Cannot Reproduce"]
    duplicate = ["Duplicate", "Duplicate: of Bug"]
    false_defect = ["Test: Test/Build Mismatch", "Test: Test Case Error", "Test: Test Change", "Test: Test Env Error", "Test Setup Error", "Partner Education", "Partner setup issue", "External Error: Incorrect Testing", "Invalid", "Submission Error"]
    if value in fixed_code_change or (value == "" and status in ["Closed", "Accepted"]):
        return "Fixed: Code Change"
    elif value in nad:
        return "NAD"
    elif value in cannot_reproduce:
        return "Cannot Reproduce"
    elif value in duplicate:
        return "Duplicate"
    elif value in false_defect:
        return "False Defect"
    elif value == "" and status not in ["Closed", "Accepted"]:
        return "Open Defects"
    else:
        return "Other"

# Function to export issues based on user input
def export_issues():
    try:
        # Create a progress bar
        progress_var.set(0)
        progress_bar.grid(row=0, column=1, padx=5, pady=5)
        logging.info("Export process started.")
        
        # Save credentials
        save_credentials()
        
        # Update progress bar for authentication
        progress_var.set(10)
        progress_bar.update()
        
        # Load credentials from user input
        jira_server = url_entry.get()
        jira_user = username_entry.get()
        jira_api_token = auth_token_entry.get()
        jql_query = jql_entry.get("1.0", tk.END).strip()
        
        # Connect to Jira
        jira = JIRA(server=jira_server, basic_auth=(jira_user, jira_api_token))
        logging.info("Connected to Jira successfully.")
        
        # Update progress bar for Jira issues search
        progress_var.set(30)
        progress_bar.update()
        
        # Fetch issues from Jira
        issues = jira.search_issues(jql_query, maxResults=False)
        logging.info(f"Fetched {len(issues)} issues from Jira.")
        
        # Update progress bar for field values extraction
        progress_var.set(50)
        progress_bar.update()
        
        # Extract issue details
        issue_data = []
        total_issues = len(issues)
        start_time = time.time()
        for idx, issue in enumerate(issues):
            if open_issues_var.get() and issue.fields.status.name in ["Closed", "Accepted"]:
                continue
            created_date = datetime.strptime(issue.fields.created[:10], '%Y-%m-%d').strftime('%m-%d-%Y')
            updated_date = datetime.strptime(issue.fields.updated[:10], '%Y-%m-%d').strftime('%m-%d-%Y')
            severity_field = getattr(issue.fields, 'customfield_10605', None)
            severity_text = severity_field.value if severity_field else ''
            applicable_product_field = getattr(issue.fields, 'customfield_13550', None)
            applicable_product_text = ', '.join([product.value for product in applicable_product_field]) if applicable_product_field else ''
            bug_resolution_field = getattr(issue.fields, 'customfield_13555', None)
            bug_resolution_text = bug_resolution_field.value if bug_resolution_field else ''
            fixed_in_build_field = getattr(issue.fields, 'customfield_11412', None)
            fixed_in_build_text = fixed_in_build_field if fixed_in_build_field else ''
            encountered_by_field = getattr(issue.fields, 'customfield_13073', None)
            encountered_by_text = encountered_by_field.value if encountered_by_field else ''
            team_watch_list_field = getattr(issue.fields, 'customfield_31502', None)
            team_watch_list_text = ', '.join([team.value for team in team_watch_list_field]) if team_watch_list_field else ''
            deferred_products_field = getattr(issue.fields, 'customfield_16203', None)
            deferred_products_text = ', '.join([product.value for product in deferred_products_field]) if deferred_products_field else ''
            how_found_field = getattr(issue.fields, 'customfield_12900', None)
            how_found_text = how_found_field.value if how_found_field else ''
            found_in_fw_version_field = getattr(issue.fields, 'customfield_11405', None)
            found_in_fw_version_text = found_in_fw_version_field if found_in_fw_version_field else ''
            reproducibility_field = getattr(issue.fields, 'customfield_11408', None)
            reproducibility_text = reproducibility_field.value if reproducibility_field else ''
            resolved_date_field = getattr(issue.fields, 'resolutiondate', None)
            resolved_date_value = datetime.strptime(resolved_date_field[:10], '%Y-%m-%d').strftime('%m-%d-%Y') if resolved_date_field else ''
            issue_links = []
            for link in issue.fields.issuelinks:
                if hasattr(link, 'outwardIssue'):
                    link_html = f'<a href="/browse/{link.outwardIssue.key}" data-issue-key="{link.outwardIssue.key}" class="issue-link link-title">{link.outwardIssue.key}</a>'
                elif hasattr(link, 'inwardIssue'):
                    link_html = f'<a href="/browse/{link.inwardIssue.key}" data-issue-key="{link.inwardIssue.key}" class="issue-link link-title">{link.inwardIssue.key}</a>'
                else:
                    link_html = ''
                soup = BeautifulSoup(link_html, 'html.parser')
                link_element = soup.find('a')
                link_title = link_element['title'] if link_element and 'title' in link_element.attrs else link_element.get_text(strip=True)
                issue_links.append(link_title)
            issue_links_text = ', '.join(issue_links)
            issue_dict = {
                'Key': issue.key,
                'Summary': issue.fields.summary,
                'Status': issue.fields.status.name,
                'Assignee': issue.fields.assignee.displayName if issue.fields.assignee else 'Unassigned',
                'Reporter': issue.fields.reporter.displayName,
                'Created': created_date,
                'Updated': updated_date,
                'Priority': issue.fields.priority.name if issue.fields.priority else '',
                'Issue Type': issue.fields.issuetype.name,
                'Labels': ', '.join(issue.fields.labels),
                'Project': issue.fields.project.name,
                'Severity': severity_text,
                'Applicable Products': applicable_product_text,
                'Bug Resolution': bug_resolution_text,
                'Fixed in Build': fixed_in_build_text,
                'Encountered By': encountered_by_text,
                'Team Watch List': team_watch_list_text,
                'Deferred Products': deferred_products_text,
                'How Found': how_found_text,
                'Found in FW Version': found_in_fw_version_text,
                'Reproducibility': reproducibility_text,
                'Issue Links': issue_links_text,
                'Resolved': resolved_date_value
            }
            # Filter issue_dict based on selected fields
            filtered_issue_dict = {key: value for key, value in issue_dict.items() if field_vars[fields.index(key)].get()}
            issue_data.append(filtered_issue_dict)
            # Update progress bar
            progress_var.set(50 + (idx + 1) / total_issues * 50)
            progress_bar.update()
        
        # Create a DataFrame from the issue data
        df = pd.DataFrame(issue_data)
        
        # Categorize issues by quarters
        df['Quarter'] = df['Created'].apply(categorize_quarters)
        
        # Rename Severity levels
        df['Severity'] = df['Severity'].replace({
            'Critical': '1 Critical',
            'High': '2 High',
            'Medium': '3 Medium',
            'Low': '4 Low'
        })
        
        # Generate pivot table for Severity with Created dates categorized in Quarters
        severity_pivot_table = pd.pivot_table(df, values='Key', index=['Severity'], columns=['Quarter'], aggfunc='count', fill_value=0)
        
        # Filter issues with [BLOCK] or [Block] in the summary
        df['Blocker'] = df['Summary'].apply(lambda x: '[BLOCK]' in x or '[Block]' in x)
        
        # Generate pivot table for Blockers with Created dates categorized in Quarters
        blockers_pivot_table = pd.pivot_table(df[df['Blocker']], values='Key', index=['Blocker'], columns=['Quarter'], aggfunc='count', fill_value=0)
        
        # Categorize Bug Resolution values
        df['Bug Resolution Category'] = df.apply(lambda row: categorize_bug_resolution(row['Bug Resolution'], row['Status']), axis=1)
        
        # Generate pivot table for Bug Resolution with Created dates categorized in Quarters
        bug_resolution_pivot_table = pd.pivot_table(df, values='Key', index=['Bug Resolution Category'], columns=['Quarter'], aggfunc='count', fill_value=0)
        
        # Generate a timestamped file name
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_file = f'jira_issues_{timestamp}.xlsx'
        
        # Export the DataFrame and pivot tables to an Excel sheet
        with pd.ExcelWriter(excel_file) as writer:
            df.to_excel(writer, sheet_name='Issues', index=False)
            if severity_var.get():
                severity_pivot_table.to_excel(writer, sheet_name='Severity')
            if blockers_var.get():
                blockers_pivot_table.to_excel(writer, sheet_name='Blockers')
            if bug_resolution_var.get():
                bug_resolution_pivot_table.to_excel(writer, sheet_name='Bug Resolution')
        
        progress_bar.grid_forget()
        logging.info("Export process completed successfully.")
        messagebox.showinfo("Success", f"Issues exported to {excel_file} with selected pivot tables")
    except Exception as e:
        progress_bar.grid_forget()
        logging.error(f"Error during export process: {str(e)}")
        messagebox.showerror("Error", str(e))

# Function to select/unselect all fields
def select_all_fields(select):
    for var in field_vars:
        var.set(select)

# Create the main window
root = tk.Tk()
root.title("Jira Wizard")
root.iconbitmap(default=None) # Remove default Tkinter window icon

# Set default window size based on overall elements listed
root.geometry("800x600")

# Add scrollbars where necessary
main_frame = ttk.Frame(root)
main_frame.pack(fill=tk.BOTH, expand=True)
canvas = tk.Canvas(main_frame)
scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
canvas.configure(yscrollcommand=scrollbar.set)
scrollable_frame = ttk.Frame(canvas)
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

def on_resize(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

scrollable_frame.bind("<Configure>", on_resize)

# Enable scrolling with mouse wheel
def on_mouse_wheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")

canvas.bind_all("<MouseWheel>", on_mouse_wheel)

# Section 1: URL, Username, Auth-Token, JQL
section1 = ttk.LabelFrame(scrollable_frame, text="Jira Credentials")
section1.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
ttk.Label(section1, text="URL:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
url_entry = ttk.Entry(section1, width=50)
url_entry.grid(row=0, column=1, padx=5, pady=5)
ttk.Label(section1, text="Username:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
username_entry = ttk.Entry(section1, width=50)
username_entry.grid(row=1, column=1, padx=5, pady=5)
ttk.Label(section1, text="Auth-Token:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
auth_token_entry = ttk.Entry(section1, width=50)
auth_token_entry.grid(row=2, column=1, padx=5, pady=5)
ttk.Label(section1, text="JQL:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
jql_entry = tk.Text(section1, width=50, height=3)
jql_entry.grid(row=3, column=1, padx=5, pady=5)

# Section 2: Field Selection
section2 = ttk.LabelFrame(scrollable_frame, text="Select Fields to Export")
section2.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
fields = ["Key", "Summary", "Status", "Assignee", "Reporter", "Created", "Updated", "Priority", "Issue Type", "Labels", "Project", "Severity", "Applicable Products", "Bug Resolution", "Fixed in Build", "Encountered By", "Team Watch List", "Deferred Products", "How Found", "Found in FW Version", "Reproducibility", "Issue Links", "Resolved"]
field_vars = [tk.BooleanVar(value=True) for _ in fields]
for i, field in enumerate(fields):
    ttk.Checkbutton(section2, text=field, variable=field_vars[i]).grid(row=i//4, column=i%4, padx=5, pady=5, sticky="w")

# Section 3: Select All/Unselect All Buttons
section3 = ttk.LabelFrame(scrollable_frame, text="Select All/Unselect All")
section3.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
ttk.Button(section3, text="Select All", command=lambda: select_all_fields(True)).grid(row=0, column=0, padx=5, pady=5)
ttk.Button(section3, text="Unselect All", command=lambda: select_all_fields(False)).grid(row=0, column=1, padx=5, pady=5)

# Section 4: Open Issues Only Option
section4 = ttk.LabelFrame(scrollable_frame, text="Select to export only Open issues")
section4.grid(row=3, column=0, padx=10, pady=10, sticky="ew")
open_issues_var = tk.BooleanVar(value=False)
ttk.Checkbutton(section4, text="Export Only Open Issues", variable=open_issues_var).grid(row=0, column=0, padx=5, pady=5, sticky="w")

# Section 5: Pivot Tables
section5 = ttk.LabelFrame(scrollable_frame, text="Pivot Tables")
section5.grid(row=4, column=0, padx=10, pady=10, sticky="ew")
severity_var = tk.BooleanVar(value=True)
blockers_var = tk.BooleanVar(value=True)
bug_resolution_var = tk.BooleanVar(value=True)
ttk.Checkbutton(section5, text="Severity", variable=severity_var).grid(row=0, column=0, padx=5, pady=5, sticky="w")
ttk.Checkbutton(section5, text="Blockers", variable=blockers_var).grid(row=0, column=1, padx=5, pady=5, sticky="w")
ttk.Checkbutton(section5, text="Bug Resolution", variable=bug_resolution_var).grid(row=0, column=2, padx=5, pady=5, sticky="w")

# Section 6: Export Button and Progress Bar
section6 = ttk.LabelFrame(scrollable_frame, text="Export into Excel")
section6.grid(row=5, column=0, padx=10, pady=10, sticky="ew")
ttk.Button(section6, text="Export", command=export_issues).grid(row=0, column=0, padx=5, pady=5)
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(section6, variable=progress_var, maximum=100)
progress_bar.grid(row=0, column=1, padx=5, pady=5)

# Section 7: Tool Description
section7 = ttk.LabelFrame(scrollable_frame, text="About Jira Wizard")
section7.grid(row=6, column=0, padx=10, pady=10, sticky="ew")
description = """Ever wished you could magically teleport your Jira issues into an Excel sheet without the browser drama and PingID acrobatics?
✨ 
Well, your wish is granted! Introducing new tool that does just that! Say goodbye to endless logins and hello to instant exports. It's like having a personal Jira wizard at your service!.
 
Dev By: Maninder Singh\nVersion: 1.0.0"""
ttk.Label(section7, text=description, wraplength=400).grid(row=0, column=0, padx=5, pady=5)

# Load credentials on startup
load_credentials()

# Run the main loop
root.mainloop()

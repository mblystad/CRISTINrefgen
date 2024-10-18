import requests
import json
import os
import tkinter as tk
from tkinter import messagebox
from tkinter import PhotoImage
from tkinter.ttk import Combobox
import win32com.client as win32
import datetime

# Function to fetch data for a specific year and filter by year
def fetch_all_data(person_id, year):
    base_url = f"https://api.cristin.no/v2/persons/{person_id}/results"
    results = []
    page = 1

    while True:
        response = requests.get(base_url, params={'page': page, 'per_page': 100})
        response.raise_for_status()  # Raise an error for bad responses
        data = response.json()

        if not data:  # Stop if no more data
            break

        # Convert the year_published field to an integer for proper comparison
        filtered_data = [item for item in data if item.get('year_published') and int(item.get('year_published')) == year]
        results.extend(filtered_data)

        page += 1

    return results

# Function to write data to a JSON file
def write_to_json(data, filename="response_data.json"):
    with open(filename, "w", encoding="utf-8") as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=4)

# Function to read data from a JSON file
def read_from_json(filename="response_data.json"):
    with open(filename, encoding="utf-8") as json_file:
        return json.load(json_file)

# Function to get the title from an item
def get_title(item):
    title_obj = item.get('title', {})
    for lang in title_obj:
        if title_obj[lang]:
            return title_obj[lang]
    return 'No title available'

# Function to format references in APA style and sort by category
def format_references_by_category(data):
    category_references = {}
    global tot_entries
    tot_entries = 0

    for item in data:
        authors = "; ".join([f"{author['first_name']} {author['surname']}" for author in
                             item.get('contributors', {}).get('preview', [])])
        year = item.get('year_published', 'No year available')
        title = get_title(item)
        publication_info = (
                    item.get('journal', {}).get('name') or item.get('event', {}).get('name') or item.get('channel',
                                                                                                         {}).get(
                'title') or 'No publication information available')
        volume_info = item.get('volume', 'N/A')
        pages = item.get('pages', {})
        pages = f"{pages.get('from', 'N/A')}-{pages.get('to', 'N/A')}"
        doi = next((link['url'] for link in item.get('links', []) if link['url_type'] == 'DOI'), 'No DOI available')
        entry_type = item.get('category', {}).get('name', {}).get('en', 'Unknown type')

        reference = f"{authors} ({year}). {title}. {publication_info}, Volume {volume_info}, pp. {pages}. {doi}. {entry_type}"

        # Group by category
        if entry_type not in category_references:
            category_references[entry_type] = []
        category_references[entry_type].append(reference)

        tot_entries += 1

    return category_references

# Function to write references to a DOC file using pywin32, sorted by category
def write_to_doc_by_category(category_references, filename):
    word = win32.Dispatch('Word.Application')
    doc = word.Documents.Add()
    word.Visible = False

    # Write references grouped by category with headings
    for category, references in category_references.items():
        doc.Range().InsertAfter(category + '\n')
        doc.Range().InsertParagraphAfter()

        for reference in references:
            doc.Range().InsertAfter(reference + '\n')
            doc.Range().InsertParagraphAfter()

        # Add extra spacing between categories
        doc.Range().InsertParagraphAfter()

    doc.SaveAs(os.path.abspath(filename), FileFormat=0)  # Save as .doc format
    doc.Close()
    word.Quit()

    print(f"Document saved as {os.path.abspath(filename)}")  # Debug statement

# Function to generate the file and optionally send it via email
def generate_file():
    person_id = entry_cristin.get()
    to_email = entry_email.get()
    selected_year = int(combo.get())

    if not person_id.isdigit():
        messagebox.showerror("Error", "Please enter a valid Cristin ID number.")
        return

    person_id = int(person_id)
    data = fetch_all_data(person_id, selected_year)
    write_to_json(data)

    # Read and process data
    data = read_from_json()
    category_references = format_references_by_category(data)
    doc_filename = f"{person_id}_formatted_references_{selected_year}.doc"

    # Write references to Word document, sorted by category
    write_to_doc_by_category(category_references, doc_filename)

    if to_email:
        send_email(to_email, doc_filename)
        messagebox.showinfo("Success",
                            f"Formatted APA references for {selected_year} have been created and sent to {to_email}.\nFile location: {os.path.abspath(doc_filename)}")
    else:
        messagebox.showinfo("Success", f"Formatted APA references for {selected_year} have been created.\nFile location: {os.path.abspath(doc_filename)}")

# Tkinter GUI setup
root = tk.Tk()
root.title("Cristin ID Reference Formatter")

# Load logo image if available
logo_path = "logo.png"
if os.path.exists(logo_path):
    logo = PhotoImage(file=logo_path)
    tk.Label(root, image=logo).pack(pady=10)
else:
    tk.Label(root, text="Logo not found").pack(pady=10)

# Entry for Cristin ID
tk.Label(root, text="Enter Cristin ID number:").pack(pady=10)
entry_cristin = tk.Entry(root)
entry_cristin.insert(0, "674004")
entry_cristin.pack(pady=5)

# Entry for email address (optional)
tk.Label(root, text="Enter your email address (optional):").pack(pady=10)
entry_email = tk.Entry(root)
entry_email.pack(pady=5)

# Dropdown for selecting the year
year = datetime.date.today().year
tk.Label(root, text="Select year:").pack(pady=5)
combo = Combobox(root, state="readonly", values=[year, year-1, year-2, year-3, year-4, year-5])
combo.current(0)  # Set the current year as default
combo.pack(pady=5)

# Button to generate and send references
tk.Button(root, text="Generate and Send References", command=generate_file).pack(pady=20)

# Run the Tkinter main loop
root.mainloop()

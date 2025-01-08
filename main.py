import requests
import json
import os
import tkinter as tk
from tkinter import messagebox
from tkinter import PhotoImage
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import win32com.client as win32
import matplotlib.pyplot as plt


def fetch_all_data(person_id):
    base_url = f"https://api.cristin.no/v2/persons/{person_id}/results"
    results = []
    page = 1

    while True:
        response = requests.get(base_url, params={'page': page, 'per_page': 100})
        response.raise_for_status()  # Raise an error for bad responses
        data = response.json()

        if not data:  # Stop if no more data
            break

        results.extend(data)
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


# Function to format references in APA style and sort by year
def format_references(data):
    formatted_references = []
    global tot_entries
    tot_entries = 0
    category_counts = {}
    for item in data:
        authors = "; ".join([f"{author['surname']}, {author['first_name']}" for author in
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

        # Count categories
        if entry_type not in category_counts:
            category_counts[entry_type] = 0
        category_counts[entry_type] += 1

        tot_entries += 1
        reference = f"{authors} ({year}). {title}. {publication_info}, Volume {volume_info}, pp. {pages}. {doi}. {entry_type}"
        formatted_references.append(reference)

    formatted_references.sort(key=lambda x: x.split('(')[1].split(')')[0], reverse=True)  # Sort by year, most recent first
    return formatted_references, tot_entries, category_counts


# Function to write categories to a text file
def write_categories_to_txt(category_counts, filename="categories.txt"):
    with open(filename, "w", encoding="utf-8") as file:
        for category, count in category_counts.items():
            file.write(f"{category} {count}\n")


# Function to read category counts from a text file
def read_category_counts(filename="categories.txt"):
    categories = []
    counts = []

    with open(filename, "r", encoding="utf-8") as file:
        for line in file:
            category, count = line.rsplit(maxsplit=1)
            categories.append(f"{category} ({count})")  # Include count in category label
            counts.append(int(count))

    return categories, counts


# Function to create and save a pie chart from category counts
def create_pie_chart(categories, counts, output_filename="category_pie_chart.png"):
    plt.figure(figsize=(15, 10))  # Increased figure size for better visibility

    wedges, texts, autotexts = plt.pie(counts, labels=categories, autopct='%1.1f%%', startangle=90,
                                       colors=plt.cm.Paired(range(len(categories))), textprops={'fontsize': 12},
                                       wedgeprops=dict(width=0.3, edgecolor='w'), pctdistance=0.85, labeldistance=1.05)

    # Make sure labels do not overlap
    for text in texts:
        text.set_fontsize(10)
    for autotext in autotexts:
        autotext.set_fontsize(10)
        autotext.set_color('white')

    plt.title('Category Distribution', fontsize=16)
    plt.savefig(output_filename, bbox_inches='tight')
    plt.close()
    print(f"Pie chart saved as '{output_filename}'")


# Function to write references to a DOC file using pywin32
def write_to_doc(references, filename, category_counts):
    word = win32.Dispatch('Word.Application')
    doc = word.Documents.Add()
    word.Visible = False

    # Insert pie chart at the beginning of the document
    pie_chart_path = os.path.abspath("category_pie_chart.png")
    if os.path.exists(pie_chart_path):
        doc.Range(0, 0).InlineShapes.AddPicture(pie_chart_path)

    # Add category counts to the beginning of the document
    doc.Range().InsertAfter(f'Current Cristin Entry count: {tot_entries}\n')
    for category, count in category_counts.items():
        doc.Range().InsertAfter(f'{category}: {count}\n')
    doc.Range().InsertParagraphAfter()

    # Add formatted references
    for reference in references:
        doc.Range().InsertAfter(reference + '\n')
        doc.Range().InsertParagraphAfter()

    doc.SaveAs(os.path.abspath(filename), FileFormat=0)  # Save as .doc format
    doc.Close()
    word.Quit()

    print(f"Document saved as {os.path.abspath(filename)}")  # Debug statement


# Function to write references to a TXT file
def write_references_to_txt(references, filename="formatted_references.txt"):
    with open(filename, "w", encoding="utf-8") as file:
        for reference in references:
            file.write(reference + "\n")

    print(f"References saved as {os.path.abspath(filename)}")


# Function to send an email with the file attached
def send_email(to_email, file_path):
    from_email = "your_email@example.com"
    from_password = "your_password"

    subject = "Formatted APA References"
    body = f"Please find the formatted APA references attached.\n\nFile location: {file_path}"

    msg = MIMEMultipart()
    msg["From"] = from_email
    msg["To"] = to_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    with open(file_path, "rb") as file:
        attachment = MIMEText(file.read(), "plain")
        attachment.add_header("Content-Disposition", "attachment", filename=os.path.basename(file_path))
        msg.attach(attachment)

    with smtplib.SMTP("smtp.example.com", 587) as server:
        server.starttls()
        server.login(from_email, from_password)
        server.send_message(msg)


# Function to generate the file and optionally send it via email
def generate_file():
    person_id = entry_cristin.get()
    to_email = entry_email.get()

    if not person_id.isdigit():
        messagebox.showerror("Error", "Please enter a valid Cristin ID number.")
        return

    person_id = int(person_id)
    data = fetch_all_data(person_id)
    write_to_json(data)

    # Read and process data
    data = read_from_json()
    doc_filename = f"{person_id}_formatted_references.doc"
    references, tot_entries, category_counts = format_references(data)

    # Store category counts in a separate variable
    global stored_category_counts
    stored_category_counts = category_counts

    # Write categories to text file
    write_categories_to_txt(category_counts)

    # Generate the text file and pie chart before creating the Word document
    generate_txt_and_chart()

    # Write references to Word document
    write_to_doc(references, doc_filename, category_counts)

    # Write references to TXT file
    write_references_to_txt(references)

    if to_email:
        send_email(to_email, doc_filename)
        messagebox.showinfo("Success",
                            f"Formatted APA references have been created and sent to {to_email}.\nFile location: {os.path.abspath(doc_filename)}")
    else:
        messagebox.showinfo("Success", f"Formatted APA references have been created.\nFile location: {os.path.abspath(doc_filename)}")


# Generate the text file and pie chart
def generate_txt_and_chart():
    categories, counts = read_category_counts("categories.txt")
    create_pie_chart(categories, counts)


tot_entries = 0
stored_category_counts = {}
# Tkinter GUI setup
root = tk.Tk()
root.title("Cristin ID Reference Formatter")



# Entry for Cristin ID
tk.Label(root, text="Enter Cristin ID number:").pack(pady=10)
entry_cristin = tk.Entry(root)
entry_cristin.insert(0, "")
entry_cristin.pack(pady=5)

# Entry for email address (optional)
tk.Label(root, text="Enter your email address (optional):").pack(pady=10)
entry_email = tk.Entry(root)
entry_email.pack(pady=5)

# Button to generate and send references
tk.Button(root, text="Generate and Send References", command=generate_file).pack(pady=20)

# Add a button to generate the text file and pie chart
tk.Button(root, text="Generate TXT and Chart", command=generate_txt_and_chart).pack(pady=20)

# Run the Tkinter main loop
root.mainloop()

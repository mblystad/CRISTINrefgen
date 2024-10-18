**Cristin ID Reference Formatter**
This Python application retrieves publication data for a specific person from the Cristin API and formats the references into APA style, grouped by category. It allows you to generate a Word document with the formatted references and optionally send the file via email. The program uses a GUI built with Tkinter for easy input and interaction.

**Features**
Retrieve Publication Data: Fetch publication results from Cristin's API based on a specific person’s Cristin ID and a chosen year.

Format References: Convert publication data into APA-style references, grouped by category (journal articles, conference papers, etc.).

Save to Word Document: Automatically generate a .doc file containing the formatted references.

Email the File: Optionally send the generated Word document via email.

GUI Interface: Simple user interface built with Tkinter for entering the Cristin ID, selecting the year, and optionally providing an email address.

**Installation
Requirements**
Python 3.x
Dependencies:
requests
json
tkinter
win32com (for Word document generation on Windows)
datetime
Install the required Python packages via pip:

bash
Copy code
pip install requests pywin32
Running the Application
Clone this repository:

bash
Copy code
git clone https://github.com/your-username/cristin-id-formatter.git
cd cristin-id-formatter
Run the script:

bash
Copy code
python cristin_formatter.py

The Tkinter GUI will open, allowing you to:
Enter a Cristin ID.
Choose a year.
Optionally enter an email address to send the generated file.

**Usage**
User Interface
Cristin ID: Enter the ID of the person whose publication data you want to retrieve.
Year: Choose the publication year for which you want to format references.
Email Address: (Optional) Enter an email to send the formatted document.
Output
A .doc file will be generated containing the APA-formatted references, grouped by publication category.
If an email address is provided, the document will be sent to that address.

**Example**
Enter Cristin ID 123456 and select the year 2023.
Click on Generate and Send References.
A Word document named 123456_formatted_references_2023.doc will be created and optionally sent via email.

**Limitations**
This program works only on Windows systems, as it relies on win32com for generating Word documents.
Requires Microsoft Word to be installed on the machine.
The application fetches publication data from the Cristin API. Changes to the API structure or availability may affect the program’s functionality.

**Contributing**
Feel free to fork this repository and submit pull requests. Contributions are welcome!

# Cristin ID Reference Formatter
The current release is in early development, and I will try to work on but it is not my top priority. Idea and first few lines of code by mblystad, most work provided by LLMs (ChatGPT / Copilot).

This repository contains a Python-based GUI application that:

1. Fetches publication data for a given Cristin person ID.  
2. Formats the fetched data into APA references.  
3. Generates various output files:  
   - A JSON file containing the raw data.  
   - A text file (`categories.txt`) listing the count of each publication category.  
   - A pie chart (`category_pie_chart.png`) visualizing the category distribution.  
   - A Word document (`.doc`) with the formatted references and the pie chart.  
   - A text file (`formatted_references.txt`) with the references.  
4. Optionally sends the Word document via email.

---

## Features

- **Automated data fetching**: Retrieves all publication data from [Cristin](https://api.cristin.no) for the specified person ID.  
- **APA reference formatting**: Converts publication details into APA-style references.  
- **Category breakdown**: Counts and visualizes how many publications fall under each category.  
- **Export to multiple formats**:  
  - **JSON** for raw data  
  - **TXT** for category counts and references  
  - **DOC** for a Word-compatible document with embedded pie chart  
- **Email sending**: Optionally attach and send the `.doc` file to a specified email address.  
- **Graphical User Interface (GUI)**: Built with Tkinter, making it easy to enter a Cristin ID and email address.

---

## Requirements

- **Python 3.7+**  
- **Pip** for package installation  
- **Operating System**: Windows (required for Word automation via `pywin32`)  

### Python Libraries

1. **requests** – for making HTTP requests to the Cristin API  
2. **json** – for handling JSON data  
3. **tkinter** – for creating the GUI  
4. **smtplib**, **email** – for sending emails  
5. **pywin32 (win32com)** – for interacting with Microsoft Word  
6. **matplotlib** – for generating the pie chart  
7. **os** – for file and path handling  

You can install the required packages (except for built-in modules) by running:

```bash
pip install requests matplotlib pywin32
```

Note: `tkinter`, `json`, `smtplib`, and `email` are part of the standard library and do not need separate installation.

---

## Installation

1. **Clone** this repository or **download** the files.
2. Ensure you have all [Requirements](#requirements) installed.
3. Place the script (e.g., `main.py`) in the same folder.

---

## Usage

1. **Open** a terminal or command prompt in the script's directory.
2. **Run** the script:

   ```bash
   python main.py
   ```

3. The GUI will appear with the following fields:
   - **Cristin ID number**: Enter a valid numeric Cristin ID (default is set to `674004` for testing).  
   - **Email address (optional)**: If provided, the script will send the generated Word document as an attachment to this email address.

4. **Click** **"Generate and Send References"** to:  
   - Fetch all publication data from Cristin.  
   - Generate and save the output files:
     - `response_data.json` (raw data)  
     - `categories.txt` (category counts)  
     - `category_pie_chart.png` (pie chart)  
     - `[CristinID]_formatted_references.doc` (Word doc with references + pie chart)  
     - `formatted_references.txt` (list of references)  
   - If an email was provided, attach and send the `.doc` file.

5. Alternatively, you can **only** generate the `.txt` file and pie chart by clicking **"Generate TXT and Chart"**.

> **Note**: If you do not have Microsoft Office installed or are on a non-Windows system, the Word document generation via `pywin32` will not work.

---

## Files Explained

- **`main.py`**  
  The main Python script that launches a Tkinter GUI and executes all functionality (fetching data, formatting references, generating outputs, sending email, etc.).

- **`response_data.json`** (generated)  
  The raw JSON data fetched from Cristin.

- **`categories.txt`** (generated)  
  A text file listing each category along with how many publications belong to it.

- **`category_pie_chart.png`** (generated)  
  A pie chart visualizing the category distribution of the publications.

- **`[CristinID]_formatted_references.doc`** (generated)  
  A Word `.doc` file containing the APA-style references plus the embedded pie chart and category counts.

- **`formatted_references.txt`** (generated)  
  A text file containing the references in APA format.

---

## Customization

- **Email sender credentials**:  
  In the `send_email()` function, replace:
  ```python
  from_email = "your_email@example.com"
  from_password = "your_password"
  ```
  with your actual email and password.  
  Also, adjust the SMTP server (`smtp.example.com`) and port as needed for your provider.

- **Default Cristin ID**:  
  Change the `entry_cristin.insert(0, "")` line to your preferred default Cristin ID.

---

## Troubleshooting

- **Invalid Cristin ID**:  
  Ensure you’ve entered numeric values. Non-numeric input will trigger an error message.
  
- **Word Document Generation Failures**:  
  - Check that Microsoft Word is installed on your Windows system.  
  - Verify that `pywin32` is installed (`pip install pywin32`).  
  - If you see a COM error, you may need to run the `pywin32_postinstall.py` script typically located in your Python Scripts folder.

- **SMTP / Email Issues**:  
  - Verify SMTP settings (server, port).  
  - Make sure you’ve allowed “less secure apps” or set up app-specific passwords if using Gmail or similar providers.  

---

## Contributing

1. **Fork** the repository  
2. **Create** your feature branch: `git checkout -b feature/new-feature`  
3. **Commit** your changes: `git commit -am 'Add some feature'`  
4. **Push** to the branch: `git push origin feature/new-feature`  
5. **Open** a Pull Request

---

## License

This project is provided under the [MIT License](LICENSE). You are free to use, modify, and distribute this software, provided you include the license notice in your version. 

---

**Thank you for using the Cristin ID Reference Formatter!** If you encounter any issues or have suggestions, please open an issue or submit a pull request.

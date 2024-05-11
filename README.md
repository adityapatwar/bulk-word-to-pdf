# bulk_convert_word_to_pdf


 Requirements Installation and Local Environment Setup

1. Install Python:
   - Ensure Python 3.8 or higher is installed on your system. You can download it from [Python's official website](https://python.org).

2. Install LibreOffice:
   - Download and install LibreOffice from [LibreOffice's download page](https://www.libreoffice.org/download/download/).
   - Ensure it is added to your system's PATH to be accessed from the command line.

3. Install Pandoc:
   - Download and install Pandoc from [Pandoc's official site](https://pandoc.org/installing.html).
   - Make sure it is added to your system's PATH.

4. Install Python Dependencies:
   - Create a `requirements.txt` file in your project directory and add the following line:
     ```
     PyPDF2==2.10.0
     ```
   - Run the following command in your terminal to install the required Python libraries:
     ```bash
     pip install -r requirements.txt
     ```
     or install manual

5. Prepare the Mapping File:
   - Create a `mappings.txt` file in the directory where your script resides. This file should contain mappings in the format `Full Phrase,Abbreviation`, one per line. Example entries might look like:
     ```
     Novel,No
     Buku Novel,Book
     ```
   - Ensure this file is correctly referenced in your script (check the `MAPPING_FILE` path).

6. Running the Script:
   - Place your DOCX files in a source directory specified by the `SOURCE` path in your script.
   - Ensure the `OUTPUT` directory exists or is creatable by your script for storing the converted PDFs.
   - Run the script using the command:
     ```bash
     python main.py
     ```

7. Expected Output:
   - The encrypted PDFs will be available in the output directory specified in the script.
   - Passwords for each PDF will be logged in the specified password log file.

 Local Environment Execution

- Open your terminal.
- Navigate to the directory containing your script.
- Execute the script using Python:
  ```bash
  python main.py
  ```
  Make sure that all paths (`SOURCE`, `OUTPUT`, `PASSWORD_LOG_PATH`, `MAPPING_FILE`) are set correctly in your script to reflect your local environment and directory structure.

 Additional Tips

- Debugging: If you encounter issues, check the console outputs for errors related to file paths, permissions, or external application calls.
- Version Control: If you are using a version control system like git, remember to add the `requirements.txt` and `mappings.txt` to your repository but exclude sensitive files like the output PDFs and logs.
- Scheduled Tasks: If you need to run this script periodically, consider setting it up as a cron job or a scheduled task depending on your operating system.
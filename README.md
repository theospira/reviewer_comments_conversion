# Reviewer Comments Conversion

## Description

This repository contains scripts to convert reviewer comments from a text file to an Excel spreadsheet and subsequently from the spreadsheet to a Word document. These scripts are designed to streamline the process of organizing, responding to, and formatting reviewer comments for academic submissions.

## Features

- Converts reviewer comments from a text file into an Excel spreadsheet with organized columns.
- Categorizes comments into "Reviewer 1", "Reviewer 2", "Major comment", and "Minor comment".
- Includes a "done?" column to track the status of each comment.
- Marks comments requiring additional input from lead authors or co-authors.
- Converts the Excel spreadsheet into a Word document formatted for submission and proofing.

## Usage

### Text File to Excel Spreadsheet

This script converts reviewer comments from a text file into an Excel spreadsheet with the following columns:

1. **Category**: Specifies the reviewer (Reviewer 1, Reviewer 2) and the type of comment (Major or Minor).
2. **Comment**: Contains the actual comment from the reviewer.
3. **Done?**: Indicates whether the comment has been addressed, needs returning to by the lead author, or requires input from co-authors.
4. **Response**: Contains the authors' response to the comment.

The script is based on the unique format of the received reviews and can be adapted as needed.

### Excel Spreadsheet to Word Document

This script converts the Excel spreadsheet into a Word document formatted for submission. The Word document is structured as follows:

- Comments from Reviewer 1 followed by the authors' responses.
- Comments from Reviewer 2 followed by the authors' responses.
- Major and minor comments are distinguished, and responses are formatted in blue.
- Additional notes are included for comments that require special attention from specific team members.

## Requirements

- Python 3.x
- Libraries:
  - `openpyxl`
  - `python-docx`
  
You can install the necessary libraries using pip:

```bash
pip install openpyxl python-docx
```

## Acknowledgements

    Coding assistance provided by ChatGPT from OpenAI.

## Contributing

If you would like to contribute to this project, please submit a pull request or open an issue for any bugs or feature requests.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

# Personalised-Document-Generator

A Streamlit-based web application that generates personalized Word and PDF documents from employee data. The application supports both Excel file uploads and manual data entry.

## Features

- Generate personalized Word (.docx) and PDF documents
- Two input methods:
  - Excel file upload (supports multiple files)
  - Manual data entry form
- Real-time data preview
- Automatic output directory creation
- Clean and intuitive user interface

## Requirements

- Python >= 3.11
- docx >= 0.2.4
- pandas >= 2.2.3
- reportlab >= 4.3.0
- streamlit

## Installation

### Clone the repository:

```sh
git clone 
cd personalized-document-generator
```

### Install dependencies:

```sh
pip install -r requirements.txt
```

## Usage

### Run the Streamlit application:

```sh
streamlit run main.py
```

### Select your preferred input method:

- **Excel Upload:** Upload Excel files containing employee data
- **Manual Entry:** Enter employee information manually through the form

### Required data fields:

- Name
- Email
- Company Name
- Position
- Joining Date

Click **"Generate Documents"** to create personalized documents in both Word and PDF formats.

## Output

Generated documents will be saved in the `output` directory with the following naming convention:

- **Word:** `{Name}_{Company_Name}.docx`
- **PDF:** `{Name}_{Company_Name}.pdf`

## Project Structure

```
.
├── main.py           # Main application file
├── output/           # Generated documents directory
├── requirements.txt  # Project dependencies
└── README.md         # Project documentation
```



## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature-branch`)
3. Commit your changes (`git commit -m "Add new feature"`)
4. Push to the branch (`git push origin feature-branch`)
5. Create a new Pull Request

## Support

For support or questions, please open an issue in the project repository.

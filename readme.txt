# Project Title: PowerPoint Text Extraction and Manipulation

## Description

This project provides a Python script to extract text from PowerPoint slides and save them into an Excel file. It also includes a script to update the text in PowerPoint slides based on the data in an Excel file.

## Installation

1. Clone the repository:
```
git clone https://github.com/your-repo-url
```

2. Install the required packages:
```
pip install -r requirements.txt
```

## Usage

### Extract text from PowerPoint to Excel

1. Place your PowerPoint file in the project directory and rename it to `CONTOSO.pptx`.

2. Run the script:
```
python extract_text.py
```

3. The extracted text will be saved in `text_boxes.xlsx` in the project directory.

### Update text in PowerPoint from Excel

1. Place your PowerPoint file in the project directory and rename it to `CONTOSO.pptx`.

2. Place your Excel file in the project directory and rename it to `presentation_data.xlsx`. The Excel file should have the following columns: `Slide Number`, `Object Name`, `Object Type`, `Text`.

3. Run the script:
```
python update_text.py
```

4. The PowerPoint file will be updated with the text from the Excel file.

## Troubleshooting

If you encounter an `IllegalCharacterError`, it means that there are illegal characters in the text that cannot be used in Excel worksheets. The script already includes a function to remove these illegal characters from the text before adding it to the worksheet.

If you encounter an error due to the 'text' variable being None, it means that you are trying to assign a None value to the text of a shape. The script already includes a check to ensure that 'text' is not None before assigning it to the text of a shape.

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License

[MIT](https://choosealicense.com/licenses/mit/)
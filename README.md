# HourGlass

Streamlit application for processing mention spreadsheets and exporting filtered Excel and CSV files.

## Run locally

1. Create and activate a virtual environment.
2. Install dependencies:

   ```bash
   python -m pip install -r requirements.txt
   ```

3. Start the app:

   ```bash
   streamlit run streamlit_app.py
   ```

Upload an `.xlsx` workbook with a `Mentions` sheet, enter the source file name, and select **Process File**.

## Project files

- `streamlit_app.py`: Streamlit user interface and processing rules.
- `Pattern_fileRaw.ipynb`: Original notebook workflow.
- `requirements.txt`: Python dependencies captured with `pip freeze`.

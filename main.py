import streamlit as st
import pandas as pd
import io
import zipfile
import openpyxl
from openpyxl.styles import Font
from copy import copy
import numpy as np


# set page configurations
st.set_page_config(
    page_title="Cleaning App",
    layout="centered",
    initial_sidebar_state="collapsed",
    page_icon=":material/mop:"
)

# the custom CSS lives here:
hide_default_format = """
    <style>
        .reportview-container .main footer {visibility: hidden;}    
        #MainMenu, header, footer {visibility: hidden;}
        div.stActionButton{visibility: hidden;}
        [class="stAppDeployButton"] {
            display: none;
        }
        [data-testid="collapsedControl"] {
            display: none
        }
        [data-testid="stFileUploaderDropzone"] {
            background-color: #9cdcf3;
            border-radius: 15px;
            border: 1px solid #1F2041;
            padding: 15px; /* Optional: Add some padding */
        }
        div[data-testid="stFileUploaderDropzoneInstructions"]>div>span {
            visibility: hidden;
        }
        div[data-testid="stFileUploaderDropzoneInstructions"]>div>span::before {
            content: "Drag & drop Survey Monkey results for cleaning.";
            visibility: visible;
        }
        div[data-testid="stFileUploaderDropzoneInstructions"]>div>small {
            visibility: hidden;
        }
        div[data-testid="stFileUploaderDropzoneInstructions"]>div>small::before {
            content: "You can upload more than 1 file at a time!";
            visibility: visible;
        }
        .stDownloadButton, div.stButton {text-align:center}
        [data-testid="stBaseButton-secondary"] {
            background-color: #9cdcf3;
            border-radius: 15px;
            border: 1px solid #1F2041;
        }
        [data-testid="stFileUploaderFile"] {
            color: #fefefe;
        }
        div[data-testid="stFileUploaderFile"] div>small {
            color: #fefefe;
        }
        [data-testid="stFileUploaderPagination"] >small {
            color: #fefefe;
        }
        [data-testid="stBaseButton-minimal"] {
            color: #fefefe;
        }
    </style>
"""

# inject the CSS
st.markdown(hide_default_format, unsafe_allow_html=True)

# heading text
# fefefe (white)
# 9cdcf3 (blue)
# 1b1b1b (black)
st.markdown(f'''
    <p style="font-size: 40px; font-weight: 900; text-align: center; margin-top: 10px; margin-bottom: 80px; color: #9cdcf3;">
        11 Ten Leadership Cleaning App
    </p>
''', unsafe_allow_html=True)


# Define helper function to extract the question number from the string for sorting
def extract_question_number(question):
    try:
        return int(question.split()[0][1:])  # Extract the number following 'Q'
    except ValueError:
        return None  # If no number is found, return None


def main():
    # declare variable for uploading files
    uploaded_files = st.file_uploader(
        label="Choose completed reporting template",
        label_visibility='collapsed',
        accept_multiple_files=True,
        type=["xlsx"],
        help="Upload Survey Monkey template(s) here."
    )

    if uploaded_files:

        # # show photo of Belichick
        # st.image('assets/bill_2.png')

        # Prepare list to hold file data
        files_to_download = []

        # Iterate through uploaded files
        for uploaded_file in uploaded_files:
            # Read the file into memory
            file_data = io.BytesIO(uploaded_file.read())

            # Load the workbook using openpyxl
            wb = openpyxl.load_workbook(file_data)
            ws = wb.active  # Assuming the data is on the active sheet

            # Iterate over all merged cells and unmerge them
            for merged_cell_range in list(ws.merged_cells.ranges):
                ws.unmerge_cells(str(merged_cell_range))

            # Add new columns & formatting
            ws["D23"] = "Question Order"
            ws["D23"].font = copy(ws["C23"].font)
            ws["D23"].fill = copy(ws["C23"].fill)
            ws["D23"].alignment = copy(ws["C23"].alignment)

            ws["E23"] = "1st-Order Category"
            ws["E23"].font = copy(ws["C23"].font)
            ws["E23"].fill = copy(ws["C23"].fill)
            ws["E23"].alignment = copy(ws["C23"].alignment)

            ws["F23"] = "2nd-Order Category"
            ws["F23"].font = copy(ws["C23"].font)
            ws["F23"].fill = copy(ws["C23"].fill)
            ws["F23"].alignment = copy(ws["C23"].alignment)

            # Extract the data starting from A23 down until the first blank cell
            row = 24
            data = []
            while ws[f"A{row}"].value is not None:
                # Append row data into the list
                data.append(
                    [ws[f"A{row}"].value, ws[f"B{row}"].value, ws[f"C{row}"].value])
                row += 1

            # Create a DataFrame from the extracted data
            df = pd.DataFrame(
                data, columns=["Questions", "Difficulty", "Average Score"])

            # Convert the 'Average Score' column to integer
            df["Average Score"] = df["Average Score"].str.replace(
                '%', '').astype(int)
            df = df.rename(columns={"Average Score": "Avg. Score (%)"})

            # Add the "Question Order" column based on the extracted question
            df["Question Order"] = df["Questions"].apply(
                extract_question_number).astype(int)

            # Sort the DataFrame by "Question Order"
            df = df.sort_values(by="Question Order")

            # Overwrite the "Question Order" column with sequential numbers starting from 1
            df["Question Order"] = range(1, len(df) + 1)

            # 1st-, 2nd-, and 3rd-Order Categories
            category_1 = ['THRIVE', 'Just Leader']
            category_2 = [
                'Trust', 'Health', 'Relationships', 'Impact', 'Value', 'Engagement', 'See the Whole Playing Field', 'Build Cultural Competency', 'Give Power Away', 'Take Bold, Courageous Action']

            # this level will only pertain to the LEADER and TEAM templates
            category_3 = [
                'Leader', 'Team', 'Leader', 'Team', 'Leader', 'Team', 'Leader', 'Team', 'Leader', 'Team', 'Leader', 'Team', 'Leader', 'Team', 'Leader', 'Team', 'Leader', 'Team', 'Leader', 'Team']

            template = ''

            # fill out templates
            if df.shape[0] == 38:  # Review template

                template = 'Review'

                # fill in the "Category" column
                repetitions_1 = [30, 8]
                category_list_1 = []
                for category, rep in zip(category_1, repetitions_1):
                    category_list_1.extend([category] * rep)
                df['1st-Order Category'] = category_list_1
                first0_summary = df.groupby(
                    '1st-Order Category')['Avg. Score (%)'].mean().reset_index()

                # Start from A22 and find the last filled cell in column A
                current_row = 22
                while ws[f"A{current_row}"].value is not None:  # Check if cell is filled
                    current_row += 1  # Move down

                # Move one row below the last filled cell
                start_row = current_row + 1

                # Insert 12 blank rows after the last row written
                ws.insert_rows(current_row, 16)

                # reset the row count to copy / paste values
                current_row = 22
                while ws[f"A{current_row}"].value is not None:  # Check if cell is filled
                    current_row += 1  # Move down
                start_row = current_row + 1

                # Write the DataFrame headers to the sheet
                headers = list(first0_summary.columns)  # Get column headers

                # Write headers in the first row
                for col_idx, header in enumerate(headers, start=1):
                    ws.cell(row=start_row, column=col_idx, value=header)

                # Write the DataFrame contents to the sheet
                for row_idx, row_data in first0_summary.iterrows():
                    for col_idx, value in enumerate(row_data, start=1):
                        ws.cell(row=start_row + row_idx + 1,
                                column=col_idx, value=value)

                start_row = start_row + 4

                repetitions_2 = [5, 5, 5, 5, 5, 5, 2, 2, 2, 2]
                category_list_2 = []
                for category, rep in zip(category_2, repetitions_2):
                    category_list_2.extend([category] * rep)
                df['2nd-Order Category'] = category_list_2
                second0_summary = df.groupby(
                    ['1st-Order Category', '2nd-Order Category'])['Avg. Score (%)'].mean().reset_index().drop(columns='1st-Order Category')

                # use a custom sort for the summary table
                second0_summary['2nd-Order Category'] = pd.Categorical(
                    second0_summary['2nd-Order Category'], [
                        "Trust", "Health", "Relationships", "Impact", "Value", "Engagement", "See the Whole Playing Field", "Build Cultural Competency", "Give Power Away", "Take Bold, Courageous Action"
                    ])
                second0_summary = second0_summary.sort_values(
                    "2nd-Order Category", ascending=False)

                # Write headers in the first row of the new section
                for col_idx, header in enumerate(second0_summary.columns, start=1):
                    ws.cell(row=start_row, column=col_idx, value=header)

                # Write the DataFrame contents row by row
                for row_idx, row_data in second0_summary.iterrows():
                    for col_idx, value in enumerate(row_data, start=1):
                        ws.cell(row=start_row + row_idx + 1,
                                column=col_idx, value=value)

                # Define the range of cells
                start_row = 24
                end_row = 77
                start_col = 1  # Column A (1-indexed)
                end_col = 7  # Column C (1-indexed)

                # Define the font style
                custom_font = Font(name="Arial", size=11)

                # Apply the font style to each cell in the range
                for row in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col):
                    for cell in row:
                        cell.font = custom_font

            elif df.shape[0] == 47:  # No leader template

                template = 'No leader'
                repetitions_1 = [35, 12]
                category_list_1 = []
                for category, rep in zip(category_1, repetitions_1):
                    category_list_1.extend([category] * rep)
                df['1st-Order Category'] = category_list_1
                first0_summary = df.groupby(
                    '1st-Order Category')['Avg. Score (%)'].mean().reset_index()

                # Start from A22 and find the last filled cell in column A
                current_row = 22
                while ws[f"A{current_row}"].value is not None:  # Check if cell is filled
                    current_row += 1  # Move down

                # Move one row below the last filled cell
                start_row = current_row + 1

                # Insert 12 blank rows after the last row written
                ws.insert_rows(current_row, 16)

                # reset the row count to copy / paste values
                current_row = 22
                while ws[f"A{current_row}"].value is not None:  # Check if cell is filled
                    current_row += 1  # Move down
                start_row = current_row + 1

                # Write the DataFrame headers to the sheet
                headers = list(first0_summary.columns)  # Get column headers

                # Write headers in the first row
                for col_idx, header in enumerate(headers, start=1):
                    ws.cell(row=start_row, column=col_idx, value=header)

                # Write the DataFrame contents to the sheet
                for row_idx, row_data in first0_summary.iterrows():
                    for col_idx, value in enumerate(row_data, start=1):
                        ws.cell(row=start_row + row_idx + 1,
                                column=col_idx, value=value)

                start_row = start_row + 4

                repetitions_2 = [5, 10, 5, 5, 5, 5, 3, 3, 3, 3]
                category_list_2 = []
                for category, rep in zip(category_2, repetitions_2):
                    category_list_2.extend([category] * rep)
                df['2nd-Order Category'] = category_list_2
                second0_summary = df.groupby(
                    ['1st-Order Category', '2nd-Order Category'])['Avg. Score (%)'].mean().reset_index().drop(columns='1st-Order Category')

                # use a custom sort for the summary table
                second0_summary['2nd-Order Category'] = pd.Categorical(
                    second0_summary['2nd-Order Category'], [
                        "Trust", "Health", "Relationships", "Impact", "Value", "Engagement", "See the Whole Playing Field", "Build Cultural Competency", "Give Power Away", "Take Bold, Courageous Action"
                    ])
                second0_summary = second0_summary.sort_values(
                    "2nd-Order Category", ascending=False)

                # Write headers in the first row of the new section
                for col_idx, header in enumerate(second0_summary.columns, start=1):
                    ws.cell(row=start_row, column=col_idx, value=header)

                # Write the DataFrame contents row by row
                for row_idx, row_data in second0_summary.iterrows():
                    for col_idx, value in enumerate(row_data, start=1):
                        ws.cell(row=start_row + row_idx + 1,
                                column=col_idx, value=value)

                # Define the range of cells
                start_row = 24
                end_row = 86
                start_col = 1  # Column A (1-indexed)
                end_col = 7  # Column C (1-indexed)

                # Define the font style
                custom_font = Font(name="Arial", size=11)

                # Apply the font style to each cell in the range
                for row in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col):
                    for cell in row:
                        cell.font = custom_font

            elif str(df["Questions"].iloc[0]).startswith("Q7"):  # Leader template

                template = 'Leader'
                # First level breakdown
                repetitions_1 = [60, 20]
                category_list_1 = []
                for category, rep in zip(category_1, repetitions_1):
                    category_list_1.extend([category] * rep)
                df['1st-Order Category'] = category_list_1
                first0_summary = df.groupby(
                    '1st-Order Category')['Avg. Score (%)'].mean().reset_index()

                # Start from A22 and find the last filled cell in column A
                current_row = 22
                while ws[f"A{current_row}"].value is not None:  # Check if cell is filled
                    current_row += 1  # Move down

                # Move one row below the last filled cell
                start_row = current_row + 1

                # Insert 12 blank rows after the last row written
                ws.insert_rows(current_row, 16)

                # reset the row count to copy / paste values
                current_row = 22
                while ws[f"A{current_row}"].value is not None:  # Check if cell is filled
                    current_row += 1  # Move down
                start_row = current_row + 1

                # Write the DataFrame headers to the sheet
                headers = list(first0_summary.columns)  # Get column headers

                # Write headers in the first row
                for col_idx, header in enumerate(headers, start=1):
                    ws.cell(row=start_row, column=col_idx, value=header)

                # Write the DataFrame contents to the sheet
                for row_idx, row_data in first0_summary.iterrows():
                    for col_idx, value in enumerate(row_data, start=1):
                        ws.cell(row=start_row + row_idx + 1,
                                column=col_idx, value=value)

                start_row = start_row + 4

                # Second level breakdown
                repetitions_2 = [10, 10, 10, 10, 10, 10, 5, 5, 5, 5]
                category_list_2 = []
                for category, rep in zip(category_2, repetitions_2):
                    category_list_2.extend([category] * rep)
                df['2nd-Order Category'] = category_list_2
                second0_summary = df.groupby(
                    ['1st-Order Category', '2nd-Order Category'])['Avg. Score (%)'].mean().reset_index().drop(columns='1st-Order Category')

                # use a custom sort for the summary table
                second0_summary['2nd-Order Category'] = pd.Categorical(
                    second0_summary['2nd-Order Category'], [
                        "Trust", "Health", "Relationships", "Impact", "Value", "Engagement", "See the Whole Playing Field", "Build Cultural Competency", "Give Power Away", "Take Bold, Courageous Action"
                    ])
                second0_summary = second0_summary.sort_values(
                    "2nd-Order Category", ascending=False)

                # Write headers in the first row of the new section
                for col_idx, header in enumerate(second0_summary.columns, start=1):
                    ws.cell(row=start_row, column=col_idx, value=header)

                # Write the DataFrame contents row by row
                for row_idx, row_data in second0_summary.iterrows():
                    for col_idx, value in enumerate(row_data, start=1):
                        ws.cell(row=start_row + row_idx + 1,
                                column=col_idx, value=value)

                # Third level breakdown
                repetitions_3 = [5, 5, 5, 5, 5, 5, 5, 5,
                                 5, 5, 5, 5, 2, 3, 2, 3, 2, 3, 2, 3]
                category_list_3 = []
                for category, rep in zip(category_3, repetitions_3):
                    category_list_3.extend([category] * rep)
                df['3rd-Order Category'] = category_list_3

                # Define the range of cells
                start_row = 24
                end_row = 119
                start_col = 1  # Column A (1-indexed)
                end_col = 7  # Column C (1-indexed)

                # Define the font style
                custom_font = Font(name="Arial", size=11)

                # Apply the font style to each cell in the range
                for row in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col):
                    for cell in row:
                        cell.font = custom_font

            else:  # Team template

                template = 'Team'
                # First level breakdown
                repetitions_1 = [60, 20]
                category_list_1 = []
                for category, rep in zip(category_1, repetitions_1):
                    category_list_1.extend([category] * rep)
                df['1st-Order Category'] = category_list_1
                first0_summary = df.groupby(
                    '1st-Order Category')['Avg. Score (%)'].mean().reset_index()

                # Start from A22 and find the last filled cell in column A
                current_row = 22
                while ws[f"A{current_row}"].value is not None:  # Check if cell is filled
                    current_row += 1  # Move down

                # Move one row below the last filled cell
                start_row = current_row + 1

                # Insert 12 blank rows after the last row written
                ws.insert_rows(current_row, 16)

                # reset the row count to copy / paste values
                current_row = 22
                while ws[f"A{current_row}"].value is not None:  # Check if cell is filled
                    current_row += 1  # Move down
                start_row = current_row + 1

                # Write the DataFrame headers to the sheet
                headers = list(first0_summary.columns)  # Get column headers

                # Write headers in the first row
                for col_idx, header in enumerate(headers, start=1):
                    ws.cell(row=start_row, column=col_idx, value=header)

                # Write the DataFrame contents to the sheet
                for row_idx, row_data in first0_summary.iterrows():
                    for col_idx, value in enumerate(row_data, start=1):
                        ws.cell(row=start_row + row_idx + 1,
                                column=col_idx, value=value)

                start_row = start_row + 4

                # Second level breakdown
                repetitions_2 = [10, 10, 10, 10, 10, 10, 5, 5, 5, 5]
                category_list_2 = []
                for category, rep in zip(category_2, repetitions_2):
                    category_list_2.extend([category] * rep)
                df['2nd-Order Category'] = category_list_2
                second0_summary = df.groupby(
                    ['1st-Order Category', '2nd-Order Category'])['Avg. Score (%)'].mean().reset_index().drop(columns='1st-Order Category')

                # use a custom sort for the summary table
                second0_summary['2nd-Order Category'] = pd.Categorical(
                    second0_summary['2nd-Order Category'], [
                        "Trust", "Health", "Relationships", "Impact", "Value", "Engagement", "See the Whole Playing Field", "Build Cultural Competency", "Give Power Away", "Take Bold, Courageous Action"
                    ])
                second0_summary = second0_summary.sort_values(
                    "2nd-Order Category", ascending=False)

                # Write headers in the first row of the new section
                for col_idx, header in enumerate(second0_summary.columns, start=1):
                    ws.cell(row=start_row, column=col_idx, value=header)

                # Write the DataFrame contents row by row
                for row_idx, row_data in second0_summary.iterrows():
                    for col_idx, value in enumerate(row_data, start=1):
                        ws.cell(row=start_row + row_idx + 1,
                                column=col_idx, value=value)

                # Third level breakdown
                repetitions_3 = [5, 5, 5, 5, 5, 5, 5, 5,
                                 5, 5, 5, 5, 2, 3, 2, 3, 2, 3, 2, 3]
                category_list_3 = []
                for category, rep in zip(category_3, repetitions_3):
                    category_list_3.extend([category] * rep)
                df['3rd-Order Category'] = category_list_3

                # Define the range of cells
                start_row = 24
                end_row = 119
                start_col = 1  # Column A (1-indexed)
                end_col = 7  # Column C (1-indexed)

                # Define the font style
                custom_font = Font(name="Arial", size=11)

                # Apply the font style to each cell in the range
                for row in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col):
                    for cell in row:
                        cell.font = custom_font

            # For the Leader and Team templates, there will be a 3rd order category
            if template == 'Leader' or template == 'Team':

                # label Column G, which won't be necessary for the other 2 templates
                ws["G23"] = "3rd-Order Category"
                ws["G23"].font = copy(ws["C23"].font)
                ws["G23"].fill = copy(ws["C23"].fill)
                ws["G23"].alignment = copy(ws["C23"].alignment)

                row = 24  # Starting from row 24
                for idx, row_data in df.iterrows():
                    ws[f"A{row}"] = row_data["Questions"]
                    ws[f"B{row}"] = row_data["Difficulty"]
                    ws[f"C{row}"] = row_data["Avg. Score (%)"]
                    ws[f"D{row}"] = row_data["Question Order"]
                    ws[f"E{row}"] = row_data["1st-Order Category"]
                    ws[f"F{row}"] = row_data["2nd-Order Category"]
                    ws[f"G{row}"] = row_data["3rd-Order Category"]
                    row += 1  # Move to the next row
                ws.column_dimensions['G'].width = 18

            else:
                row = 24  # Starting from row 24
                for idx, row_data in df.iterrows():
                    ws[f"A{row}"] = row_data["Questions"]
                    ws[f"B{row}"] = row_data["Difficulty"]
                    ws[f"C{row}"] = row_data["Avg. Score (%)"]
                    ws[f"D{row}"] = row_data["Question Order"]
                    ws[f"E{row}"] = row_data["1st-Order Category"]
                    ws[f"F{row}"] = row_data["2nd-Order Category"]
                    row += 1  # Move to the next row

            # any final adjustments to the table
            ws["C23"] = "Avg. Score (%)"
            ws.column_dimensions['A'].width = 50
            ws.column_dimensions['B'].width = 10
            ws.column_dimensions['C'].width = 16
            ws.column_dimensions['D'].width = 14
            ws.column_dimensions['E'].width = 18
            ws.column_dimensions['F'].width = 18

            # Add "_clean" suffix to the file name before the extension
            clean_file_name = f"{uploaded_file.name.rsplit('.', 1)[0]}_clean.xlsx"

            # Save the modified workbook to a BytesIO object
            cleaned_file = io.BytesIO()
            wb.save(cleaned_file)
            cleaned_file.seek(0)

            # Store the new name and file data
            files_to_download.append((clean_file_name, cleaned_file))

        # Handle downloading the files
        if len(files_to_download) > 1:
            # Create a ZIP file for multiple uploads
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for file_name, file_data in files_to_download:
                    zip_file.writestr(file_name, file_data.getvalue())
            zip_buffer.seek(0)

            # Provide download button for the ZIP file
            st.download_button(
                label="Download Files",
                data=zip_buffer,
                file_name="uploaded_files_clean.zip",
                mime="application/zip"
            )

        else:
            # Provide download button for a single file
            file_name, file_data = files_to_download[0]
            st.download_button(
                label="Download File",
                data=file_data,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# Run the app
if __name__ == "__main__":
    main()

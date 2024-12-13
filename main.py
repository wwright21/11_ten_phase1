import streamlit as st
import pandas as pd
import io
import zipfile
import openpyxl
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
        # Prepare list to hold file data
        files_to_download = []

        # Iterate through uploaded files
        for uploaded_file in uploaded_files:
            # Read the file into memory
            file_data = io.BytesIO(uploaded_file.read())

            # Load the workbook using openpyxl
            wb = openpyxl.load_workbook(file_data)
            ws = wb.active  # Assuming the data is on the active sheet

            # Add new columns & formatting
            ws["D23"] = "Question Order"
            ws["D23"].font = copy(ws["C23"].font)
            ws["D23"].fill = copy(ws["C23"].fill)
            ws["D23"].alignment = copy(ws["C23"].alignment)

            ws["E23"] = "Category"
            ws["E23"].font = copy(ws["C23"].font)
            ws["E23"].fill = copy(ws["C23"].fill)
            ws["E23"].alignment = copy(ws["C23"].alignment)

            ws["F23"] = "Category Average"
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

            # define Category sequence to be used by 3 of the 4 templates below
            categories = [
                'Trust', 'Health', 'Relationships', 'Impact', 'Value', 'Engagement', 'Just Leader'
            ]

            # Category sequence for the Team template
            categories2 = [
                'Trust-L', 'Trust-E', 'Health-L', 'Health-E', 'Relationships-L', 'Relationships-E', 'Impact-L', 'Impact-E', 'Value-L', 'Value-E', 'Engagement-L', 'Engagement-E', 'Just Leader'
            ]

            # fill in the 'Category' column for each template
            if df.shape[0] == 38:  # Review template

                # fill in the "Category" column
                repetitions = [5, 5, 5, 5, 5, 5, 8]
                category_list = []
                for category, rep in zip(categories, repetitions):
                    category_list.extend([category] * rep)
                df['Category'] = category_list
            elif df.shape[0] == 45:  # No leader template

                # fill in the "Category" column
                repetitions = [5, 10, 5, 5, 5, 5, 10]
                category_list = []
                for category, rep in zip(categories, repetitions):
                    category_list.extend([category] * rep)
                df['Category'] = category_list
            elif str(df["Questions"].iloc[0]).startswith("Q7"):  # Leader template

                # fill in the "Category" column
                repetitions = [10, 10, 10, 10, 10, 10, 20]
                category_list = []
                for category, rep in zip(categories, repetitions):
                    category_list.extend([category] * rep)
                df['Category'] = category_list
            else:  # Team template

                # fill in the "Category" column
                repetitions = [5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 20]
                category_list = []
                for category, rep in zip(categories2, repetitions):
                    category_list.extend([category] * rep)
                df['Category'] = category_list

            # Group by 'Categories' and calculate the mean 'Avg. Score (%)'
            grouped_df = df.groupby('Category')[
                'Avg. Score (%)'].mean().reset_index()

            # Merge the grouped DataFrame back to the original, aligning on 'Categories'
            df = df.merge(grouped_df, on='Category')

            # Rename the merged 'Avg. Score (%)' to 'Category Average'
            df = df.rename(columns={'Avg. Score (%)_y': 'Category Average'})
            df = df.rename(columns={'Avg. Score (%)_x': 'Avg. Score (%)'})

            # don't need averages for the Just Leader questions
            df.loc[df['Category'] == 'Just Leader',
                   'Category Average'] = np.nan

            # Write the sorted DataFrame back into the Excel sheet starting at A24
            row = 24  # Starting from row 24
            for idx, row_data in df.iterrows():
                ws[f"A{row}"] = row_data["Questions"]
                ws[f"B{row}"] = row_data["Difficulty"]
                ws[f"C{row}"] = row_data["Avg. Score (%)"]
                ws[f"D{row}"] = row_data["Question Order"]
                ws[f"E{row}"] = row_data["Category"]
                ws[f"F{row}"] = row_data["Category Average"]
                row += 1  # Move to the next row

            # any final adjustments to the table
            ws["C23"] = "Avg. Score (%)"
            ws.column_dimensions['B'].width = 10
            ws.column_dimensions['C'].width = 16
            ws.column_dimensions['D'].width = 14
            ws.column_dimensions['E'].width = 14
            ws.column_dimensions['F'].width = 16

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

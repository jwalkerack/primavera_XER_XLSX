import streamlit as st
from openpyxl import Workbook
import io
def main():
    st.title("Primavera XER -- Excel")
    st.sidebar.header("Choose a Transformation")

    option = st.sidebar.selectbox("Select an option", ("XER to Excel", "Excel to XER"))

    if option == "XER to Excel":
        xer_to_excel()
    elif option == "Excel to XER":
        st.warning("This feature is not yet implemented.")

import streamlit as st

def xer_to_excel():
    st.header("XER to Excel Transformation")

    # File uploader for the XER text file
    xer_file = st.file_uploader("Upload XER Text File", type=['xer'])

    # Trigger the transformation
    if st.button("Transform XER to Excel"):
        if xer_file:
            try:
                # Read the uploaded XER file and process it
                lines = read_file(xer_file)
                row_ids = Generate_List_index(lines)
                key_dict = Generate_dict_index(row_ids)
                data_dict = generate_data_dict(key_dict, lines)

                # Generate the Excel file in memory
                excel_file = create_dict_to_excel(data_dict)

                # Provide a download button for the generated Excel file
                st.download_button(
                    label="Download Excel File",
                    data=excel_file,
                    file_name="generated_file.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"An error occurred: {e}")
        else:
            st.error("Please upload a file to proceed.")


def read_file(file_obj):
    # Reads lines from the uploaded file-like object
    lines = file_obj.getvalue().decode('latin-1').splitlines()
    return lines

def Generate_List_index(lines):
    RowIds = []
    for line in lines:
        row_element = line[:2]
        RowIds.append(row_element)
    return RowIds

def Generate_dict_index(RowIds):
    DataDict = {}
    A = RowIds
    for Aindex, value in enumerate(A):
        if value == "%T":
            Table = Aindex
            D = []
            SecondPass = A[Aindex + 1:]
            for Bindex, Bvalue in enumerate(SecondPass):
                if Bvalue == "%F":
                    Fields = Aindex + 1
                elif Bvalue == "%R":
                    D.append((Aindex + 1) + Bindex)
                elif Bvalue == "%T" or Bvalue == "%E":
                    break
            DataDict[Table] = {'F': Fields , 'D':D }
    return DataDict
def generate_data_dict(gen_dict,rd_file):
    TableDict = {}
    for item in gen_dict:
        T = []
        f = gen_dict[item]['F']
        d = gen_dict[item]['D']

        table = rd_file[item].split("\t")[-1].strip()
        fields = rd_file[f].split("\t")[1:]
        fields[-1] = fields[-1].strip()
        T.append(fields)
        for v in d:
            row = rd_file[v].split("\t")[1:]
            row[-1] = row[-1].strip()
            T.append(row)
        TableDict[table] = T
    return TableDict

def write_dict_to_excel(data_dict, file_path):
    from openpyxl import Workbook

    # Create a new workbook
    wb = Workbook()

    # Remove the default sheet created in the workbook
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    for table_name, table_data in data_dict.items():
        # Add a new sheet with the table name
        ws = wb.create_sheet(title=table_name)

        # Write the data (header and rows) into the sheet
        for row in table_data:
            ws.append(row)

    # Save the workbook to the specified file path
    wb.save(file_path)
    print(f"Excel file created at: {file_path}")

def create_dict_to_excel(data_dict):
    # Create a new workbook
    wb = Workbook()

    # Remove the default sheet created in the workbook
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    for table_name, table_data in data_dict.items():
        # Add a new sheet with the table name
        ws = wb.create_sheet(title=table_name)

        # Write the data (header and rows) into the sheet
        for row in table_data:
            ws.append(row)

    # Save the workbook to an in-memory file
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)  # Move to the beginning of the in-memory file

    return output

if __name__ == "__main__":
    main()

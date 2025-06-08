import pandas as pd

def create_dummy_excel():
    file_path = "your_file.xlsx"

    # Create sample data for Sheet1 with an 'ID' column
    data1 = {
        'ID': [1, 2, 3, 4, 5, 6, 7],
        'Product': ['A', 'B', 'C', 'D', 'E', 'F', 'G'],
        'Sales': [100, 150, 200, 50, 300, 250, 120]
    }
    df1 = pd.DataFrame(data1)

    # Create sample data for Sheet2 with an 'ID' column
    # It will have some common IDs with Sheet1, some unique
    data2 = {
        'ID': [3, 4, 5, 8, 9, 10, 11],
        'Customer': ['X', 'Y', 'Z', 'P', 'Q', 'R', 'S'],
        'Region': ['North', 'South', 'East', 'West', 'North', 'East', 'South']
    }
    df2 = pd.DataFrame(data2)

    # Create a third sheet to ensure the main script only reads the first two
    data3 = {
        'ignore_me_id': [100, 200, 300],
        'data': ['x', 'y', 'z']
    }
    df3 = pd.DataFrame(data3)

    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df1.to_excel(writer, sheet_name='Orders', index=False)
            df2.to_excel(writer, sheet_name='Customer_Details', index=False)
            df3.to_excel(writer, sheet_name='Internal_Sheet', index=False)
        print(f"Dummy Excel file '{file_path}' created successfully with sheets: Orders, Customer_Details, Internal_Sheet and 'ID' columns.")
    except Exception as e:
        print(f"Error creating dummy Excel file: {e}")

if __name__ == "__main__":
    create_dummy_excel()

import pandas as pd
import os

input("Rename requires an Excel sheet that is structured as following:"
      "Column 1 must be named FileOLD, column 2 must be named FileNEW (simply insert a cell above and type the "
      "collumn names)."
      "The Excel sheet and the files that are to be renamed must be in the same folder as Rename. "
      "Press Enter to continue...")

def get_excel_path():

    while True:
        try:
            excel_path = input("Please enter the name of your Excel file (don't forget .xlsx at the end): ")
            if os.path.isfile(excel_path):
                return excel_path
            else:
                print("The file does not exist. Please try again")
        except KeyboardInterrupt:
            print("\nUser has canceled the entry.")
            exit(1)


excel_file_path = get_excel_path()
df = pd.read_excel(excel_file_path)

print(df.columns)


for index, row in df.iterrows():
    old_name = str(row['FileOLD']) + ".wav"
    new_name = str(row['FileNEW'])


    try:
        os.rename(old_name, new_name)
    except FileNotFoundError:
        print(f"File {old_name} not found.")
    except Exception as e:
        print(f"Error occurred: {e}")

#run auto-py-to-exe in console to convert the script into an executable
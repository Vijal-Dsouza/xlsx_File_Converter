import os
import sys
import pandas as pd
import requests
import re


def main():

    args = sys.argv
    if len(args) < 2 :
        print("Please Enter the file path")

    file_path =  args[1]


    if os.path.exists(file_path):
        file_name = os.path.basename(file_path).removesuffix(".xlsx")

    else:
        try:
            sheet_key = re.search(r'/d/([a-zA-Z0-9-_]+)', file_path).group(1)
            file_path = f"https://docs.google.com/spreadsheets/d/{sheet_key}/export?format=xlsx"
            resp = requests.get(file_path)
            # file_name = " ".join(resp.headers.get("Content-Disposition").split(";")[2].split("'")[2].split("%20"))

            file_name = "downloaded_sheet"
        except Exception:
            print("Cannot find sheet key. Provide correct URL")
   
    try:
        os.mkdir(file_name)
    except FileExistsError:
        print("File already Exists")


    buf = pd.ExcelFile(file_path)
    for sheet in buf.sheet_names:
        pd.read_excel(buf, sheet_name=sheet).to_csv(f"{file_name}/{sheet}.csv", index=False)
    
    print(f"Sheets have been saved under {file_name}/")
    
   
if __name__ == "__main__":
    main()
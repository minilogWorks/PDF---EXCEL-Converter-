import tabula
import pandas
import os

BASE_FOLDER_PATH = "C:\\Dev\\Python\\pdf_to_excel\\ATTENDANCE SHEET"
OUTPUT_FOLDER = "C:\\Dev\\Python\\pdf_to_excel\\OUTPUT"

#converts pdf file at src to xlsx file at dest
def pdf_2_excel(src, dest):
    df = tabula.read_pdf(src, pages='all')
    agg = pandas.concat(df[:])
    agg.to_excel(dest, index = False)
    # df.to_excel(dest)

# read files 2 folders deep from base_path 
# into base_path with new name
def collect_pdf_files(base_path):
    for dir in os.listdir(base_path):
        new_path = base_path + "\\" + dir
        for sub_dir in os.listdir(new_path):
            new_path_2 = new_path + "\\" + sub_dir
            os.chdir(new_path_2)
            for file in os.listdir():
                if file.endswith(".pdf"):
                    file_source_path = new_path_2 + "\\" + file
                    file_name_list = file.split(" ")
                    file_name_str = "".join(file_name_list)
                    new_file_name = sub_dir + "-" + file_name_str
                    file_dest_path = base_path + "\\" + new_file_name
                    os.rename(file_source_path, file_dest_path)

def convert_pdf_files_2_xlsx(base_path, out_path):
    os.chdir(base_path)
    for file in os.listdir():
        if file.endswith(".pdf"):
            file_name = file.split(".")[0]
            src_file_path = base_path + "\\" + file
            dest_file_path = out_path + "\\" + file_name + ".xlsx"
            pdf_2_excel(src_file_path, dest_file_path)
        
convert_pdf_files_2_xlsx(BASE_FOLDER_PATH, OUTPUT_FOLDER)

# pdf_2_excel("C:\\Dev\\Python\\pdf_to_excel\\ATTENDANCE SHEET\\ae1-AE153.pdf", "output.xlsx")



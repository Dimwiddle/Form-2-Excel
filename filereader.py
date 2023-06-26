import csv
from target_fields import target_fields, ignore_fields
import pandas as pd
import os, sys



class Converter(object):

    def __init__(self, file=None, folder=None):
        if file:
            self.file_path = file
        elif folder:
            self.folder_path = folder
        self.fields = target_fields
        self.ignore = ignore_fields

    def read_file(self, file=None):
        """Returns the list of lines in the events form"""
        if file:
            f = open(file, "r")
        else:
            f = open(self.file_path, "r")
        read = f.readlines()
        return read
    
    def save_as_text_file(self, text):
        filename = "email_temp.txt"
        new_file = open(filename, "w+")
        new_file.write(text)
        new_file.close()
        return filename
        
    def clean_data(self, line):
        clean_string = line.split(':')[1].lstrip()
        clean_string = clean_string.replace(" ","")
        clean_string = clean_string.strip("\n")
        return clean_string
    
    def convert_to_excel(self, destination, dictionary=None):
        convert_dict = self.convert_to_dict()
        dest = f"{destination}\\converted_form.xlsx" 
        if dictionary:
            convert_dict = dictionary
        df = pd.DataFrame().from_dict(convert_dict)
        df.to_excel(dest, engine='xlsxwriter')
    
    def convert_folder_to_excel(self, folder, destination):
        list_dict = self.convert_folder_to_dict(folder)
        excel_df = pd.DataFrame()
        dest = f"{destination}converted_forms.xlsx" 
        complete = [False, dest]
        for d in list_dict:
            df = pd.DataFrame().from_dict(d)
            excel_df = excel_df.append(df, ignore_index=True)
        try:
            excel_df.to_excel(dest, engine='xlsxwriter')
            print("Conversion Done")
            complete = [True, dest]
            return complete
        except:
            e = sys.exc_info()[0]
            print(f"Conversion failed. Please contact support... \n {e}")
            return complete

    
    def convert_to_dict(self, file=None):
        target_dict = {}
        arr = self.read_file(file=file)
        skip = False
        for line in arr:
            for field in self.fields:
                if field in line:
                    for ignore in self.ignore:
                        if ignore not in line:
                            data = self.clean_data(line)
                            target_dict[field] = data
                            break
                        else:
                            
                            break
        dict_object = [target_dict]
        return dict_object               

    def convert_folder_to_dict(self, folder, target_docx=None, target_pdf=None):
        list_dict = []
        target_folder = os.listdir(folder)
        for file in target_folder:
            file_ext = os.path.splitext(file)
            file_ext = file_ext[1]
            if file_ext == ".txt":
                find_file = folder + file
                list_dict.append(self.convert_to_dict(find_file)) 
        return list_dict

    # def read_file(self):
    #     """Returns the list of lines in the events form"""
    #     file_ext = os.path.splitext(self.file_path)
    #     file_ext = file_ext[1]
    #     if ".pdf" == file_ext:
    #         pdf_file = open(self.file_path, "rb")
    #         text = slate.PDF(pdf_file)
    #         temp_file = self.save_as_text(text[0])
    #         self.file_path = temp_file
    #     elif ".docx" == file_ext:
    #         read = ""
    #     f = open(self.file_path, "r")
    #     read = f.readlines()
    #     return read
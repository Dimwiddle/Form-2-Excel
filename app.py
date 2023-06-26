import tkinter as tk
from tkinter import filedialog, Text, RIGHT, RAISED, LEFT
import filereader
import os, sys


class FormApp(object):

    def __init__(self):
        self.root = tk.Tk()
        icon = tk.PhotoImage(file="icons\icons8-yellow-file-50.png")
        self.root.iconphoto(False, icon)
        self.root.title("Form2Excel")

        canvas = tk.Canvas(self.root, height=600, width=700, bg="#0b4f6c")
        canvas.pack()
        self.frame = tk.Frame(self.root, bg="white")
        self.frame.place(relwidth=0.8, relheight=0.8, rely=0.1, relx=0.1)
        label = tk.Label(self.frame, text="Hello! To begin please click the button: 'Select Folder To Convert'", bg="white")
        label.pack()
        openFolder = tk.Button(self.root, text="Select Folder To Convert", padx=10, pady=5, fg="white", bg="#757575", command=self.select_folder)
        openFolder.pack(side=LEFT)
        runConversion = tk.Button(self.root, text="Convert to Spreadsheet", padx=10, pady=5, fg="white", bg="#757575", command=self.launch_convert)
        runConversion.pack(side=RIGHT)
        self.root.mainloop()
    
    def select_folder(self):
        for widget in self.frame.winfo_children():
            widget.destroy()

        self.folder = filedialog.askdirectory(initialdir="/", title="Select Folder")
        folder_text = f"Folder Selected: {self.folder}. \n Ready to convert to a spreadsheet"
        label = tk.Label(self.frame, text=folder_text,  bg="white")
        label.pack()
        return self.folder
    
    def launch_convert(self):
        folder = f"{self.folder}/"
        complete = [False, "temp_file_path"]
        try:
            complete = filereader.Converter().convert_folder_to_excel(folder=folder, destination=folder)
            label = tk.Label(self.frame, text="Running...",  bg="white")
            label.pack()
            if complete[0]:
                label = tk.Label(self.frame, text=f"Job completed successfully \n Spreadsheet is located in: {complete[1]}",  bg="white")
                label.pack()
                os.startfile(complete[1])
            else:
                assertion_error = "Could not convert folder. Please try again"
                raise AssertionError(assertion_error)
        except:
            if assertion_error:
                e = assertion_error
            else:
                e = sys.exc_info()[0]
            label = tk.Label(self.frame, text=f"Job failed... \n {e}",  bg="white")
            label.pack()
        
        
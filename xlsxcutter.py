import pandas as pd
import os
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox

def select_input_file():
    """Open a file dialog to select the input file and update the input file path entry widget."""
    file_path = filedialog.askopenfilename()
    if not file_path.endswith('.xlsx'):
        messagebox.showwarning("Warning", "The selected file is not an xlsx file.")
        input_file_path_entry.delete(0, tk.END) #clear the content of the input file path entry
        return
    input_file_path_entry.delete(0, tk.END)
    input_file_path_entry.insert(0, file_path)

def select_output_folder():
    """Open a directory dialog to select the output folder and update the output folder path entry widget."""
    folder_path = filedialog.askdirectory()
    output_folder_path_entry.delete(0, tk.END)
    output_folder_path_entry.insert(0, folder_path)

def split_excel_file():
    """Read the specified sheet from the input Excel file into a pandas DataFrame,
    split the DataFrame into chunks of the specified number of rows (keeping the header of the original file),
    and write each chunk to a new Excel file in the output folder."""
    input_file_path = input_file_path_entry.get()
    output_folder_path = output_folder_path_entry.get()
    input_sheet_name = input_sheet_name_entry.get()
    rows_per_file = int(rows_per_file_entry.get())

    df = pd.read_excel(input_file_path, sheet_name=input_sheet_name, header=0)

    for i, chunk in enumerate(np.array_split(df, len(df) // rows_per_file + 1)):
        output_file_name = f"output_file_{i+1}.xlsx"
        output_file_path = os.path.join(output_folder_path, output_file_name)
        chunk.to_excel(output_file_path, index=False)

#Check the state of the variables to enable/disable the button
def check_variables():
    if all((input_file_path_entry.get(), output_folder_path_entry.get(), input_sheet_name_entry.get())):
        split_button.configure(state="normal")
    else:
        split_button.configure(state="disabled")

#Create the main window
window = tk.Tk()
image = tk.PhotoImage(file=".\knife.png")
window.iconphoto(True, image)
window.title("Excel File Cutter")

#Create the widgets for selecting the input file
input_file_label = tk.Label(window, text="What file to chop:")
input_file_label.grid(row=0, column=0, sticky="w")
input_file_path_entry = tk.Entry(window)
input_file_path_entry.grid(row=0, column=1, sticky="we")
input_file_browse_button = tk.Button(window, text="Browse", command=select_input_file)
input_file_browse_button.grid(row=0, column=2)

#Create the widgets for selecting the output folder
output_folder_label = tk.Label(window, text="Where to save chops:")
output_folder_label.grid(row=1, column=0, sticky="w")
output_folder_path_entry = tk.Entry(window)
output_folder_path_entry.grid(row=1, column=1, sticky="we")
output_folder_browse_button = tk.Button(window, text="Browse", command=select_output_folder)
output_folder_browse_button.grid(row=1, column=2)

#Create the widget for specifying the sheet to read
input_sheet_name_label = tk.Label(window, text="Excel sheet name:")
input_sheet_name_label.grid(row=2, column=0, sticky="w")
input_sheet_name_entry = tk.Entry(window)
input_sheet_name_entry.grid(row=2, column=1, sticky="we")

#Create the widget for specifying the number of rows per file
rows_per_file_label = tk.Label(window, text="Number of rows:")
rows_per_file_label.grid(row=3, column=0, sticky="w")
rows_per_file_entry = tk.Entry(window)
rows_per_file_entry.grid(row=3, column=1, sticky="we")

#Create the button to start the splitting process
split_button = tk.Button(window, text="Cut xlsx!", command=split_excel_file, state="disabled")
split_button.grid(row=3, column=1)

#Bind the check_variables function to any change in the variables
input_file_path_entry.bind("<KeyRelease>", lambda event: check_variables())
output_folder_path_entry.bind("<KeyRelease>", lambda event: check_variables())
input_sheet_name_entry.bind("<KeyRelease>", lambda event: check_variables())
split_button.grid(row=5, column=1)

#Set the window to resizeable
window.grid_columnconfigure(1, weight=1)
window.resizable(False, False)

# Set the title and geometry of the window
window.geometry("264x128")

#Start the GUI event loop
window.mainloop()

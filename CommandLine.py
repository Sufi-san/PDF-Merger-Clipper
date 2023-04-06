import tkinter as tk
import os
from tkinter import filedialog, messagebox
from PyPDF2 import PdfWriter

root = tk.Tk()
root.withdraw()
merger = PdfWriter()

contains_other = False
output = None
fact = int(input("For Merging Press '1', for Clipping Press '2': "))
print(os.getcwd())

if fact == 1:
    file_paths = filedialog.askopenfilenames()
    for file_path in file_paths:
        if os.path.splitext(file_path)[1] == '.pdf':
            file_name = os.path.basename(file_path)
            merger.append(file_path)
        else:
            contains_other = True
    output = open("Merged Files/Merged.pdf", "wb")
    merger.write(output)
    route = os.path.join(os.getcwd(), "Merged Files", "Merged.pdf")
    print(os.getcwd())
    print(route)
    os.startfile(route)
elif fact == 2:
    print("Clip initiated")
    file_path = filedialog.askopenfilename()
    start = int(input("Enter Start Page: "))
    end = int(input("Enter End Page: "))
    if os.path.splitext(file_path)[1] == '.pdf':
        file_name = os.path.basename(file_path)
        merger.merge(position=0, fileobj=file_path, pages=(start-1, end))
    else:
        contains_other = True
    output = open("Clipped Files/Clipped.pdf", "wb")
    merger.write(output)
if contains_other:
    messagebox.showinfo("Info", "Non-pdf files were removed.")

print(merger)
merger.close()
if output:
    output.close()


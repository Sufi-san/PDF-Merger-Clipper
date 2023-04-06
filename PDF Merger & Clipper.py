import os.path
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from PIL import ImageTk, Image
from PyPDF2 import PdfWriter


# Function to Select Files from anywhere in the system and append names in the listbox.
def select_files(action):
    global file_paths
    global file_path
    if action == 1:  # Indicates files are selected in Merge Section
        not_pdf = False
        new_paths = filedialog.askopenfilenames()  # tuple of newly selected files
        new_paths = sorted(new_paths, key=os.path.getmtime)
        file_paths += tuple(new_paths)  # tuple of old and new selected files
        file_paths_list = list(file_paths)  # converted file_paths tuple to list
        if new_paths:
            for path in file_paths:
                if os.path.splitext(path)[1] != ".pdf":  # excluding non-pdf files
                    file_paths_list.remove(path)
                    not_pdf = True
            file_paths = tuple(file_paths_list)  # updating file_paths tuple using modified file_paths_list
            # The above procedure is used many times for updating file_paths tuple
            if not_pdf:
                messagebox.showinfo("Only PDF Allowed", "Non-pdf files have been removed.")
            if file_paths:
                ls_box1.delete(0, END)
                for path in file_paths:
                    file_name = os.path.basename(path)
                    ls_box1.insert(END, file_name)
        else:
            messagebox.showinfo("Message", "No files selected.")

    elif action == 2:  # Indicates file selection in Clip Section
        if ls_box2.size() < 1:
            file_path = filedialog.askopenfilename()
            if file_path == "":
                messagebox.showinfo("Message", "No File Selected")
            else:
                if os.path.splitext(file_path)[1] != ".pdf":
                    file_path = ""
                    messagebox.showinfo("Only PDF Allowed", "Please select PDF files only")
                elif file_path == "":
                    messagebox.showinfo("Message", "No file selected.")
                else:
                    file_name = os.path.basename(file_path)
                    ls_box2.insert(END, file_name)
        else:
            messagebox.showinfo("Selection Limit", "Only 1 File can be selected for clipping.")


def remove_selected():  # Deletes selected items inside the ls_box1 list box and file_paths tuple
    global file_paths
    selection = ls_box1.curselection()  # Getting indices of items selected by user
    if selection:
        file_paths_list = list(file_paths)
        for index in reversed(selection):
            file_paths_list[index] = ""
        file_paths_list = [path for path in file_paths_list if path]
        file_paths = tuple(file_paths_list)  # Deleted selected items in file_paths tuple
        ls_box1.delete(0, END)  # Delete all elements of ls_box1 listbox
        for path in file_paths:
            file_name = os.path.basename(path)
            ls_box1.insert(END, file_name)  # Updating ls_box listbox as per updated file_paths tuple
    else:
        messagebox.showinfo("Message", "Please Select Files to Remove.")


def update_selected():  # Update selected items inside the ls_box1 list box and file_paths tuple
    global file_paths
    selection = ls_box1.curselection()  # Getting indices of items selected by user
    if selection:
        new_path = filedialog.askopenfilename()
        file_paths_list = list(file_paths)
        for index in reversed(selection):
            file_paths_list[index] = new_path
        file_paths = tuple(file_paths_list)  # Updated selected items in file_paths tuple
        ls_box1.delete(0, END)  # Delete all elements of ls_box1 listbox
        for path in file_paths:
            file_name = os.path.basename(path)
            ls_box1.insert(END, file_name)  # Updating ls_box listbox as per updated file_paths tuple
    else:
        messagebox.showinfo("Message", "Please select Files to Update.")


def pdf_merge():  # Merge all files selected by user.

    def merge_process():
        merger = PdfWriter()
        for path in file_paths:
            merger.append(path)
        txt_lbl1.config(text="Files Merged!")
        icon_lbl1.config(image=file_icon)
        txt_lbl3.config(text=f"{f_name}.pdf")
        btn_5.config(state=NORMAL)
        output = open(f"Merged Files/{f_name}.pdf", "wb")
        merger.write(output)
        merger.close()

    global file_paths
    if ls_box1.size() > 1:
        f_name = res_name_entry1.get()
        if f_name != "":
            exist_f = os.path.join("Merged Files", f"{f_name}.pdf")
            if not os.path.exists(exist_f):
                merge_process()
            else:
                response = messagebox.askquestion("File Already Exists", "Do you want to overwrite existing file?")
                if response == "yes":
                    merge_process()
                else:
                    messagebox.showinfo("Message", "Please Select a different file name for resultant file.")
        else:
            messagebox.showinfo("No Name", "Please Enter name for the resultant file and Try Again.")
    else:
        messagebox.showinfo("Files Missing", "Please Select more than 1 file to merge.")


def pdf_clip():  # Clips selected file according to start and end page entered.

    def clip_process():
        clipper = PdfWriter()
        start = int(start_entry.get())
        end = int(end_entry.get())
        clipper.merge(position=0, fileobj=file_path, pages=(start-1, end))
        txt_lbl2.config(text="File Clipped!")
        icon_lbl2.config(image=file_icon)
        txt_lbl4.config(text=f"{f_name}.pdf")
        btn_9.config(state=NORMAL)
        output = open(f"Clipped Files/{f_name}.pdf", "wb")
        clipper.write(output)
        clipper.close()

    global file_path
    if ls_box2.size() == 1:
        f_name = res_name_entry2.get()
        if f_name and start_entry.get() and end_entry.get():
            exist_f = os.path.join("Clipped Files", f"{f_name}.pdf")
            if not os.path.exists(exist_f):
                clip_process()
            else:
                response = messagebox.askquestion("File Already Exists", "Do you want to overwrite existing file?")
                if response == "yes":
                    clip_process()
                else:
                    messagebox.showinfo("Message", "Please Select a different file name for resultant file.")
        else:
            messagebox.showinfo("Data Incomplete", "Please fill all the required fields and Try Again.")
    else:
        messagebox.showinfo("Files Missing", "Please Select 1 file to clip.")


def view_files(action):  # View the resulting merged/clipped file.
    if action == 1:
        try:
            path = os.path.join(os.getcwd(), "Merged Files", f"{res_name_entry1.get()}.pdf")
            os.startfile(path)
        except FileNotFoundError:
            messagebox.showinfo("File Not Found", "File does not exist or has been removed.")
    elif action == 2:
        try:
            path = os.path.join(os.getcwd(), "Clipped Files", f"{res_name_entry2.get()}.pdf")
            os.startfile(path)
        except FileNotFoundError:
            messagebox.showinfo("File Not Found", "File does not exist or has been removed.")


def clear_data(action):  # Clears all the information filled/selected by user
    global file_paths
    global file_path
    if action == 1:
        file_paths = ()
        ls_box1.delete(0, END)
        res_name_entry1.delete(0, END)
        btn_5.config(state=DISABLED)
        txt_lbl1.config(text="")
        icon_lbl1.config(image="")
        txt_lbl3.config(text="")
    elif action == 2:
        file_path = ""
        ls_box2.delete(0, END)
        res_name_entry2.delete(0, END)
        start_entry.delete(0, END)
        end_entry.delete(0, END)
        btn_9.config(state=DISABLED)
        txt_lbl2.config(text="")
        icon_lbl2.config(image="")
        txt_lbl4.config(text="")


def validate_if_num(new_value):  # Restricts text inside the entry to be integers only
    if new_value.isdigit() or new_value == "":
        return True
    else:
        return False


def open_folder(action):
    if action == 1:
        os.startfile("Merged Files")
    elif action == 2:
        os.startfile("Clipped Files")


# Basic information of root window
root = Tk()
root.title("PDF Merger & Clipper")
root.geometry(f"545x{root.winfo_screenheight()}+39+0")
root.resizable(False, False)
# Providing Image paths
root.iconbitmap("Images/exe_icon.ico")
bg_pic = ImageTk.PhotoImage(Image.open("Images/bg_img.jpg"))
file_icon = ImageTk.PhotoImage(Image.open("Images/file_icon.jpg"))
root_icon = ImageTk.PhotoImage(Image.open("Images/root_icon.jpg"))
# Creating module-level variables to use as global inside functions
file_paths = ()
file_path = ""
# Creating a frame widget covering entire root window
main_frame = Frame(root)
main_frame.pack(fill=BOTH, expand=1)

my_canvas = Canvas(main_frame)  # Placing a canvas widget on the main_frame widget
my_canvas.pack(side=LEFT, fill=BOTH, expand=1)

my_scr_bar = ttk.Scrollbar(main_frame, orient=VERTICAL, command=my_canvas.yview)  # Attaching scrollbar to main_frame
my_scr_bar.pack(side=RIGHT, fill=Y)

my_canvas.configure(yscrollcommand=my_scr_bar.set)  # Binding scrollbar to canvas
my_canvas.bind("<Configure>", lambda e: my_canvas.configure(scrollregion=my_canvas.bbox("all")))

root_frame = Frame(my_canvas)  # Creating another frame widget on the canvas widget

my_canvas.create_window((0, 0), window=root_frame, anchor="nw")

# Creating label widgets to provide basic software information.
bg_lbl = Label(root_frame, image=bg_pic)
bg_lbl.pack(fill=BOTH)
icon_lbl = Label(root_frame, image=root_icon)
icon_lbl.place(x=215, y=10, width=100, height=100)
title_lbl = Label(root_frame, text="Merger & Clipper", fg="#ffffff", font=("Impact", 20), bg="#080808")
title_lbl.place(x=1, y=110, width=527, height=30)

# Creating and placing required frame widgets.
merge_frame = LabelFrame(root_frame, text="Merge Section", bg="#c8c8c8")
merge_frame.place(x=10, y=150, width=510, height=265)
clip_frame = LabelFrame(root_frame, text="Clip Section", bg="#c8c8c8")
clip_frame.place(x=10, y=450, width=510, height=225)

# Creating and placing required button widgets.
btn_1 = Button(root_frame, text="Select Files", command=lambda: select_files(1))
btn_1.place(x=20, y=170, width=77, height=25)
btn_2 = Button(root_frame, text="Remove Selected", command=remove_selected)
btn_2.place(x=20, y=310, width=112, height=30)
btn_3 = Button(root_frame, text="Update Selected", command=update_selected)
btn_3.place(x=160, y=310, width=113, height=30)
btn_4 = Button(root_frame, text="Merge PDFs", command=pdf_merge)
btn_4.place(x=20, y=380, width=285, height=30)
btn_5 = Button(root_frame, text="View File", state=DISABLED, command=lambda: view_files(1))
btn_5.place(x=340, y=380, width=75, height=30)
btn_6 = Button(root_frame, text="Clear Data", command=lambda: clear_data(1))
btn_6.place(x=430, y=380, width=75, height=30)
btn_7 = Button(root_frame, text="Select File", command=lambda: select_files(2))
btn_7.place(x=20, y=480, width=77, height=25)
btn_8 = Button(root_frame, text="Clip PDF", command=pdf_clip)
btn_8.place(x=20, y=640, width=285, height=30)
btn_9 = Button(root_frame, text="View File", state=DISABLED, command=lambda: view_files(2))
btn_9.place(x=340, y=640, width=75, height=30)
btn_10 = Button(root_frame, text="Clear Data", command=lambda: clear_data(2))
btn_10.place(x=430, y=640, width=75, height=30)
btn_11 = Button(root_frame, text="Open Save Location", command=lambda: open_folder(1))
btn_11.place(x=160, y=164, width=145, height=30)
btn_12 = Button(root_frame, text="Open Save Location", command=lambda: open_folder(2))
btn_12.place(x=160, y=468, width=145, height=33)

# Creating and placing required label widgets.
line_lbl1 = Label(root_frame, bg="#000000")
line_lbl1.place(x=0, y=140, width=530, height=1)
line_lbl2 = Label(root_frame, bg="#000000")
line_lbl2.place(x=310, y=160, width=1, height=247)
line_lbl3 = Label(root_frame, bg="#000000")
line_lbl3.place(x=0, y=430, width=530, height=1)
line_lbl4 = Label(root_frame, bg="#000000")
line_lbl4.place(x=310, y=460, width=1, height=200)
icon_lbl1 = Label(root_frame, bg="#c8c8c8")
icon_lbl1.place(x=340, y=210, width=161, height=114)
icon_lbl2 = Label(root_frame, bg="#c8c8c8")
icon_lbl2.place(x=340, y=490, width=161, height=114)
txt_lbl1 = Label(root_frame, text="", font=("Times", 15), bg="#c8c8c8")
txt_lbl1.place(x=360, y=180, width=128, height=30)
txt_lbl2 = Label(root_frame, text="", font=("Times", 15), bg="#c8c8c8")
txt_lbl2.place(x=360, y=460, width=120, height=30)
txt_lbl3 = Label(root_frame, text="", font=("Times", 10), bg="#c8c8c8")
txt_lbl3.place(x=340, y=330, width=161, height=30)
txt_lbl4 = Label(root_frame, text="", font=("Times", 10), bg="#c8c8c8")
txt_lbl4.place(x=340, y=610, width=161, height=25)
txt_lbl5 = Label(root_frame, text="Start Page:", font=("Times", 12), bg="#c8c8c8")
txt_lbl5.place(x=20, y=550, width=70, height=25)
txt_lbl6 = Label(root_frame, text="End Page:", font=("Times", 12), bg="#c8c8c8")
txt_lbl6.place(x=160, y=550, width=70, height=25)
txt_lbl7 = Label(root_frame, text="Resultant File Name:", font=("Times", 12), bg="#c8c8c8")
txt_lbl7.place(x=20, y=350)
txt_lbl8 = Label(root_frame, text="Resultant File Name:", font=("Times", 12), bg="#c8c8c8")
txt_lbl8.place(x=20, y=600)
txt_lbl9 = Label(root_frame, text=".pdf", font=("Caliber", 12), bg="#c8c8c8")
txt_lbl9.place(x=270, y=348)
txt_lbl10 = Label(root_frame, text=".pdf", font=("Caliber", 12), bg="#c8c8c8")
txt_lbl10.place(x=270, y=598)

# Creating and placing required listbox widgets.
ls_box1 = Listbox(root_frame, selectmode=MULTIPLE)
ls_box1.place(x=20, y=200, width=253, height=101)
ls_box2 = Listbox(root_frame)
ls_box2.place(x=20, y=510, width=253, height=30)

# Variable required for validate command.
validate = root_frame.register(validate_if_num)

# Creating and placing required entry widgets.
start_entry = Entry(root_frame, justify=CENTER, validate="key", validatecommand=(validate, "%P"))
start_entry.place(x=90, y=550, width=37, height=30)
end_entry = Entry(root_frame, justify=CENTER, validate="key", validatecommand=(validate, "%P"))
end_entry.place(x=230, y=550, width=37, height=30)
res_name_entry1 = Entry(root_frame, justify=CENTER)
res_name_entry1.place(x=160, y=348, width=110, height=25)
res_name_entry2 = Entry(root_frame, justify=CENTER)
res_name_entry2.place(x=160, y=598, width=110, height=25)

root.mainloop()

from tkinter import filedialog
from tkinter import *

def browse_button1():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global file_path1
    # global file_path2

    filename = filedialog.askopenfilename(title = "Select file",filetypes = (("xls files","*.xls"),("all files","*.*")))
    file_path1.set(filename)
    print(filename)

def browse_button2():
    # Allow user to select a directory and store it in global var
    # called folder_path
    # global file_path1
    global file_path2

    filename = filedialog.askopenfilename(title = "Select file",filetypes = (("xls files","*.xls"),("all files","*.*")))
    file_path2.set(filename)
    print(filename)

root = Tk()
root.geometry('600x450')
file_path1 = StringVar()
file_path2 = StringVar()
lbl1 = Label(master=root,textvariable=file_path1,width=40)
lbl1.grid(row=2, column=1)
button2 = Button(text="Browse old file", command=browse_button1)
button2.grid(row=2, column=4)
lbl1 = Label(master=root,textvariable=file_path2,width=40)
lbl1.grid(row=4, column=1)
button2 = Button(text="Browse new file", command=browse_button2)
button2.grid(row=4, column=4)

mainloop()

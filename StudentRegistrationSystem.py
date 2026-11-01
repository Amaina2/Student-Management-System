from tkinter import * 
from datetime import date
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk
import os
from tkinter.ttk import Combobox
import openpyxl , xlrd
from openpyxl import Workbook
import pathlib

background = "#0C2340"
button_color = "#C9A227"
button_text = "#FFFFFF"
framebg = "grey"
framefg = "#000000"

root = Tk()
root.title("Student Registration System")
root.geometry("1250x700+210+100")
root.config(bg=background)

file = pathlib.Path('student_data.xlsx')
if file.exists():
    pass
else:
    file = Workbook()
    sheet = file.active
    sheet['A1'] = "Registration No."
    sheet['B1'] = "Name"
    sheet['C1'] = "Classification"
    sheet['D1'] = "Gender"
    sheet['E1'] = "Date of Birth"
    sheet['F1'] = "Date of Registration"
    sheet['G1'] = "Ethnicity"
    sheet['H1'] = "Contact No."
    sheet['I1'] = "Major"
    sheet['J1'] = "Car registration No."
    sheet['K1'] = "Hostel room No."
    sheet['L1'] = "Occupation"

    file.save('student_data.xlsx')

student_images_dir = pathlib.Path('Student Images')
student_images_dir.mkdir(exist_ok=True)

#Exit window
def Exit():
    root.destroy()


def get_existing_image_path(registration_number):
    if registration_number is None:
        return None

    matches = sorted(student_images_dir.glob(f'{registration_number}.*'))
    for match in matches:
        if match.suffix.lower() in {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff', '.webp'}:
            return match
    return None

#Upload image
def showimage():
    global filename
    global img
    filename = filedialog.askopenfilename(
        initialdir=os.getcwd(),
        title="Select Image File",
        filetypes=(("Image Files", "*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.tif;*.tiff;*.webp"), ("All Files", "*.*"))
    )

    if not filename:
        return

    img = Image.open(filename)
    resized_image = img.resize((190, 190))
    photo2 = ImageTk.PhotoImage(resized_image)
    lbl.config(image = photo2)
    lbl.image = photo2

#Automate entering of registration number
def registration_no():
    file = openpyxl.load_workbook('student_data.xlsx')
    sheet = file.active
    row = sheet.max_row

    max_row_value = sheet.cell(row=row, column=1).value

    try:
        Registration.set(max_row_value + 1)

    except:
        Registration.set("1")

#Clear
def clear():
    global img
    Name.set('')
    DOB.set('DD/MM/YYYY')
    dob_entry.config(fg = '#8B8B8B')
    Ethnicity.set('')
    Major.set('')
    ContactNo.set('')
    CarRegNo.set('')
    HostelRoomNo.set('')
    Occupation.set('')
    Classification.set("Select Classification")

    registration_no()

    saveButton.config(state = 'normal')

    img1 = PhotoImage(file = "Images/upload photo.png")
    lbl.config(image = img1)
    lbl.image = img1

    img = ""

#Save to excel file
def save():
    global gender
    R1 = Registration.get()
    N1 = Name.get()
    C1 = Classification.get()
    try:
        G1 = gender
    except:
        messagebox.showerror("Error", "Please select gender")
    D2 = DOB.get()
    D1 = Date.get()
    E1 = Ethnicity.get()
    C2 = ContactNo.get()
    M1 = Major.get()
    C3 = CarRegNo.get()
    H1 = HostelRoomNo.get()
    O1 = Occupation.get()

    if N1 == "" or C1 == "Select Classification" or D2 == "" or E1 == "" or C2 == "" or M1 == "" or C3 == "" or H1 == "" or O1 == "":
        messagebox.showerror("Error", "Please fill all the fields")
    else:
        file = openpyxl.load_workbook('student_data.xlsx')
        sheet = file.active
        sheet.cell(column=1, row=sheet.max_row + 1, value=R1)
        sheet.cell(column=2, row=sheet.max_row, value=N1)
        sheet.cell(column=3, row=sheet.max_row, value=C1)
        sheet.cell(column=4, row=sheet.max_row, value=G1)
        sheet.cell(column=5, row=sheet.max_row, value=D2)
        sheet.cell(column=6, row=sheet.max_row, value=D1)
        sheet.cell(column=7, row=sheet.max_row, value=E1)
        sheet.cell(column=8, row=sheet.max_row, value=C2)
        sheet.cell(column=9, row=sheet.max_row, value=M1)
        sheet.cell(column=10, row=sheet.max_row, value=C3)
        sheet.cell(column=11, row=sheet.max_row, value=H1)
        sheet.cell(column=12, row=sheet.max_row, value=O1)

        file.save('student_data.xlsx')

        try:
            if 'filename' in globals() and filename:
                extension = pathlib.Path(filename).suffix.lower() or '.jpg'
                if extension not in {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff', '.webp'}:
                    extension = '.jpg'
                image_path = student_images_dir / f'{R1}{extension}'
                img.save(image_path)
                messagebox.showinfo("Info", "Data saved successfully")
            else:
                messagebox.showinfo("Info", "Data saved successfully. No image was selected.")
        except Exception:
            messagebox.showinfo("Info", "Image is not available")


        clear()

        registration_no()

#Search student data
def search():
    text = Search.get().strip()
    if text == "":
        messagebox.showerror("Error", "Please enter a registration number")
        return

    try:
        search_number = int(text)
    except ValueError:
        messagebox.showerror("Error", "Registration number must be numeric")
        return

    clear()
    saveButton.config(state = 'disabled')

    file = openpyxl.load_workbook('student_data.xlsx')
    sheet = file.active
    found_row = None

    for row in sheet.iter_rows(min_row=2, values_only=False):
        if row[0].value == search_number:
            found_row = row
            break

    if found_row is None:
        messagebox.showerror("Invalid", "Registration number not found")
        return

    x1 = sheet.cell(row=found_row[0].row, column=1).value
    x2 = sheet.cell(row=found_row[0].row, column=2).value
    x3 = sheet.cell(row=found_row[0].row, column=3).value
    x4 = sheet.cell(row=found_row[0].row, column=4).value
    x5 = sheet.cell(row=found_row[0].row, column=5).value
    x6 = sheet.cell(row=found_row[0].row, column=6).value
    x7 = sheet.cell(row=found_row[0].row, column=7).value
    x8 = sheet.cell(row=found_row[0].row, column=8).value
    x9 = sheet.cell(row=found_row[0].row, column=9).value
    x10 = sheet.cell(row=found_row[0].row, column=10).value
    x11 = sheet.cell(row=found_row[0].row, column=11).value
    x12 = sheet.cell(row=found_row[0].row, column=12).value

    Registration.set(x1)
    Name.set(x2)
    Classification.set(x3)

    if x4 == "Female":
        R2.select()
    else:
        R1.select()

    DOB.set(x5)
    dob_entry.config(fg = '#000000')
    Date.set(x6)
    Ethnicity.set(x7)
    ContactNo.set(x8)
    Major.set(x9)
    CarRegNo.set(x10)
    HostelRoomNo.set(x11)
    Occupation.set(x12)

    existing_image_path = get_existing_image_path(str(x1))
    if existing_image_path is not None:
        try:
            img = Image.open(existing_image_path)
            resized_image = img.resize((190, 190))
            photo2 = ImageTk.PhotoImage(resized_image)
            lbl.config(image = photo2)
            lbl.image = photo2
        except Exception:
            pass

#Update
def update():
    R1 = Registration.get()
    N1 = Name.get()
    C1 = Classification.get()
    selection()
    G1 = gender
    D2 = DOB.get()
    D1 = Date.get()
    E1 = Ethnicity.get()
    C2 = ContactNo.get()
    M1 = Major.get()
    C3 = CarRegNo.get()
    H1 = HostelRoomNo.get()
    O1 = Occupation.get()

    file = openpyxl.load_workbook('student_data.xlsx')
    sheet = file.active

    for row in sheet.iter_rows(min_row=2, values_only=False):
        if row[0].value == R1:
            reg_number = row[0].row
            break
    else:
        messagebox.showerror("Invalid", "Registration number not found")
        return

    sheet.cell(column=1, row=reg_number, value=R1)
    sheet.cell(column=2, row=reg_number, value=N1)
    sheet.cell(column=3, row=reg_number, value=C1)
    sheet.cell(column=4, row=reg_number, value=G1)
    sheet.cell(column=5, row=reg_number, value=D2)
    sheet.cell(column=6, row=reg_number, value=D1)
    sheet.cell(column=7, row=reg_number, value=E1)
    sheet.cell(column=8, row=reg_number, value=C2)
    sheet.cell(column=9, row=reg_number, value=M1)
    sheet.cell(column=10, row=reg_number, value=C3)
    sheet.cell(column=11, row=reg_number, value=H1)
    sheet.cell(column=12, row=reg_number, value=O1)

    file.save('student_data.xlsx')

    try:
        existing_image_path = get_existing_image_path(str(R1))
        extension = pathlib.Path(filename).suffix.lower() if 'filename' in globals() and filename else (existing_image_path.suffix.lower() if existing_image_path is not None else '.jpg')
        if extension not in {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff', '.webp'}:
            extension = '.jpg'
        image_path = student_images_dir / f'{R1}{extension}'
        if 'img' in globals() and img is not None and filename:
            img.save(image_path)
        elif existing_image_path is not None:
            if existing_image_path != image_path:
                existing_image_path.rename(image_path)
    except Exception:
        pass

    messagebox.showinfo("Info", "Data updated successfully")

    clear()

        

#gender selection
def selection():
    global gender
    value = radio.get()
    if value == 1:
        gender = "Male"
    else:
        gender = "Female"


#Top Frame
Label(root, text = "Email: austinmaina05@gmail.com", width = 10, height = 3, bg = "#000000", fg = "#FFFFFF", anchor = "e").pack(side = TOP, fill = X)
header_title = Label(root, text = "STUDENT REGISTRATION", width = 22, height = 2, bg = background, fg = "#FFFFFF", font = ("Helvetica", 20, "bold"))
header_title.place(relx = 0.5, y = 45, anchor = 'n')

#search box to update and delete student data
Search = StringVar()
search_entry = Entry(root, textvariable = Search, width = 15, bd = 2, font = ("Helvetica", 20), bg = "#F5F5F5")
imageicon3 = PhotoImage(file = "Images/search.png")
srch = Button(root, text = "Search", compound = LEFT, image = imageicon3, width = 123, bg = button_color, fg = button_text, font = ("Helvetica", 12, "bold"), command = search)
imageicon4 = PhotoImage(file = "Images/layer 4.png")
Update_button = Button(root, image = imageicon4, bg = button_color, command = update)
Update_button.place(x = 110, y = 64)


def update_layout(event = None):
    if root.winfo_width() > 1:
        window_w = root.winfo_width()

        header_title.place_configure(relx = 0.5, y = 45, anchor = 'n')

        search_entry_width = 220
        search_x = max(900, min(window_w - 350, window_w - 350))
        search_entry.place_configure(x = search_x, y = 70, width = search_entry_width)
        srch.place_configure(x = search_x + search_entry_width + 10, y = 66)

        left_panel_w = min(900, max(560, int(window_w * 0.58)))
        left_panel_x = 30
        obj.place_configure(x = left_panel_x, y = 200, width = left_panel_w)
        obj2.place_configure(x = left_panel_x, y = 470, width = left_panel_w)

        right_column_x = max(700, min(window_w - 300, int(window_w * 0.72)))
        right_button_x = right_column_x + 18
        f.place_configure(x = right_column_x, y = 150)
        upload_button.place_configure(x = right_button_x, y = 370)
        saveButton.place_configure(x = right_button_x, y = 450)
        reset_button.place_configure(x = right_button_x, y = 530)
        exit_button.place_configure(x = right_button_x, y = 610)

root.bind('<Configure>', update_layout)

#registration and date
Label(root, text = "Registration No:", fg = framebg, bg = background, font = ("Helvetica", 13)).place(x = 30, y = 150)
Label(root, text = "Date:", fg = framebg, bg = background, font = ("Helvetica", 13)).place(x = 500, y = 150)

Registration = IntVar()
Date = StringVar()

reg_entry = Entry(root, textvariable = Registration, width = 15, font = ("Helvetica", 10), bd = 2)
reg_entry.place(x = 160, y=150)

registration_no()


today = date.today()
di = today.strftime("%d/%m/%Y")
date_entry = Entry(root, textvariable = Date, width = 15, font = ("Helvetica", 10) , bd = 2)
date_entry.place(x = 550, y=150)

Date.set(di)

#Student details
obj = LabelFrame(root, text = "Student's Details", width = 900, height = 250, bg = framebg, fg = framefg, font = ("Helvetica", 20), relief = GROOVE, bd = 2)
obj.place(x = 30, y = 200)

Label(obj, text = "Full Name:", font = ("Helvetica", 13), bg = framebg, fg = framefg).place(x = 30, y = 50)
Label(obj, text = "Date of Birth:", font = ("Helvetica", 13), bg = framebg, fg = framefg).place(x = 30, y = 100)
Label(obj, text = "Gender:", font = ("Helvetica", 13), bg = framebg, fg = framefg).place(x = 30, y = 150)

Label(obj, text = "Classification:", font = ("Helvetica", 13), bg = framebg, fg = framefg).place(x = 500, y = 50)
Label(obj, text = "Ethnicity:", font = ("Helvetica", 13), bg = framebg, fg = framefg).place(x = 500, y = 100)
Label(obj, text = "Major:", font = ("Helvetica", 13), bg = framebg, fg = framefg).place(x = 500, y = 150)

Name = StringVar()
name_entry = Entry(obj, textvariable = Name, width = 20, font = ("Helvetica", 10))
name_entry.place(x = 160, y = 50)

DOB = StringVar(value = "DD/MM/YYYY")
dob_entry = Entry(obj, textvariable = DOB, width = 20, font = ("Helvetica", 10), fg = '#8B8B8B')
dob_entry.place(x = 160, y = 100)

def on_dob_focus_in(event):
    if DOB.get() == "DD/MM/YYYY":
        DOB.set('')
        dob_entry.config(fg = '#000000')

def on_dob_focus_out(event):
    if DOB.get() == '':
        DOB.set('DD/MM/YYYY')
        dob_entry.config(fg = '#8B8B8B')


dob_entry.bind('<FocusIn>', on_dob_focus_in)
dob_entry.bind('<FocusOut>', on_dob_focus_out)

radio = IntVar()
R1 = Radiobutton(obj, text = "Male", variable = radio, value = 1, bg = framebg, fg = framefg, command = selection, selectcolor = "#FFFFFF")
R1.place(x = 150, y = 150)

R2 = Radiobutton(obj, text = "Female", variable = radio, value = 2, bg = framebg, fg = framefg, command = selection, selectcolor = "#FFFFFF")
R2.place(x = 200, y = 150)

Ethnicity = StringVar()
ethnicity_entry = Entry(obj, textvariable = Ethnicity, width = 20, font = ("Helvetica", 10))
ethnicity_entry.place(x = 630, y = 100)

Major = StringVar()
major_entry = Entry(obj, textvariable = Major, width = 20, font = ("Helvetica", 10))
major_entry.place(x = 630, y = 150)

Classification = Combobox(obj, value = ["Freshman", "Sophomore", "Junior", "Senior"], font = ("Roboto", 10), width = 17, state = "r")
Classification.place(x = 630, y = 50)
Classification.set("Select Classification")

#student secondary details
obj2 = LabelFrame(root, text = "More Student Details", width = 900, height = 220, bg = framebg, fg = framefg, font = ("Helvetica", 20), relief = GROOVE, bd = 2)
obj2.place(x = 30, y = 470)

Label(obj2, text = "Contact No:", font = ("Helvetica", 13), bg = framebg, fg = framefg).place(x = 30, y = 50)
Label(obj2, text = "Car Reg No:", font = ("Helvetica", 13), bg = framebg, fg = framefg).place(x = 30, y = 100)

ContactNo = StringVar()
contact_entry = Entry(obj2, textvariable = ContactNo, width = 20, font = ("Helvetica", 10))
contact_entry.place(x = 160, y = 50)

CarRegNo = StringVar()
car_reg_entry = Entry(obj2, textvariable = CarRegNo, width = 20, font = ("Helvetica", 10))
car_reg_entry.place(x = 160, y = 100)

Label(obj2, text = "Hostel Room No:", font = ("Helvetica", 13), bg = framebg, fg = framefg).place(x = 500, y = 50)
Label(obj2, text = "Occupation:", font = ("Helvetica", 13), bg = framebg, fg = framefg).place(x = 500, y = 100)

HostelRoomNo = StringVar()
hostel_room_entry = Entry(obj2, textvariable = HostelRoomNo, width = 20, font = ("Helvetica", 10))
hostel_room_entry.place(x = 630, y = 50)

Occupation = StringVar()
occupation_entry = Entry(obj2, textvariable = Occupation, width = 20, font = ("Helvetica", 10))
occupation_entry.place(x = 630, y = 100)

#Image
f = Frame(root, width = 200, height = 200, bg = framebg, relief = GROOVE, bd = 2)
img = PhotoImage(file = "Images/upload photo.png")
lbl = Label(f, image = img, bg = framebg)
lbl.place(x = 0, y = 0)

# Button
upload_button = Button(root, text = "Upload Photo", width = 19, height = 2, bg = button_color, fg = button_text, font = ("Helvetica", 10, "bold"), command = showimage)

saveButton = Button(root, text = "Save", width = 19, height = 2, bg = button_color, fg = button_text, font = ("Helvetica", 10, "bold"), command = save)
reset_button = Button(root, text = "Reset", width = 19, height = 2, bg = button_color, fg = button_text, font = ("Helvetica", 10, "bold"), command = clear)
exit_button = Button(root, text = "Exit", width = 19, height = 2, bg = button_color, fg = button_text, font = ("Helvetica", 10, "bold"), command = Exit)

f.place(x = 980, y = 150)
upload_button.place(x = 980, y = 370)
saveButton.place(x = 980, y = 450)
reset_button.place(x = 980, y = 530)
exit_button.place(x = 980, y = 610)

root.mainloop()
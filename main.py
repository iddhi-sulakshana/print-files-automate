import subprocess
from customtkinter import CTk, CTkButton, CTkLabel, CTkScrollableFrame, CTkComboBox, CTkFrame, CTk
import win32print
from tkinter import filedialog, messagebox

selected_files = []
printer_properties = None


def set_button_state(button, state):
    if state == "normal":
        button.configure(state="normal")
    elif state == "disabled":
        button.configure(state="disabled")


def display_files():
    i = 0
    for child in scrollable_frame.winfo_children():
        child.destroy()
    for file in selected_files:
        frame = CTkFrame(master=scrollable_frame, width=500, height=25)
        frame.pack(fill="x")

        label = CTkLabel(master=frame, text=file.split("/")[-1])
        label.pack(pady=2, side="left")

        delete_button = CTkButton(master=frame, text="Delete", command=lambda index=i: delete_callback(index))
        delete_button.pack(pady=2, side="right")
        i = i + 1
    root.update()


def delete_callback(file):
    scrollable_frame.winfo_children()[file].destroy()
    selected_files.pop(file)
    display_files()


def select_files(event):
    set_button_state(event, "disabled")
    global selected_files
    files = filedialog.askopenfilenames(title="Select Files", filetypes=())
    if len(files) > 0:
        selected_files = list(files)
        display_files()
    root.update()
    set_button_state(event, "normal")


def print_files(event):
    set_button_state(event, "disabled")
    if len(selected_files) == 0:
        messagebox.showwarning(title="File Select", message="No files selected to print")
        set_button_state(event, "normal")
        return
    printer_handle = win32print.OpenPrinter(printer_name)
    attributes = win32print.GetPrinter(printer_handle)[13]
    if (attributes & 0x00000400) >> 10 == 0:
        print("offline")
    else:
        print("online")
    for file_path in selected_files:
        print_file(file_path, printer_handle)
    win32print.ClosePrinter(printer_handle)
    messagebox.showinfo(title="Success", message="Printing")
    set_button_state(event, "normal")


def print_file(file_path, printer_handle):
    file_handle = open(file_path, 'rb')

    win32print.StartDocPrinter(printer_handle, 1, (file_path, None, "RAW"))

    win32print.StartPagePrinter(printer_handle)
    win32print.WritePrinter(printer_handle, file_handle.read())
    win32print.EndPagePrinter(printer_handle)
    win32print.EndDocPrinter(printer_handle)

    file_handle.close()


def get_defaults():
    global printer_name
    global printer_properties
    PRINTER_DEFAULTS = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
    pHandle = win32print.OpenPrinter(printer_name, PRINTER_DEFAULTS)
    printer_properties = win32print.GetPrinter(pHandle, 2)


def set_defaults():
    global printer_name
    global printer_properties
    PRINTER_DEFAULTS = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
    pHandle = win32print.OpenPrinter(printer_name, PRINTER_DEFAULTS)
    win32print.SetPrinter(pHandle, 2, printer_properties, 0)


def clear_files(event):
    set_button_state(event, "disabled")
    global selected_files
    selected_files = []
    for child in scrollable_frame.winfo_children():
        child.destroy()
    set_defaults()
    root.update()
    set_button_state(event, "normal")


def get_printers():
    global printer_name
    printers = []
    for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL):
        printers.append(printer[2])
    printer_name = printers[0]
    get_defaults()
    return printers


def combobox_callback(value):
    global printer_name
    set_defaults()
    printer_name = value
    get_defaults()


def open_settings():
    global printer_name
    subprocess.run(f"RUNDLL32 PRINTUI.DLL,PrintUIEntry /e /n \"{printer_name}\"", shell=True)


def window_init():
    global root
    root = CTk()
    root.title("File Printing")

    printers = get_printers()
    combobox = CTkComboBox(master=root, values=printers, command=combobox_callback)
    combobox.pack(padx=10, pady=5, fill="x")

    frame1 = CTkFrame(master=root, width=500, height=25)
    frame1.pack(fill="x")

    select_button = CTkButton(frame1, text="Select Files", command=lambda: select_files(select_button))
    select_button.pack(side="left", pady=5, padx=10)

    clear_button = CTkButton(frame1, text="New Job", command=lambda: clear_files(clear_button))
    clear_button.pack(side="right", pady=5, padx=10)

    settings_button = CTkButton(root, text="Settings", command=open_settings)
    settings_button.pack(padx=10, pady=5, fill="x")

    global scrollable_frame
    scrollable_frame = CTkScrollableFrame(master=root, width=500, height=300)
    scrollable_frame.pack(padx=10, pady=5, fill="both")

    print_button = CTkButton(root, text="Print Files", command=lambda: print_files(print_button))
    print_button.pack(padx=10, pady=5, fill="x")

    root.mainloop()


window_init()
set_defaults()

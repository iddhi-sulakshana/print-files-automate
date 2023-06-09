import subprocess
from customtkinter import CTk, CTkButton, CTkLabel, CTkScrollableFrame, CTkComboBox, CTkFrame
import win32print
from tkinter import filedialog, messagebox


class FilePrinterApp:
    def __init__(self):
        self.selected_files = []
        self.printer_properties = None
        self.printer_name = None

        self.root = CTk()
        self.root.title("File Printing")

        self.printers = self.get_printers()
        self.combobox = CTkComboBox(master=self.root, values=self.printers, command=self.combobox_callback)
        self.combobox.pack(padx=10, pady=5, fill="x")

        frame1 = CTkFrame(master=self.root, width=500, height=25)
        frame1.pack(fill="x")

        self.select_button = CTkButton(frame1, text="Select Files", command=self.select_files)
        self.select_button.pack(side="left", pady=5, padx=10)

        clear_button = CTkButton(frame1, text="New Job", command=self.clear_files)
        clear_button.pack(side="right", pady=5, padx=10)

        settings_button = CTkButton(self.root, text="Settings", command=self.open_settings)
        settings_button.pack(padx=10, pady=5, fill="x")

        self.scrollable_frame = CTkScrollableFrame(master=self.root, width=500, height=300)
        self.scrollable_frame.pack(padx=10, pady=5, fill="both")

        self.print_button = CTkButton(self.root, text="Print Files", command=self.print_files)
        self.print_button.pack(padx=10, pady=5, fill="x")

        self.root.mainloop()

    def combobox_callback(self, value):
        self.set_defaults()
        self.printer_name = value
        self.get_defaults()

    def get_printers(self):
        printers = []
        for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL):
            printers.append(printer[2])
        self.printer_name = printers[0]
        self.get_defaults()
        return printers

    def get_defaults(self):
        printer_defaults = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
        p_handle = win32print.OpenPrinter(self.printer_name, printer_defaults)
        self.printer_properties = win32print.GetPrinter(p_handle, 2)

    def set_defaults(self):
        printer_defaults = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
        p_handle = win32print.OpenPrinter(self.printer_name, printer_defaults)
        win32print.SetPrinter(p_handle, 2, self.printer_properties, 0)

    def select_files(self):
        self.select_button.configure(state="disabled")
        files = filedialog.askopenfilenames(title="Select Files", filetypes=())
        if len(files) > 0:
            self.selected_files = list(files)
            self.display_files()
        self.root.update()
        self.select_button.configure(state="normal")

    def print_files(self):
        self.print_button.configure(state="disabled")
        if len(self.selected_files) == 0:
            messagebox.showwarning(title="File Select", message="No files selected to print")
            self.print_button.configure(state="normal")
            return
        printer_handle = win32print.OpenPrinter(self.printer_name)
        attributes = win32print.GetPrinter(printer_handle)[13]
        if (attributes & 0x00000400) >> 10 == 0:
            print("offline")
        else:
            print("online")
        for file_path in self.selected_files:
            self.print_file(file_path, printer_handle)
        win32print.ClosePrinter(printer_handle)
        self.clear_files()
        messagebox.showinfo(title="Success", message="Printing")
        self.print_button.configure(state="normal")

    @staticmethod
    def print_file(file_path, printer_handle):
        file_handle = open(file_path, 'rb')

        win32print.StartDocPrinter(printer_handle, 1, (file_path, None, "RAW"))

        win32print.StartPagePrinter(printer_handle)
        win32print.WritePrinter(printer_handle, file_handle.read())
        win32print.EndPagePrinter(printer_handle)
        win32print.EndDocPrinter(printer_handle)

        file_handle.close()

    def display_files(self):
        i = 0
        for child in self.scrollable_frame.winfo_children():
            child.destroy()
        for file in self.selected_files:
            frame = CTkFrame(master=self.scrollable_frame, width=500, height=25)
            frame.pack(fill="x")

            label = CTkLabel(master=frame, text=file.split("/")[-1])
            label.pack(pady=2, side="left")

            delete_button = CTkButton(master=frame, text="Delete", command=lambda index=i: self.delete_callback(index))
            delete_button.pack(pady=2, side="right")
            i = i + 1
        self.root.update()

    def delete_callback(self, file):
        self.scrollable_frame.winfo_children()[file].destroy()
        self.selected_files.pop(file)
        self.display_files()

    def open_settings(self):
        subprocess.run(f"RUNDLL32 PRINTUI.DLL,PrintUIEntry /e /n \"{self.printer_name}\"", shell=True)

    def clear_files(self):
        self.select_button.configure(state="disabled")
        self.selected_files = []
        for child in self.scrollable_frame.winfo_children():
            child.destroy()
        self.set_defaults()
        self.root.update()
        self.select_button.configure(state="normal")


if __name__ == "__main__":
    app = FilePrinterApp()

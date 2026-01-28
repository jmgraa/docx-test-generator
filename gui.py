from datetime import date
from tkinter import filedialog

from customtkinter import CTk, CTkButton, CTkEntry, CTkLabel, CTkToplevel

from setup import mix_tests


class Popup(CTkToplevel):
    def __init__(self, title, description):
        super().__init__()
        self.geometry("300x150")
        self.title(title)
        self.attributes("-topmost", True)
        self.label = CTkLabel(self, text=description)
        self.label.pack(pady=20)
        self.close_btn = CTkButton(self, text="Zamknij", command=self.destroy)
        self.close_btn.pack(pady=10)


class App(CTk):
    def __init__(self):
        super().__init__()
        self.popup_window = None
        self.title("Generator testów")
        self.geometry("400x500")

        self.label_num = CTkLabel(self, text="Liczba kopii:")
        self.label_num.pack(pady=(20, 0))
        self.entry_num = CTkEntry(self, placeholder_text="Liczba...")
        self.entry_num.pack(pady=5)

        self.label_text = CTkLabel(self, text="Nagłówek testu:")
        self.label_text.pack(pady=(10, 0))
        self.entry_text = CTkEntry(
            self,
            placeholder_text="Nagłówek...",
        )
        self.entry_text.pack(pady=5)

        self.date_label = CTkLabel(self, text="Data testu (RRRR-MM-DD):")
        self.date_label.pack(pady=(10, 0))
        self.date_text = CTkEntry(self)
        self.date_text.pack(pady=5)
        self.date_text.insert(0, date.today().isoformat())

        self.file_path = ""
        self.btn_file = CTkButton(
            self, text="Wybierz plik Word", command=self.pick_file
        )
        self.btn_file.pack(pady=(20, 5))
        self.label_file = CTkLabel(self, text="Nie wybrano pliku", font=("Arial", 10))
        self.label_file.pack()

        self.dir_path = ""
        self.btn_dir = CTkButton(
            self, text="Wybierz folder wyjściowy", command=self.pick_dir
        )
        self.btn_dir.pack(pady=(20, 5))
        self.label_dir = CTkLabel(self, text="Nie wybrano folderu", font=("Arial", 10))
        self.label_dir.pack()

        self.btn_submit = CTkButton(
            self, text="Generuj pliki", fg_color="green", command=self.submit
        )
        self.btn_submit.pack(pady=20)

    def pick_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Plik Word", "*.docx")])
        if self.file_path:
            self.label_file.configure(text=f"Plik: {self.file_path.split('/')[-1]}")

    def pick_dir(self):
        self.dir_path = filedialog.askdirectory()
        if self.dir_path:
            self.label_dir.configure(text=f"Folder: {self.dir_path}")

    def submit(self):
        data = [
            self.file_path,
            self.entry_text.get(),
            "2025-02-06",
            self.entry_num.get(),
            self.dir_path,
        ]
        try:
            mix_tests(data)
            self.__open_popup("Sukces", "Pliki zostały wygenerowane!")
        except Exception as e:
            self.__open_popup("Błąd", str(e))

    def __open_popup(self, title, description):
        if self.popup_window is None or not self.popup_window.winfo_exists():
            self.popup_window = Popup(title, description)
        else:
            self.popup_window.focus()


if __name__ == "__main__":
    app = App()
    app.mainloop()

from datetime import date
from tkinter import filedialog

from customtkinter import (
    CTk,
    CTkButton,
    CTkEntry,
    CTkLabel,
    CTkTextbox,
    CTkToplevel,
    set_widget_scaling,
)

from setup import mix_tests
from validating import validate_input

DEFAULT_INTRO = "PRZED PRZYSTĄPIENIEM DO WYPEŁNIANIA FORMULARZA ODPOWIEDZI PROSZĘ PRZECZYTAĆ INSTRUKCJĘ. PROSZĘ PAMIĘTAĆ O CZYTELNYM PODPISANIU FORMULARZA ODPOWIEDZI W PRAWYM-DOLNYM ROGU.\nPROSZĘ PAMIĘTAĆ O POPRAWNYM OZNACZENIU NUMERU TESTU I NUMERU IDENTYFIKACYJNEGO STUDENTA"
LABEL_FONT = ("Arial", 20, "bold")

set_widget_scaling(1.05)


class InputParams:
    def __init__(
        self,
        file_path,
        title,
        date,
        copy_number,
        margin_v,
        margin_h,
        output_dir,
        intro=DEFAULT_INTRO,
    ):
        self.file_path = file_path
        self.title = title
        self.date = date
        self.copy_number = copy_number
        self.margin_v = margin_v
        self.margin_h = margin_h
        self.output_dir = output_dir
        self.intro = intro

    def validate_input(self):
        validation = validate_input(
            self.file_path,
            self.date,
            self.copy_number,
            self.output_dir,
            self.margin_v,
            self.margin_h,
        )

        errors = "\n".join(e for e in validation if isinstance(e, str))

        if not errors:
            self.__format_input()

        return errors

    def __format_input(self):
        self.file_path = self.file_path.strip()
        self.title = self.title.strip()
        self.date = self.date.strip()
        if isinstance(self.copy_number, str):
            self.copy_number = int(self.copy_number.strip())
        if isinstance(self.margin_v, str):
            self.margin_v = float(self.margin_v.strip())
        if isinstance(self.margin_h, str):
            self.margin_h = float(self.margin_h.strip())
        self.output_dir = self.output_dir.strip()


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


class MarginsPopup(CTkToplevel):
    def __init__(self, parent_app):
        super().__init__()
        self.parent_app = parent_app
        self.geometry("320x280")
        self.title("Ustawienia strony")
        self.attributes("-topmost", True)

        self.label_margin_v = CTkLabel(
            self, text="Margines pionowy (cm)", font=LABEL_FONT
        )
        self.label_margin_v.pack(pady=(10, 4))
        self.entry_margin_v = CTkEntry(self, placeholder_text="np. 1.5", width=200)
        self.entry_margin_v.pack(pady=(0, 8))
        self.entry_margin_v.insert(0, getattr(parent_app, "margin_v", "1"))

        self.label_margin_h = CTkLabel(
            self, text="Margines poziomy (cm)", font=LABEL_FONT
        )
        self.label_margin_h.pack(pady=(20, 4))
        self.entry_margin_h = CTkEntry(self, placeholder_text="np. 1.5", width=200)
        self.entry_margin_h.pack(pady=(0, 8))
        self.entry_margin_h.insert(0, getattr(parent_app, "margin_h", "1"))

        self.btn_ok = CTkButton(self, text="Zapisz", command=self._save_and_close)
        self.btn_ok.pack(pady=16)

    def _save_and_close(self):
        self.parent_app.margin_h = self.entry_margin_h.get().strip() or "1"
        self.parent_app.margin_v = self.entry_margin_v.get().strip() or "1"
        self.destroy()


class App(CTk):
    def __init__(self):
        super().__init__()
        self.popup_window = None
        self.margin_h = "1"
        self.margin_v = "1"
        self.title("Generator testów")
        self.geometry("700x850")

        self.label_num = CTkLabel(self, text="Liczba kopii", font=LABEL_FONT)
        self.label_num.pack(pady=(20, 0))
        self.entry_num = CTkEntry(self, placeholder_text="Liczba...")
        self.entry_num.pack(pady=8)

        self.label_text = CTkLabel(self, text="Nagłówek testu", font=LABEL_FONT)
        self.label_text.pack(pady=(10, 0))
        self.entry_text = CTkEntry(
            self,
            placeholder_text="Nagłówek...",
        )
        self.entry_text.pack(pady=8)

        self.date_label = CTkLabel(
            self, text="Data testu (RRRR-MM-DD)", font=LABEL_FONT
        )
        self.date_label.pack(pady=(10, 0))
        self.date_text = CTkEntry(self)
        self.date_text.pack(pady=8)
        self.date_text.insert(0, date.today().isoformat())

        self.textbox_text = CTkLabel(self, text="Wstęp do testu", font=LABEL_FONT)
        self.textbox_text.pack(pady=(10, 0))
        self.textbox = CTkTextbox(self, width=500, height=200)
        self.textbox.pack(padx=20, pady=20)
        self.textbox.insert("0.0", DEFAULT_INTRO)

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

        self.btn_margins = CTkButton(
            self, text="Ustawienia marginesów", command=self._open_margins_popup
        )
        self.btn_margins.pack(pady=(10, 5))

        self.btn_submit = CTkButton(
            self, text="Generuj pliki", fg_color="green", command=self.submit
        )
        self.btn_submit.pack(pady=20)

        self.progrss_bar = CTkLabel(self, text="", font=("Arial", 10))
        self.progrss_bar.pack()

    def pick_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Plik Word", "*.docx")])
        if self.file_path:
            self.label_file.configure(text=f"Plik: {self.file_path.split('/')[-1]}")

    def pick_dir(self):
        self.dir_path = filedialog.askdirectory()
        if self.dir_path:
            self.label_dir.configure(text=f"Folder: {self.dir_path}")

    def submit(self):
        input_params = InputParams(
            self.file_path,
            self.entry_text.get(),
            self.date_text.get(),
            self.entry_num.get(),
            self.margin_v,
            self.margin_h,
            self.dir_path,
            self.textbox.get("1.0", "end-1c"),
        )

        errors = input_params.validate_input()
        if not errors:
            self.__run_generation(input_params)
        else:
            self.__open_info_popup("Błąd", errors)

    def __run_generation(self, data):
        try:
            for i in mix_tests(data):
                self.progrss_bar.configure(text=f"Postęp: {i}/{data.copy_number}")
                self.update_idletasks()
            self.__open_info_popup("Sukces", "Pliki zostały wygenerowane!")
        except Exception as e:
            self.__open_info_popup("Błąd", str(e))

    def __open_info_popup(self, title, description):
        if self.popup_window is None or not self.popup_window.winfo_exists():
            self.popup_window = Popup(title, description)
        else:
            self.popup_window.focus()

    def _open_margins_popup(self):
        self.popup_window = MarginsPopup(self)


if __name__ == "__main__":
    app = App()
    app.mainloop()

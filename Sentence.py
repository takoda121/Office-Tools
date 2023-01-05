import tkinter
import tkinter.messagebox
import customtkinter
import docx
from tkinter import filedialog

customtkinter.set_appearance_mode("System") 
customtkinter.set_default_color_theme("green") 


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("Sentence")
        self.geometry(f"{1100}x{580}")
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)
        self.logo_label = customtkinter.CTkLabel(self, text="Sentence", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(2, 0))
        self.sidebar_button_2 = customtkinter.CTkButton(self, command=self.savelol, text="Save File")
        self.sidebar_button_2.grid(row=3, column=0, padx=20, pady=10)
        self.entry = customtkinter.CTkEntry(self, placeholder_text="Title")
        self.entry.grid(row=3, column=1, columnspan=2, padx=(20, 0), pady=(20, 20), sticky="nsew")
        self.textbox = customtkinter.CTkTextbox(self, width=250, height=500)
        self.textbox.grid(row=0, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")
        self.textbox.insert("0.0", "Your text")



    def savelol(self):
        file_path = filedialog.asksaveasfilename(filetypes=[('Word Documents', '*.docx')])
        print(file_path)
        open(file_path, 'w')
        doc = docx.Document()
        nigger = self.entry.get()
        doc.add_heading(nigger, 0)
        doc.add_paragraph(self.textbox.get(1.0, "end"))
        if not file_path.endswith(".docx"):
            nothing = file_path
            file_path = nothing + ".docx"
        doc.save(file_path)


if __name__ == "__main__":
    print("Welcome to sentence!\nLightweight tool\nSaves to .docx")
    app = App()
    app.mainloop()

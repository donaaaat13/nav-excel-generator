import customtkinter
from tkinter import filedialog
from read_nav_document import *

customtkinter.set_appearance_mode("System")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("NAV excel generátor")
        self.geometry("400x240")
        self.grid_rowconfigure(0, weight=1)  # configure grid system
        self.grid_columnconfigure(0, weight=1)

        self.button = customtkinter.CTkButton(master=self, text="Fájlok kiválasztása", command=self.button_function)
        self.button.grid(row=0, column=0, sticky="nsew")
        self.textbox = customtkinter.CTkTextbox(master=self, width=400, corner_radius=0)
        self.textbox.grid(row=1, column=0, sticky="nsew")

    def button_function(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        logger = logging.getLogger()

        invoice_data = []
        # Create an instance of the PDFDataExtractor class
        for file in files:
            self.textbox.insert("end", f"PDF beolvasva:\n{file}\n\n")
            pdf_extractor = PDFDataExtractor(file, logger)

            # Use the methods to extract information
            invoice_data.append(pdf_extractor.get_data())
        
        invoice_data_df = pd.DataFrame(invoice_data)
        print(invoice_data_df)
        # Get the current date and time
        current_datetime = datetime.now()
        # Format the date and time as a string (YYYY-MM-DD_HH-MM-SS)
        formatted_datetime = current_datetime.strftime("%Y-%m-%d_%H-%M-%S")
        # writing to Excel
        excel_file_path = f'szamlak_{formatted_datetime}.xlsx'

        writer = pd.ExcelWriter(excel_file_path, engine='openpyxl')
        invoice_data_df.to_excel(writer, index=False)  # send df to writer
        writer.close()
        self.textbox.insert("end", f"EXCEL LEGENERÁLVA!\n{excel_file_path}")


app =  App()
app.mainloop()
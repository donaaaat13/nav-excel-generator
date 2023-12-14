# importing required modules
import os
import PyPDF2
import pandas as pd
from datetime import datetime

class PDFDataExtractor:
    def __init__(self, file_path):
        self.file_path = file_path
        self.lines = self.extract_lines()

    def extract_lines(self):
        with open(self.file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ''
            for page in pdf_reader.pages:
                text += page.extract_text()
            return text.split('\n')

    def __extract_vevo(self):
        for idx, line in enumerate(self.lines):
            if "VEVŐ:" in line:
                return self.lines[idx + 1]

    def __extract_kiallitas_datuma(self):
        for line in self.lines:
            if "Kiállítás dátuma: " in line:
                return line.replace("Kiállítás dátuma: ", "")

    def __extract_sorszam(self):
        for line in self.lines:
            if "Sorszám: " in line:
                return line.replace("Sorszám: ", "")

    def __extract_osszeg(self):
        for idx, line in enumerate(self.lines):
            if line == "Összesen:":
                return self.lines[idx + 1]
            
    def get_data(self):
        data = {
            "Vevő": self.__extract_vevo(),
            "Számla sorszáma": self.__extract_sorszam(),
            "Kiállítás dátuma": self.__extract_kiallitas_datuma(),
            "Összeg": self.__extract_osszeg(),
        }
        return data

if __name__ == "__main__":
    directory = "../sample_files"
    files = os.listdir(directory)

    invoice_data = []
    # Create an instance of the PDFDataExtractor class
    for file in files:
        pdf_extractor = PDFDataExtractor(f"{directory}/{file}")

        # Use the methods to extract information
        invoice_data.append(pdf_extractor.get_data())
    
    invoice_data_df = pd.DataFrame(invoice_data)
    print(invoice_data_df)

    # Get the current date and time
    current_datetime = datetime.now()
    # Format the date and time as a string (YYYY-MM-DD_HH-MM-SS)
    formatted_datetime = current_datetime.strftime("%Y-%m-%d_%H-%M-%S")
    # writing to Excel
    excel_file_path = f'../excels/szamlak_{formatted_datetime}.xlsx'

    writer = pd.ExcelWriter(excel_file_path, engine='openpyxl')
    invoice_data_df.to_excel(writer, index=False)  # send df to writer
    writer.close()

    #invoice_data_df.to_excel(excel_file_path, index=False, auto_width=True)
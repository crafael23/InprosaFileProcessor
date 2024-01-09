import PyPDF2
import pdfplumber
import pandas as pd
import camelot
import os


def split_pdf(archivo: str):
    pdf_file = open(archivo, "rb")
    pdf_reader = PyPDF2.PdfFileReader(pdf_file)

    for page in range(len(pdf_reader.pages)):
        pdf_writer = PyPDF2.PdfFileWriter()
        pdf_writer.addPage(pdf_reader.pages[page])

        output_filename = f"output/Temp/{archivo.split('.pdf')[0]}_page_{page + 1}.pdf"

        with open(output_filename, "wb") as output_pdf:
            pdf_writer.write(output_pdf)


def pdf_to_xlsx(pdf_file_name, xlsx_file_name):
    print("Converting: ", pdf_file_name)
    tables = camelot.read_pdf(pdf_file_name, pages="all", flavor="stream")

    for i, table in enumerate(tables):  # type: ignore
        df = table.df
        rows_to_delete = []

        if "\nActividad:" in df.iloc[1, 1]:
            # Significa que la actividad y el codigo de actividad estan juntos y hay que separarlos

            actividad = df.iloc[1, 1].split("\n")[1]
            codigo_actividad = df.iloc[1, 1].split("\n")[0]
            nombre_ficha = df.iloc[1, 2]
            ##Replace the values in the dataframe on the 1,1  1,2 and 1,3 positions
            df.iloc[1, 1] = codigo_actividad
            df.iloc[1, 2] = actividad
            df.iloc[1, 3] = nombre_ficha

        ##Revisa si la fila esta vacia y solo tiene un valor en el nombre del insumo, si es asi, lo agrega al nombree dle insumo anterior y elimina la fila que no tenia nada mas que eso

        try:
            for index in range(len(df)):
                if (
                    (df.iloc[index, 0] == "")
                    & (df.iloc[index, 2] == "")
                    & (df.iloc[index, 3] == "")
                    & (df.iloc[index, 4] == "")
                    & (df.iloc[index, 5] == "")
                    & (df.iloc[index, 6] == "")
                ):
                    text = df.iloc[index, 1]
                    df.iloc[index - 1, 1] = df.iloc[index - 1, 1] + " " + text
                    rows_to_delete.append(index)
        except Exception as e:
            print("Error en el archivo: ", pdf_file_name)
            print("Error: ", e)

            print(df)
            return

        df = df.drop(rows_to_delete)

        df.to_excel(xlsx_file_name, index=False)


def main():
    path = "Fichas1.pdf"

    # split_pdf(path)

    files = os.listdir("output/Temp")
    pdf_files = [file for file in files if file.endswith(".pdf")]

    print(f"Total number of PDF files: {len(pdf_files)}")
    

    for i in range(1, len(pdf_files) + 1):
        # Your code here
        # Replace '...' with the code you want to execute for each iteration
        pdf_filename= f"output/Temp/Fichas1_page_{i}.pdf"
        xlsx_output_filename = f"output/Xlsx/Fichas1_page_{i}.xlsx"
        print(f"Converting {pdf_filename} to {xlsx_output_filename}")
        pdf_to_xlsx(pdf_filename, xlsx_output_filename)
        
        

    # for filename in os.listdir("output/Temp"):
    #     if filename.endswith(".pdf"):
    #         os.remove(f"output/Temp/{filename}")

    # for filename in os.listdir("output"):
    #     if filename.endswith(".xlsx"):
    #         os.remove(f"output/{filename}")

    print("Done :) ")
    pass


if __name__ == "__main__":
    main()

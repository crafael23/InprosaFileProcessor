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

        linea1 = df.iloc[0].values
        linea2 = df.iloc[1].values
        linea3 = df.iloc[2].values

        if "Dirección de Proyectos" in linea1:
            rows_to_delete.append(0)
        if "Unidad de Costos" in linea2:
            rows_to_delete.append(1)
        if "Reporte De Fichas" in linea3:
            rows_to_delete.append(2)
        if "Reporte De Fichas" in linea1:
            rows_to_delete.append(0)

        df = df.drop(rows_to_delete)
        df = df.reset_index(drop=True)

        rows_to_delete = []

        if "\nActividad:" in df.iloc[0, 1]:
            df.iloc[0, 1] = df.iloc[0, 1].replace("\nActividad:", "")

        if "Actividad:" in df.iloc[0, 2]:
            df.iloc[0, 2] = df.iloc[0, 2].replace("Actividad:", "")
            NombreFichaFinal = df.iloc[0, 3]
            print("\033[94mNombre de ficha esta en la posicion 3\033[0m")
            # This print statement should display the message "Nombre de ficha esta en la posicion 3" in the console with the color blue (94).
            print("Nombre Ficha Final =", NombreFichaFinal)
        elif df.iloc[0, 2] == "":
            NombreFichaFinal = df.iloc[0, 3]
            # This print statement should display the message "Nombre de ficha esta en la posicion 3" in the console with the color yellow (93).
            print("\033[93mNombre de ficha esta en la posicion 3\033[0m")
            print("Nombre Ficha Final =", NombreFichaFinal)
        elif df.iloc[0, 2] != "":
            NombreFichaFinal = df.iloc[0, 2]
            # This print statement should display the message "Nombre de ficha esta en la posicion 2" in the console with the color red (91).
            print("Nombre Ficha Final =", NombreFichaFinal)
        else:
            print("\033[91mError en el archivo: \033[0m", pdf_file_name)
            print("Error: ", e)
            print(df)
            return

        valor_CodigoActividad = df.iloc[0, 1]
        print("Codigo Actividad =", valor_CodigoActividad)

        for col in df.columns:
            df.iloc[0, col] = ""

        df.iloc[0, 0] = valor_CodigoActividad
        df.iloc[0, 1] = NombreFichaFinal

        # print(df.head(4))

        ##Revisa si la fila esta vacia y solo tiene un valor en el nombre del insumo, si es asi, lo agrega al nombree dle insumo anterior y elimina la fila que no tenia nada mas que eso

        try:
            for index in range(len(df)):
                if (
                    (df.iloc[index, 0] == "")
                    & (df.iloc[index, 2] == "")
                    & (df.iloc[index, 3] == "")
                    & (df.iloc[index, 4] == "")
                    & (df.iloc[index, 5] == "")
                ):
                    print("Condicion vacia cumplida")
                    text = df.iloc[index, 1]
                    df.iloc[index - 1, 1] = df.iloc[index - 1, 1] + " " + text
                    rows_to_delete.append(index)
        except Exception as e:
            print("Error en el archivo: ", pdf_file_name)
            print("Error: ", e)

            print(df)
            return

        df = df.drop(rows_to_delete)

        print("\n\n")

        df.to_excel(xlsx_file_name, index=False)


def main():
    path = "Fichas1.pdf"

    split_pdf(path)

    files = os.listdir("output/Temp")
    pdf_files = [file for file in files if file.endswith(".pdf")]

    print(f"Total number of PDF files: {len(pdf_files)}")

    for i in range(1, len(pdf_files) + 1):
        # for i in range(1, 200):

        pdf_filename = f"output/Temp/Fichas1_page_{i}.pdf"
        xlsx_output_filename = f"output/Xlsx/Fichas1_page_{i}.xlsx"
        print("\n\n")

        print(f"Converting {pdf_filename} to {xlsx_output_filename}")

        pdf_to_xlsx(pdf_filename, xlsx_output_filename)

    for filename in os.listdir("output/Temp"):
        if filename.endswith(".pdf"):
            os.remove(f"output/Temp/{filename}")

    # for filename in os.listdir("output/Xlsx"):
    #     if filename.endswith(".xlsx"):
    #         os.remove(f"output/Xlsx/{filename}")

    print("Done :) ")
    pass


if __name__ == "__main__":
    main()

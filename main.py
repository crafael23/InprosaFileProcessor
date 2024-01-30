import PyPDF2
import pandas as pd
import camelot
import os
import concurrent.futures


def split_pdf(archivo: str):
    with open(archivo, "rb") as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)

        for page in range(len(pdf_reader.pages)):
            pdf_writer = PyPDF2.PdfWriter()
            pdf_writer.add_page(pdf_reader.pages[page])

            output_filename = (
                f"output/Temp/{archivo.split('.pdf')[0]}_page_{page + 1}.pdf"
            )

            with open(output_filename, "wb") as output_pdf:
                pdf_writer.write(output_pdf)


def pdf_to_dataFrame(pdf_file_name):
    tables = camelot.read_pdf(pdf_file_name, pages="all", flavor="stream")

    for i, table in enumerate(tables):  # type: ignore
        df = table.df
        df = df.iloc[:, :5]
        rows_to_delete = []
        try:
            # Agarra los valores de las primeras 3 filas y los guarda en variables
            linea1 = df.iloc[0].values
            linea2 = df.iloc[1].values
            linea3 = df.iloc[2].values

            # Revisa si las primeras 3 filas tienen los valores que se esperan, si no, los agrega a la lista de filas a eliminar
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

            # Revisa si en el espacio 0,1 existe \nActividad: y lo elimina si es asi.
            if "\nActividad:" in df.iloc[0, 1]:
                df.iloc[0, 1] = df.iloc[0, 1].replace("\nActividad:", "")

            # Revisa si en el espacio 0,2 existe Actividad: y lo elimina si es asi. Luego guarda el valor de la celda en la variable
            # NombreFichaFinal y lo reemplaza adonde tiene que estar
            if "Actividad:" in df.iloc[0, 2]:
                df.iloc[0, 2] = df.iloc[0, 2].replace("Actividad:", "")
                NombreFichaFinal = df.iloc[0, 3]
                # print("\033[94mNombre de ficha esta en la posicion 3\033[0m")
                # This print statement should display the message "Nombre de ficha esta en la posicion 3" in the console with the color blue (94).
                # print("Nombre Ficha Final =", NombreFichaFinal)
            elif df.iloc[0, 2] == "":
                NombreFichaFinal = df.iloc[0, 3]
                # This print statement should display the message "Nombre de ficha esta en la posicion 3" in the console with the color yellow (93).
                # print("\033[93mNombre de ficha esta en la posicion 3\033[0m")
                # print("Nombre Ficha Final =", NombreFichaFinal)
            elif df.iloc[0, 2] != "":
                NombreFichaFinal = df.iloc[0, 2]
                # This print statement should display the message "Nombre de ficha esta en la posicion 2" in the console with the color red (91).
                # print("Nombre Ficha Final =", NombreFichaFinal)
            else:
                print("\033[91mError en el archivo: \033[0m", pdf_file_name)
                print("Error: ", e)
                print(df)
                return
            valor_CodigoActividad = df.iloc[0, 1]
            # print("Codigo Actividad =", valor_CodigoActividad)

            # limpiar la primera fila
            for col in df.columns:
                df.iloc[0, col] = ""

            # Reemplazar los valores de la primera fila con los valores que se guardaron en las variables
            df.iloc[0, 0] = valor_CodigoActividad
            df.iloc[0, 1] = NombreFichaFinal

            # Esto revisa si hay algun valor del nombre de la ficha que fue truncated hacia otra celda just abajo de la primera fila
            # De ser asi lo agrega al nombre de la ficha y elimina la fila que tenia el valor truncado
            dato: str = ""
            if (df.iloc[1, 0] == "") & (df.iloc[1, 1] == ""):
                if df.iloc[1, 2] == "":
                    dato = df.iloc[1, 3]
                else:
                    dato = df.iloc[1, 2]
                df.iloc[0, 1] = df.iloc[0, 1] + " " + dato
                df.drop(1, inplace=True)
                df.reset_index(drop=True, inplace=True)

            df.iloc[0, 2] = df.iloc[1, 1]
            df.iloc[0, 3] = 0
            df.iloc[0, 4] = 0

            tipodeinsumo = df.iloc[2, 1]

            vocabulario_excepciones = [
                valor_CodigoActividad,
                "Unidad:",
                "Tipo De Insumo:",
                "Codigo_Insumo",
                "",
            ]
            # Revisa si la fila esta vacia y solo tiene un valor en el nombre del insumo, si es asi, lo agrega al nombree dle insumo anterior y elimina la fila que no tenia nada mas que eso
            last_column_number = len(df.columns)

            vocabulario_borrar = [
                "Tipo De Insumo:",
                "Codigo_Insumo",
                "LICITACIONES\nUsuario:",
            ]

            df[str(last_column_number)] = ""
            df.iloc[0, last_column_number] = "+"
            for index in range(len(df)):
                if df.iloc[index, 0] == "Tipo De Insumo:":
                    tipodeinsumo = df.iloc[index, 1]
                if (
                    (df.iloc[index, 0] == "")
                    & (df.iloc[index, 2] == "")
                    & (df.iloc[index, 3] == "")
                    & (df.iloc[index, 4] == "")
                ):
                    # print("Condicion vacia cumplida")
                    text = df.iloc[index, 1]
                    df.iloc[index - 1, 1] = df.iloc[index - 1, 1] + " " + text
                    rows_to_delete.append(index)

                if df.iloc[index, 0] not in vocabulario_excepciones:
                    df.iloc[index, last_column_number] = tipodeinsumo

                if df.iloc[index, 0] in vocabulario_borrar:
                    rows_to_delete.append(index)

        except Exception as e:
            print("Error en el archivo: ", pdf_file_name)
            print("Error: ", e)
            print(df)
            return

        rows_to_delete.append(1)
        df = df.drop(rows_to_delete)
        df.reset_index(drop=True, inplace=True)



        return df


def merge(path: str, output: str):
    print("Merging files")
    files = os.listdir(path)
    xlsx_files = [file for file in files if file.endswith(".xlsx")]
    xlsx_files.sort()  # Sort the list alphabetically

    # List to hold all the individual DataFrames
    df_list = []

    # Read each Excel file into a DataFrame and add it to the list
    for file in xlsx_files:
        df = pd.read_excel(os.path.join(path, file))
        df_list.append(df)

        empty_df = pd.DataFrame([[""] * len(df.columns)] * 2, columns=df.columns)
        df_list.append(empty_df)
    # Concatenate all the DataFrames together
    merged_df = pd.concat(df_list, ignore_index=True)

    # Write the merged DataFrame to a new Excel file
    merged_df.to_excel(output, index=False)


def process_file(file):
    dataframe = pdf_to_dataFrame(f"output/Temp/{file}")
    if dataframe is not None:
        empty_df = pd.DataFrame(
            [[""] * 5] * 2, columns=[f"Column{i+1}" for i in range(5)]
        )
        return dataframe, empty_df
    return None, None


def main():
    path = "Fichas_final.pdf"
    path2 = "Fichas1.pdf"

    split_pdf(path)
    split_pdf(path2)

    files = os.listdir("output/Temp")
    pdf_files = [file for file in files if file.endswith(".pdf")]

    df_list = []

    print("Converting pdf to dataframe")
    with concurrent.futures.ThreadPoolExecutor(max_workers=16) as executor:
        futures = [executor.submit(process_file, file) for file in pdf_files]
        for future in concurrent.futures.as_completed(futures):
            dataframe, empty_df = future.result()
            if dataframe is not None:
                df_list.append(dataframe)
                df_list.append(empty_df)

    print("Merging files")
    merged_df = pd.concat(df_list, ignore_index=True)

    print("Writing to excel")
    merged_df.to_excel("output/End/FinalFinal.xlsx", index=False, header=False)

    print("Deleting temp files")
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

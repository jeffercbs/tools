import pathlib
import os
import pandas as pd
import time

COLUMN_NAMES = ['Nº documento', 'Doc.compensación']
FILE_PATH = 'soportes.xlsx'
DIR_SOPORTS = "soports"
DIR_RENOMBRADOS = "renombrados"

def read_db():
    df = pd.read_excel(FILE_PATH, dtype={COLUMN_NAMES[1]: str})
    df[COLUMN_NAMES[1]] = df[COLUMN_NAMES[1]].apply(lambda x: str(x).split('.')[0])
    return df

def rename_and_move_file(file, nu_doc):
    file = pathlib.Path(file)
    new_file_name = f'soporte de pago {nu_doc}.pdf'
    new_file = pathlib.Path(DIR_RENOMBRADOS) / new_file_name
    new_file.parent.mkdir(parents=True, exist_ok=True)
    file.rename(new_file)

def read_dir_soports():
    start_time = time.time()
    db = read_db()
    renamed_count = 0
    not_found_count = 0

    for file in os.listdir(DIR_SOPORTS):
        file_path = os.path.join(DIR_SOPORTS, file)
        code = file.replace('.pdf', '')

        db_code = db[COLUMN_NAMES[1]]
        row = db.loc[db_code == code]

        if row.empty:
            print(f'No se encontro el codigo {code}')
            not_found_count += 1
            continue
        else:
            rename_and_move_file(file_path, row[COLUMN_NAMES[0]].values[0])
            renamed_count += 1

    end_time = time.time()
    elapsed_time = end_time - start_time

    print(f'Archivos renombrados: {renamed_count}')
    print(f'Archivos no encontrados: {not_found_count}')
    print(f'Tiempo de ejecución: {elapsed_time:.2f} segundos')

if __name__ == '__main__':
    read_dir_soports()
from fpdf import FPDF
import pandas as pd
import openpyxl


def data_access():

    print('reading from data_access')

    workbook = openpyxl.load_workbook('../fichas-Macro.xlsm')

    # Selecciona la hoja de trabajo que contiene los datos
    sheet = workbook['DATOS-indice 2018']

    # Itera sobre cada fila y la imprime en la consola
    for row in sheet.iter_rows(min_row=3, values_only=True):
        print(row)

    # Cierra el archivo de Excel
    workbook.close()


print('hello world')


def ticket_builder(num=0, names='', last_names='', born_date='', born_place='', baptism_date='', baptism_place='', godfather='', godmother='s', mother='', father='', y=10, minister='', hijo='Primero'):
    print('building...')

    complete_name = names.split(' ')
    complete_last_name = last_names.split(' ')

    first_name = 'x'
    second_name = 'x'

    first_last_name = 'x'
    second_last_name = 'x'

    first_name = complete_name[0]

    if len(complete_name) > 1:
        second_name = complete_name[1]

    first_last_name = complete_last_name[0]

    if len(complete_last_name) > 1:
        second_last_name = complete_last_name[1]

    complete_name_str = names + ' ' + last_names

# region first block

    pdf.rect(x=5, y=y, w=200, h=115)
    # pdf.rect(x=5, y=105, w=200, h=90)
    # pdf.rect(x=5, y=200, w=200, h=90)

    pdf.cell(w=10, h=6, txt='No.', align='L',)
    pdf.cell(w=55, h=6, txt=str(num), align='C', border='B')
    # 2
    pdf.cell(w=10, h=6, txt='', align='R', border='', )
    pdf.cell(w=30, h=6, txt='En la parroquia de:', align='L', border='0')
    pdf.multi_cell(w=85, h=5, txt='San Isidro Labrador', align='C', border='B')

    pdf.cell(w=25, h=6, txt='1° Apellido', align='L',)
    pdf.cell(w=40, h=6, txt=first_last_name, align='C', border='B')
    # 2
    pdf.cell(w=10, h=6, txt='', align='C ', border='', )
    pdf.cell(w=50, h=6, txt=baptism_place, align='C', border='B', )
    pdf.cell(w=10, h=6, txt='de', align='C', border='0')
    pdf.multi_cell(w=55, h=6, txt='Diócesis de Bluefields',
                   align='C', border='B')

    pdf.cell(w=25, h=6, txt='2° Apellido', align='L',)
    pdf.cell(w=40, h=6, txt=second_last_name, align='C', border='B')
    # 2
    pdf.cell(w=10, h=6, txt='', align='C', border='', )
    pdf.cell(w=30, h=6, txt='En la fecha de:', align='L',)
    pdf.multi_cell(w=85, h=6, txt=baptism_date, align='C', border='B')

    pdf.cell(w=25, h=6, txt='1° Nombre', align='L',)
    pdf.cell(w=40, h=6, txt=first_name, align='C', border='B')
    # 2
    pdf.cell(w=10, h=6, txt='', align='C', border='', )
    pdf.cell(w=30, h=6, txt='El Ministro: ', align='L',)
    pdf.multi_cell(w=85, h=6, txt=minister, align='C', border='B')

    pdf.cell(w=25, h=6, txt='2° Nombre', align='L',)
    pdf.multi_cell(w=40, h=6, txt=second_name, align='C', border='B')
    pdf.cell(w=65, h=12, txt='', align='C', border='0')
    # 2
    pdf.cell(w=10, h=6, txt='', align='C', border='')
    pdf.multi_cell(
        w=130, h=6, txt='Administró el Sacramento del Bautismo a:', align='C',)
    pdf.cell(w=75, h=12, txt='', align='C', border='0')
    pdf.multi_cell(w=115, h=6, txt=complete_name_str, align='C', border='B')

# endregion

# region second block
    pdf.cell(
        w=65, h=6, txt='Recibió la primera comunión en la', align='L', border='0')
    # 2
    pdf.cell(w=10, h=6, txt='', align='L', border='0')
    pdf.cell(w=30, h=6, txt='Quién nació en: ', align='L', border='0')
    pdf.multi_cell(w=85, h=6, txt=born_place, align='C', border='B')

    pdf.cell(w=30, h=6, txt='parroquia de', align='L', border='0')
    pdf.cell(w=35, h=6, txt=' ', align='C', border='B')
    # 2
    pdf.cell(w=10, h=6, txt='', align='L', border='0')
    pdf.cell(w=30, h=6, txt='en la fecha de: ', align='L', border='0')
    pdf.multi_cell(w=85, h=6, txt=born_date, align='C', border='B')

    pdf.cell(w=10, h=6, txt='el', align='L', border='0')
    pdf.cell(w=25, h=6, txt='', align='L', border='B')
    pdf.cell(w=10, h=6, txt='de', align='L', border='0')
    pdf.cell(w=20, h=6, txt='', align='L', border='B')
    # 2
    pdf.cell(w=10, h=6, txt='', align='L', border='0')
    pdf.cell(w=20, h=6, txt='Hijo: ', align='L', border='0')
    pdf.cell(w=30, h=6, txt=hijo, align='C', border='B')
    pdf.cell(w=30, h=6, txt='de', align='C', border='0')
    pdf.multi_cell(w=35, h=6, txt=mother, align='C', border='B')

    pdf.cell(w=10, h=6, txt='del', align='L', border='0')
    pdf.cell(w=55, h=6, txt='', align='L', border='B')
    # 2
    pdf.cell(w=10, h=6, txt='', align='L', border='0')
    pdf.cell(w=10, h=6, txt='Y de: ', align='L', border='0')
    pdf.multi_cell(w=105, h=6, txt=father, align='L', border='B')
    pdf.multi_cell(w=20, h=12, txt='', align='L', border='0')

# endregion

# region third block
    pdf.cell(w=65, h=6, txt='Contrajo matrimonio en la',
             align='L', border='T')
    # 2
    pdf.cell(w=10, h=6, txt='', align='L', border='0')
    pdf.cell(w=20, h=6, txt='Padrinos', align='L', border='0')
    pdf.multi_cell(w=95, h=6, txt=godfather, align='C', border='B')

    pdf.cell(w=30, h=5, txt='parroquia de', align='L', border='0')
    pdf.cell(w=35, h=6, txt=' ', align='C', border='B')
    # 2
    pdf.cell(w=10, h=6, txt='', align='L', border='0')
    pdf.cell(w=10, h=6, txt='Y', align='L', border='0')
    pdf.multi_cell(w=105, h=6, txt=godmother, align='C', border='B')

    pdf.cell(w=15, h=6, txt='con', align='L', border='0')
    pdf.cell(w=50, h=6, txt='', align='L', border='B')
    # 2
    pdf.multi_cell(
        w=130, h=6, txt='A quienes se les advirtió su obligación y parentesco espiritual', border='0', align='C')

    pdf.cell(w=10, h=6, txt='el', align='L', border='0')
    pdf.cell(w=25, h=6, txt='', align='L', border='B')
    pdf.cell(w=10, h=6, txt='de', align='L', border='0')
    pdf.multi_cell(w=20, h=6, txt='', align='L', border='B')
    pdf.cell(w=10, h=5, txt='del', align='L', border='0')
    pdf.multi_cell(w=55, h=6, txt='', align='L', border='B')
    pdf.cell(w=10, h=6, txt='Nota: ', align='L', border='0')
    pdf.multi_cell(w=55, h=6, txt='', align='L', border='B')
    pdf.multi_cell(w=55, h=20, txt='', align='L', border='0')


# endregion

def read_excel_data(file_path=''):
    print('Reading excel data...')
    data = pd.x(file_path=file_path, header=3, sheet_name=1)
    print(data)


pdf = FPDF(orientation='P', unit='mm', format='A4')

pdf.add_page()

pdf.set_font('Arial', '', 10)

ticket_builder(names='Brandon Isaac', last_names='Fonseca', baptism_date='18/12/2003',
               baptism_place='La Palma', born_date='09/01/99', born_place='Juigalpa',
               godfather='Nelson Oteron', godmother='Martha Cabrera', father='No se sabe',
               mother='Nuvia Fonseca', num=1, minister='Rayan')
ticket_builder(y=143)
pdf.add_page()
ticket_builder()


pdf.output('./report/fichas.pdf')
read_excel_data('../fichas-Macro.xlsm')
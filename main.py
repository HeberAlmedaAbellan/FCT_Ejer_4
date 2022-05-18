import time

from licitaciones import *


def excel_contratacion_estado():
    licitacion = Licitaciones()

    licitacion.nav()
    licitacion.get_num_rows()
    licitacion.get_num_cols()
    licitacion.get_data()
    licitacion.data_to_excel()
    time.sleep(3)
    licitacion.move_file()


excel_contratacion_estado()
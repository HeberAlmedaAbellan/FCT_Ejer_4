import time

from licitaciones import *

licitacion = Licitaciones()

licitacion.nav()
licitacion.get_num_rows()
licitacion.get_num_cols()
licitacion.get_data()
licitacion.data_to_excel()
time.sleep(5)
licitacion.move_file()
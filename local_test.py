from read_hotelgest import read_excel
from read_new import read_csv
from write_NCS import write_excel

excel = read_csv("uploads\in1.csv").read_csv()
write_excel(excel).write()
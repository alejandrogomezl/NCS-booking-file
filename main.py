from read_hotelgest import read_excel
from write_NCS import write_excel

excel = read_excel("in1.xlsx").read_excel()
write_excel(excel).write()

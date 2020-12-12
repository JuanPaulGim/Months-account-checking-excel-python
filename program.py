import sys
import xlrd
import xlsxwriter

loc, wb, sheet, n_wb = None, None, None, None
ioct = 1
oct_rows =[]
creds_in_sep = {}
cred_saldo = {}
row_values_sep = {}
creds_in_oct = {}


def fill_dics():
	for i in range(1,sheet.nrows):
		creds_in_sep[sheet.cell_value(i,4)] = False
		creds_in_oct[sheet.cell_value(i,4)] = False


def get_group1():
	global loc, wb, sheet, n_wb, oct_rows, creds_sep, ioct
	dic_sep = []
	wi = 1
	ws = n_wb.add_worksheet("Segmento 1")
	ws.write_row(0,0,sheet.row_values(0))
	for i in range(1,sheet.nrows):
		if sheet.cell_value(i,sheet.ncols-1) == "septiembre":
			if sheet.cell_value(i,10) > 0:
				id_cred = sheet.cell_value(i,4)
				saldo = sheet.cell_value(i,10)
				creds_in_sep[id_cred] = True
				cred_saldo[id_cred] = saldo
				row_values_sep[id_cred] = sheet.row_values(i)
				dic_sep.append(id_cred)
			ioct+=1
		elif sheet.cell_value(i,sheet.ncols-1) == "octubre":
			id_cred = sheet.cell_value(i,4)
			creds_in_oct[id_cred] = True
			if id_cred not in dic_sep:
				rowoct = sheet.row_values(i)
				oct_rows.append(rowoct)
				ws.write_row(wi,0,sheet.row_values(i))
				wi+=1
def get_group11():
	global loc, wb, sheet, n_wb, oct_rows, creds_sep, ioct
	ws = n_wb.add_worksheet("Segmento 1.1")
	ws.write_row(0,0,sheet.row_values(0))
	wi = 1
	for r in oct_rows:
		if r[10] == r[11]:
			ws.write_row(wi,0,r)
			wi+=1

def get_group12():
	global loc, wb, sheet, n_wb, oct_rows, creds_sep, ioct
	ws = n_wb.add_worksheet("Segmento 1.2")
	ws.write_row(0,0,sheet.row_values(0))
	wi = 1
	for r in oct_rows:
		if r[10] < r[11]:
			ws.write_row(wi,0,r)
			wi+=1

def get_group13():
	ws = n_wb.add_worksheet("Segmento 1.3")
	ws.write_row(0,0,sheet.row_values(0))
	wi = 1
	for r in oct_rows:
		if r[10] > r[11]:
			ws.write_row(wi,0,r)
			wi+=1

def get_group2():
	global loc, wb, sheet, n_wb, oct_rows, creds_sep, ioct
	ws2 = n_wb.add_worksheet("Segmento 2")
	ws2.write_row(0,0,sheet.row_values(0))
	wi = 1
	for i in range(ioct,sheet.nrows):
		id_cred = sheet.cell_value(i,4)
		if creds_in_sep[id_cred] and creds_in_oct[id_cred] == False:
			print("enter 2")
			ws2.write_row(wi,0,sheet.row_values(i))
			wi+=1

def get_group3():
	global loc, wb, sheet, n_wb, oct_rows, creds_sep, ioct
	ws = n_wb.add_worksheet("Segmento 3")
	first_row = sheet.row_values(0)
	for val in first_row[3:]:
		first_row.append(val)
	ws.write_row(0,0,first_row)
	wi = 1
	for i in range(ioct,sheet.nrows):
		if sheet.cell_value(i,14) == "K":
			id_cred = sheet.cell_value(i,4)
			row = sheet.row_values(i)
			rowf = row_values_sep[id_cred]
			for j in range(3,len(row)):
				rowf.append(row[j])
			ws.write_row(wi,0,rowf)
			wi+=1

def get_group41_group42_group43():
	global loc, wb, sheet, n_wb, oct_rows, creds_sep, ioct
	first_row = sheet.row_values(0)
	for val in first_row[3:]:
		first_row.append(val)
	ws41 = n_wb.add_worksheet("Segmento 4.1")
	ws41.write_row(0,0,first_row)
	ws42 = n_wb.add_worksheet("Segmento 4.2")
	ws42.write_row(0,0,first_row)
	ws43 = n_wb.add_worksheet("Segmento 4.3")
	ws43.write_row(0,0,first_row)
	wi41,wi42,wi43 = 1,1,1
	for i in range(ioct,sheet.nrows):
		id_cred = sheet.cell_value(i,4)
		if creds_in_sep[id_cred] and creds_in_oct[id_cred]:
			saldo = sheet.cell_value(i,10)
			if cred_saldo[id_cred] > saldo:
				row = sheet.row_values(i)
				rowf = row_values_sep[id_cred]
				for j in range(3,len(row)):
					rowf.append(row[j])
				ws41.write_row(wi41,0,rowf)
				wi41+=1
			elif cred_saldo[id_cred] < saldo:
				row = sheet.row_values(i)
				rowf = row_values_sep[id_cred]
				for j in range(3,len(row)):
					rowf.append(row[j])
				ws42.write_row(wi42,0,rowf)
				wi42+=1
			elif cred_saldo[id_cred] == saldo:
				row = sheet.row_values(i)
				rowf = row_values_sep[id_cred]
				for j in range(3,len(row)):
					rowf.append(row[j])
				ws43.write_row(wi43,0,rowf)
				wi43+=1
def main():
	global loc, wb, sheet, n_wb, oct_rows, creds_sep, ioct
	loc = "C:/Users/Juan Pablo/Desktop/ManageBDExcel/database.xlsx"
	wb = xlrd.open_workbook(loc)
	sheet = wb.sheet_by_index(2)
	n_wb = xlsxwriter.Workbook('newbd.xlsx')
	fill_dics()
	print("Generando Segmento 1 . . .")
	get_group1()
	print("Finalizado.")
	print("Generando Segmento 1.1 . . .")
	get_group11()
	print("Finalizado.")
	print("Generando Segmento 1.2 . . .")
	get_group12()
	print("Finalizado.")
	print("Generando Segmento 1.3 . . .")
	get_group13()
	print("Finalizado.")
	print("Generando Segmento 2 . . .")
	get_group2()
	print("Finalizado.")
	print("Generando Segmento 3 . . .")
	get_group3()
	print("Finalizado.")
	print("Generando Segmento 4 . . .")
	get_group41_group42_group43()
	n_wb.close()
	for k in creds_in_sep.keys():
		if creds_in_sep[k] and creds_in_oct[k] == False:
			print(k)
main()
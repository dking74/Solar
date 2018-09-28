from portinfo import ExcelSheet , PortEntity
from openpyxl.utils import get_column_letter

def convertDuplexNum ( rowDataList ):

	if   rowDataList [ 3 ] == 0: return "Unknown"
	elif rowDataList [ 3 ] == 1: return "Full Duplex"
	elif rowDataList [ 3 ] == 2: return "Half Duplex"
	
def adjustColumnWidth ( excelInstance ):

	for column_cells in excelInstance.worksheet.columns:
		length = max ( len ( str ( cell.value ) or "" ) for cell in column_cells )
		excelInstance.worksheet.column_dimensions [ column_cells [ 0 ].column ].width = length + 2

def main ( ):

	sheet = ExcelSheet ( "H:\Scripts\Solar_Scripts\PortInfo\AMAG Panels.xlsx" )
	num_rows = sheet.worksheet.max_row + 1
	#for row in range ( 1 , num_rows ):
	data  = sheet.readRowFromWorkbook ( 1 , 2 , sheet.worksheet )
	if data [ 1 ] != None:
		port = PortEntity ( "solarwinds.ameren.com" , "e141674" , "Kings74" , data [ 1 ].strip ( ) )
		info = port.getPortInfo ( )
		print ( info )
	try:
		rowDataList = list ( info [ 0 ].values ( ) )
		rowDataList [ 3 ] = convertDuplexNum ( rowDataList )
		sheet.writeToWorkbook ( sheet.worksheet , row , rowDataList , 4 )
	except ( AttributeError , IndexError ) as detail:
		pass
	
	#adjustColumnWidth ( sheet )
	#sheet.saveWorkbook ( )
	
if __name__ == "__main__":
	main ( )
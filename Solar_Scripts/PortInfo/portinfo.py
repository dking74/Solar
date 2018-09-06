from openpyxl import Workbook , load_workbook

class ExcelSheet ( ):

	'''
	Class name    : ExcelSheet
	
	Class Purpose : To handle reading and writing to Excel spreadsheet
	'''

	def __init__ ( self , workbook ):

		'''
		Method name: __init__
		
		Method Purpose: To initialize an excel instance
		
		Parameters:
			- workbook (string): The name of the workbook
		
		Returns: None
		'''

		self.__workbookName = workbook
		self.__workbook     = None

	def readFromWorkbook ( self , numRows , numColumns , startRow=0 ):

		'''
		Method name: readFromWorkbook
		
		Method Purpose: To read from the inputted workbook
		
		Parameters:
			- numRows (integer): The number of rows to read
			- numColumns (integer): The number of columns to read
			- startRow (integer): The row to start reading from
		
		Returns: None
		'''

		if self.__workbook != None:
			self.__workbook  = load_workbook ( self.__workbookName )
		sheet = self.__workbook.active

		for sheet_row in range  ( 
									startRow if startRow < numRows else 0 , 
									numRows if numRows < self.__workbook.max_row + 1
								):
			sheet.
	def writeToWorkbook  ( self ):

		'''
		Method name: writeToWorkbook
		
		Method Purpose: To write to the inputted workbook
		
		Parameters: None
		
		Returns: None
		'''

	def removeSheetFromWorkbook ( self , sheetName ):

		'''
		Method name: removeSheetFromWorkbook
		
		Method Purpose: To remove an inputted sheet from the inputted workbook
		
		Parameters:
			- sheetName (string): The name of the sheet to remove
		
		Returns: None
		'''

		if self.__workbook != None:
			self.__workbook = load_workbook ( self.__workbookName )

		try:
			self.__workbook.remove ( 
					self.__workbook.get_sheet_by_name ( sheetName )
				)
		except:
			print ( "The sheet is unable to be deleted because the name is incorrect." )




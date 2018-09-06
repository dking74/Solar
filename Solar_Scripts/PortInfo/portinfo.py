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

	def readFromWorkbook ( self , numRows , numColumns ):

		'''
		Method name: readFromWorkbook
		
		Method Purpose: To read from the inputted workbook
		
		Parameters:
			- numRows (integer): The number of rows to read
			- numColumns (integer): The number of columns to read
		
		Returns: None
		'''

		self.__workbook  = load_workbook ( self.__workbookName )
		sheet            = self.__workbook.active

	def writeToWorkbook  ( self ):

		'''
		Method name: writeToWorkbook
		
		Method Purpose: To write to the inputted workbook
		
		Parameters: None
		
		Returns: None
		'''

	def removeSheetFromWorkbook ( self , sheetNum ):

		'''
		Method name: removeSheetFromWorkbook
		
		Method Purpose: To remove an inputted sheet from the inputted workbook
		
		Parameters:
			- sheetNum (integer): The sheet number to remove
		
		Returns: None
		'''

		self.__workbook  = load_workbook ( self.__workbookName )




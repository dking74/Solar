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

	def readFromWorkbook ( self , numRows , numColumns , startRow=1 ):

		'''
		Method name: readFromWorkbook
		
		Method Purpose: To read from the inputted workbook
		
		Parameters:
			- numRows (integer): The number of rows to read
			- numColumns (integer): The number of columns to read
			- startRow (integer): The row to start reading from
		
		Returns: A list of all data
		'''

		if self.__workbook == None: 
			self.__workbook  = load_workbook ( self.__workbookName )
		sheet = self.__workbook.active

		spreadsheet_data = []
		for sheet_row in range ( *self.__findReadRange ( sheet.max_row , numRows , startRow ) ):
			spreadsheet_data.append (
						[ 
							sheet.cell ( row=sheet_row , column=column ).value \
							for column in range ( \
								*self.__findReadRange ( sheet.max_column , numColumns ) 
							) 
						]
					)
		return spreadsheet_data		

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

		if self.__workbook == None: self.__workbook = load_workbook ( self.__workbookName )
		try:
			self.__workbook.remove ( 
					self.__workbook.get_sheet_by_name ( sheetName )
					)
		except:
			print ( "The sheet is unable to be deleted because the name is incorrect." )

	def __findReadRange ( self , maxCell , inputCell , startCell=1 ):

		'''
		Method name: __findReadRange
		
		Method Purpose: To find the range being read from Excel Worksheet;
						either the row or column range
		
		Parameters:
			- maxCell (integer): The max number of cells that can be read
			- startCell (integer): The starting cell (row or column)
			- inputCell (integer): The user inputted range to be read
		
		Returns: A tuple containing the read range
		'''

		cellRange = ( 
				startCell    if startCell < inputCell and startCell > 0
							else 1 , 
				inputCell + 1 if inputCell  < maxCell + 1 and inputCell < maxCell + 1 \
							else maxCell + 1
			)
		return cellRange


from openpyxl import Workbook , load_workbook

class ExcelSheet ( ):

	'''
	Class name    : ExcelSheet
	
	Class Purpose : To handle reading and writing to Excel spreadsheet
	'''

	def __init__                ( self , workbook ):

		'''
		Method name: __init__
		
		Method Purpose: To initialize an excel instance
		
		Parameters:
			- workbook (string): The name of the workbook
		
		Returns: None
		'''

		self.__workbookName = workbook
		self.__workbook     = None

	def readFullWorkbook        ( self , numRows , numColumns , sheet , startRow=1 ):

		'''
		Method name: readFromWorkbook
		
		Method Purpose: To read from the inputted workbook
		
		Parameters:
			- numRows (integer): The number of rows to read
			- numColumns (integer): The number of columns to read
			- startRow (integer): The row to start reading from
			- sheet (unknown): The worksheet to read from
		
		Returns: A list of all data
		'''

		spreadsheet_data = []
		for sheet_row in range ( *self.__getReadRange ( sheet.max_row , numRows , startRow ) ):
			rowData = self.readRowFromWorkbook ( sheet_row , numColumns , sheet )
			spreadsheet_data.append ( rowData )
		return spreadsheet_data	

	def readRowFromWorkbook     ( self , rowNum , numColumns , sheet ):

		'''
		Method name: readRowFromWorkbook
		
		Method Purpose: To read an individual row from Workbook
		
		Parameters: 
			- rowNum (integer): The numbered row to read
			- numColumns (integer): The number of columns to read from row
			- sheet (unknown): The sheet that we are reading from
		
		Returns: A list of the data in the row
		'''	

		rowData = 	[ 
						sheet.cell  ( 
							row=rowNum if rowNum < sheet.max_row + 1 else sheet.max_row, \
							column=column 
						).value for column in range ( \
							*self.__getReadRange ( sheet.max_column , numColumns ) 
						)
					]
		return rowData

	def writeToWorkbook         ( self ):

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

		try:
			self.__openWorkbook    ( )
			self.__workbook.remove ( self.__workbook.get_sheet_by_name ( sheetName ) )
		except:
			print ( "The sheet is unable to be deleted because the name is incorrect." )

	def __getReadRange          ( self , maxCell , inputCell , startCell=1 ):

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

	def getWorkbookActiveSheet  ( self ):

		'''
		Method name: __getWorkbookActiveSheet
		
		Method Purpose: Get the Workbooks active sheet
		
		Parameters: None
		
		Returns: The active sheet of the workbook
		'''
		
		self.__openWorkbook ( )
		return self.__workbook.active

	def __openWorkbook          ( self ):

		'''
		Method name: _openWorkbook
		
		Method Purpose: To open the workbook if it is not already
		
		Parameters: None
		
		Returns: None
		'''

		if self.__workbook == None: 
			self.__workbook = load_workbook ( "theking7.xlsx" )#self.__workbookName )
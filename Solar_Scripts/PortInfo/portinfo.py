from openpyxl import Workbook , load_workbook
from orionsdk import SwisClient
import requests

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

		try:
			self.__workbook = self.__openWorkbook ( workbook )
		except Exception as detail:
			print ( detail )
			raise SystemExit ( )

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

		return self.__workbook.active

	def __openWorkbook          ( self , workbookName ):

		'''
		Method name: _openWorkbook

		Method Purpose: To open the workbook if it is not already

		Parameters:
			- workbookName (string): The name of the workbook to open

		Returns: None
		'''

		try:
			workbook = load_workbook ( workbookName )
		except FileNotFoundError:
			raise Exception ( "The file name entered could not be found in the file system." )
		else:
			return workbook

class PortDetails ( ):

	'''
	Class name: PortDetails

	Class Purpose: To obtain individual port details from Solarwinds
	'''

	def __init__ ( self , ipAddress , username , password ):

		'''
		Method name: __init__

		Method Purpose: To initialize a port detail instance

		Parameters:
			- ipAddress (string): The ip address the port is bound to
			- username (string): The username of the user in solarwinds
			- password (string): The password of the user in solarwinds

		Returns: None
		'''

		# the server is unverified --> allow without warning
		verify = False
		if not verify:
			from requests.packages.urllib3.exceptions import InsecureRequestWarning
			requests.packages.urllib3.disable_warnings ( InsecureRequestWarning )

		self.__ipAddress  = ipAddress
		self.__solarwinds = SwisClient ( "solarwinds.ameren.com" , username , password )

	def getPortInfo ( self ):

		'''
		Method name: getPortInfo

		Method Purpose: To get the important port information from Solarwinds

		Parameters: None

		Returns: A dictionary of the details of the query if it is found
				 'None' if results are not found
		'''

		portQueryResults = self.__solarwinds.query  ( 	"""
														SELECT
															e.Ports.Port.Name,
															e.Ports.Port.PortDescription,
															e.Ports.ConnectionType,
															e.MACAddress,
															e.Ports.Port.Speed,
															e.Ports.Port.Duplex
														FROM
															Orion.UDT.Endpoint e
														WHERE
															e.IPAddresses.IPAddress='10.177.216.5'
														"""
													)

		return portQueryResults [ 'results' ]

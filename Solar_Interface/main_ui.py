from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtPrintSupport import *
from dictionaries import *
from orionsdk import SwisClient
import abc , requests , time

class Main_UI         ( QMainWindow ):

	"""
	Class Name    : Main_UI

	Class Purpose : To create the User Interface for the program
	"""
	
	def __init__ ( self , background_pic ):
	
		"""
		Method Name          : __init__

		Method Purpose       : To initialize the instance of the interface
		
		Parameters           :
			- background_pic : The picture to go in the background
		"""

		# call the parent class constructor to have a baseline for the UI
		super( ).__init__( )

		# the server is unverified --> allow without warnings
		verify = False
		if not verify:
			from requests.packages.urllib3.exceptions import InsecureRequestWarning
			requests.packages.urllib3.disable_warnings ( InsecureRequestWarning )

		# create instance variables for checking state of main ui
		self._connected  = False
		self.constructed = True

		# setup screen for logging into SolarWinds
		self._setupLoginScreen ( )

		# if we were able to connect, go into main view
		if self._connected == True:

			# initialize the layout elements
			self.__initializeCentralWidget (                )
			self.__setupMenuBar            (                )
			self.__setupBackground         ( background_pic )
			self.__setupButtonWidget       (                )
			self.__mainWindowDisplay       (                )

		else:
			self.constructed = False

	def setupSolarWindsConnection ( self ):

		'''
		Method name    : setupSolarWindsConnection 
		
		Method Purpose : To ensure SolarWinds connection is setup
		
		Parameters     : None
		
		Returns        :
			- True     : If login was successful
			- False    : If login was unsuccessful
		'''

		# Initialize SolarWinds console
		try:
			self._solar = SwisClient ( "solarwinds.ameren.com" , self.__username_text.text ( ) , self.__password_text.text ( ) )
			self._solar.query        (                              "SELECT AccountID FROM Orion.Accounts"                     )

		# handler exception with bringing up user dialog
		except requests.exceptions.HTTPError:
			error_msg      = QMessageBox (                                                        )
			error_text     =             ( "The username and/or password is incorrect.\n\n" 
							               "Please enter username as 'ameren\(employee #)'\n" 
							               "and the associated password (network password)"
							             )
			error_msg.setText            (                      error_text                        )
			error_msg.setIcon            (                   QMessageBox.Warning                  )
			error_msg.setDefaultButton   (                     QMessageBox.Ok                     )
			error_msg.setWindowTitle     (                     "Error Message"                    )
			error_msg.exec_              (                                                        )
			return False
		
		return True

	def _setupLoginScreen ( self ):

		'''
			Method name    : _setupLoginScreen
		
			Method Purpose : To provide the interface for entering username and password
		
			Parameters     : None
		
			Returns        :
				-username  : The username in the text field
				-password  : The password in the text field
		'''

		self.__information_dialog = QDialog      (           )
		self.__information_dialog.resize         ( 300 , 300 )
		self.__information_dialog.setWindowTitle (  "Login"  )

		# create layout for ameren logo and add logo to it
		ameren_logo = QVBoxLayout (                                                             )
		am_label    = QLabel      (                                                             )
		pix_map     = QPixmap     (                      "Ameren_Missouri_4c.jpg"               )
		am_label.resize           (                             150 , 150                       )
		am_label.setPixmap        ( pix_map.scaled ( am_label.size ( ) , Qt.IgnoreAspectRatio ) )
		ameren_logo.addWidget     (                             am_label                        )
		ameren_logo.setAlignment  (                           Qt.AlignCenter                    )

		# create username and password fields
		username_label       = QLabel     (     "Username: "   )
		password_label       = QLabel     (     "Password: "   )
		self.__username_text  = QLineEdit (                    )
		self.__password_text  = QLineEdit (                    ) 
		self.__password_text.setEchoMode  ( QLineEdit.Password )

		# create container boxes
		username_box = QHBoxLayout ( )
		password_box = QHBoxLayout ( )
		info_box     = QVBoxLayout ( )

		# add username and password fields to containers
		username_box.addWidget (    username_label     )
		username_box.addWidget ( self.__username_text  )
		password_box.addWidget (    password_label     )
		password_box.addWidget ( self.__password_text  )
		info_box.addLayout     (      username_box     )
		info_box.addLayout	   (      password_box     )

		# create a button for qualifying credentials
		but_layout  = QHBoxLayout   (                                )
		done_button = QPushButton   (             "Login"            )
		done_button.clicked.connect (        self.on_done_click      )
		but_layout.addWidget        (            done_button         )

		# create main box for all elements
		main_box = QVBoxLayout (             )
		main_box.addLayout     ( ameren_logo )
		main_box.addLayout     (  info_box   )
		main_box.addLayout     ( but_layout  )

		# set the layout of the dialog and execute
		self.__information_dialog.setLayout ( main_box )
		self.__information_dialog.exec_     (          )

	def __initializeCentralWidget ( self ):

		"""
		Method Name       : __initializeCentralWidget
		Method Purpose    : To setup the central widget with properties
		
		Parameters        : None 
		"""

		# create the main/central widget constrained by the window --> give it properties
		self.__central_widget = QWidget     (                           )
		self.__central_widget.setStyleSheet ( "background-color: black" )
		self.__central_widget.resize        (        1300 , 800         )

	def __setupMenuBar ( self ):

		"""
		Method Name       : __setupMenuBar

		Method Purpose    : To setup the menu bar for the layout
		
		Parameters        : None 
		"""

		# add a menu bar
		mainMenu = self.menuBar       (          )
		fileMenu = mainMenu.addMenu   (  'File'  )
		editMenu = mainMenu.addMenu   (  'Edit'  )
		viewMenu = mainMenu.addMenu   (  'View'  )
		searchMenu = mainMenu.addMenu ( 'Search' )
		toolsMenu = mainMenu.addMenu  (  'Tools' )
		helpMenu = mainMenu.addMenu   (  'Help'  )
		
		# add functionality for a close button
		exitButton = QAction         ( QIcon ( 'exit24.png' ), 'Exit', self )
		exitButton.setShortcut       (              'Ctrl+C'                )
		exitButton.setStatusTip      (         'Exit application'           )
		exitButton.triggered.connect (             self.close               )
		fileMenu.addAction           (             exitButton               )

	def __setupButtonBox ( self ):
		
		"""
		# Method Name       : __setupButtonBox

		# Method Purpose    : To create the box for the buttons
		
		# Parameters        : None 

		Returns           : The box to be added
		"""

		# create the box containing the buttons and add the buttons
		button_box = QVBoxLayout     (                                                                       )
		button_1   = Button_Behavior ( "View Information on node"      , "Node Query"      , 0 , self._solar ) 
		button_2   = Button_Behavior ( "View Information on group"     , "Group Query"     , 1 , self._solar )
		button_3   = Button_Behavior ( "View Information on interface" , "Interface Query" , 2 , self._solar )
		button_4   = Button_Behavior ( "Add/Delete node"               , "Add Node"        , 3 , self._solar ) 
		button_5   = Button_Behavior ( "Add/Delete group"              , "Add Group"       , 4 , self._solar )

		# create a list of all the buttons and then modify the size of each
		button_list = [ button_1 , button_2 , button_3 , button_4 , button_5 ]
		for button in button_list: 
			Button_Behavior.button_sizing ( button , QSizePolicy.Preferred , QSizePolicy.Preferred , type_var="GROW" )

		# add buttons to vertical box
		button_box.setSpacing (    25    )
		button_box.addWidget  ( button_1 )
		button_box.addWidget  ( button_2 )
		button_box.addWidget  ( button_3 )
		button_box.addWidget  ( button_4 )
		button_box.addWidget  ( button_5 )

		return button_box

	def __setupBackground ( self , background_pic ):

		"""
		Method Name       : __setupBackground

		Method Purpose    : To initialize the background for the application
		
		Parameters        : None
		"""

		# create a box containing the picture
		box = QVBoxLayout ( )

		# define a label and add the label to the box
		self.__label     = QLabel                        (                                                				   )
		self.__label.setGeometry                         ( 0, 0 , box.geometry ( ).width ( ) , box.geometry ( ).height ( ) )
		self.__label.setAlignment                        (                         Qt.AlignCenter                          )
		self.__pixel_map = QPixmap                       (                         background_pic         		           )
		self.__label.setPixmap ( self.__pixel_map.scaled (           self.__label.size ( ) , Qt.IgnoreAspectRatio )        )
		box.addWidget                                    (                          self.__label                           )
 
		# add the box to the main widget		 
		self.__central_widget.setLayout ( box )

	def __setupButtonWidget ( self ):

		"""
		Method Name       : __setupButtonWidget

		Method Purpose    : To show the widget for the buttons on top of background
		
		Parameters        : None
		"""

		# create a child widget to add to main widget and give properties
		self.__mid_widget = QWidget     (   self.__central_widget   )
		self.__mid_widget.setLayout     ( self.__setupButtonBox ( ) )
		self.__mid_widget.resize        (         400 , 400         )
		self.__mid_widget.move          (
										 self.__central_widget.rect ( ).center ( ).x ( ) - ( self.__mid_widget.frameSize ( ).width  ( ) / 2 ) , 
										 self.__central_widget.rect ( ).center ( ).y ( ) - ( self.__mid_widget.frameSize ( ).height ( ) / 2 )
										)
		self.__mid_widget.setStyleSheet ( 
										 "QWidget { background-color: white;"
											"border-width: 8px;"
											"border-style: outset;"
											"border-color: black; }"
										 "QPushButton { background-color: green;" 
											"color: white;"
											"border-width: 5px;" 
											"border-style: outset;"
											"border-color: black;"
											"border-radius: 10px;"
											"font: bold 14px; }" 
										)

	def __mainWindowDisplay ( self ):

		"""
		Method Name       : __mainWindowDisplay

		Method Purpose    : To set the main window properties and show it
		
		Parameters        : None 
		"""

		# set properties of the main window
		self.setSizePolicy ( QSizePolicy.Preferred , QSizePolicy.Preferred )
		self.setMinimumHeight (                     600                    )
		self.setMinimumWidth  (                     850                    )  
		self.setWindowTitle   (               "SolarWinds App"             )
		self.setStyleSheet    (          "background-color: white"         )
		self.setGeometry      (             0 , 0 , 1300 , 800             )
		self.setCentralWidget (            self.__central_widget           ) 

		# move application to middle of screen
		self.moveScreenToMiddle ( self )

		# show the screen 
		self.show ( )

	def resizeEvent ( self , event ):
		
		"""
		Method Name       : resizeEvent

		Method Purpose    : To override the parent method and to handle when the window size is being adjusted
		
		Parameters        :
			- event       : The event we are receiving
		"""

		# adjust the size of the layout and move the screen back to the middle
		self.__central_widget.resize (                        self.rect ( ).width ( ) , self.rect ( ).height ( )                         )
		self.__label.resize          ( self.__central_widget.rect ( ).width ( ) - 100 , self.__central_widget.rect ( ).height ( ) - 100  )
		self.__label.setPixmap       (          self.__pixel_map.scaled ( self.__label.size ( ) , Qt.IgnoreAspectRatio )                 )
		self.__mid_widget.resize     ( self.__central_widget.rect ( ).width ( ) / 5 , self.__central_widget.rect ( ).height ( ) * 9 / 10 )
		self.__mid_widget.move       (                                            0 , 30                                                 )
		self.moveScreenToMiddle      (                                            self                                                   )

	@pyqtSlot ( )
	def on_done_click ( self ):	

		'''
			Method name    : on_done_click
		
			Method Purpose : To handle the login button click
		
			Parameters     : None
		
			Returns        : None
		'''
		
		# get the connection status of SolarWinds
		connected = self.setupSolarWindsConnection ( )

		# determine if connected properly; if so, close the dialog and continue
		if connected == True:
			self.__information_dialog.close ( )
			self._connected = True

		# set the text back to nothing and start looking at data
		self.__username_text.setText ( "" )
		self.__password_text.setText ( "" )

	@staticmethod
	def moveScreenToMiddle ( node ):
		
		"""
		Method Name       : __moveScreenToMiddle

		Method Purpose    : To move the layout to the middle of the screen
		
		Parameters        : None
		"""

		qtRectangle = node.frameGeometry                             (                         )
		centerPoint = QDesktopWidget( ).availableGeometry ( ).center (                         )
		qtRectangle.moveCenter                                       (       centerPoint       )
		node.move                                                    ( qtRectangle.topLeft ( ) )

class Button_Behavior ( QPushButton ):
	
	"""
	Class Name    : Button_Behavior

	Class Purpose : To define the behavior of when a button is pressed
	"""
	
	def __init__ ( self , name , message , type_num , solarwinds ):
		
		"""
		Method Name      : __init__

		Method Purpose   : To initialize the instance of a button
		
		Parameters       :
			- name       : The name displayed on the button
			- message    : The message to display when the button is clicked
			- type_num   : The type of button that was clicked on main screen
		"""

		# initialize as a Push button and connect clicked method
		super ( ).__init__   (                         name                              )
		self.clicked.connect ( lambda: self.on_click ( message , type_num , solarwinds ) )
	
	@staticmethod
	def button_sizing ( button_num , policy_x , policy_y , width=100 , height=100 , type_var=None ):

		"""
		Method Name       : __button_sizing

		Method Purpose    : To adjust the size of the buttons
		
		Parameters        :
			- button_num  : The button that we are adjusting
		"""

		button_num.setSizePolicy ( 
			                      policy_x ,
    							  policy_y
							     )
		
		if type_var != "GROW":

			# set the button constraints
			button_num.setMinimumWidth  (      width     )
			button_num.setMaximumWidth  (      width     )
			button_num.setMinimumHeight (     height     )
			button_num.setMaximumHeight (     height     )
			button_num.resize           ( width , height )

	@pyqtSlot ( )
	def on_click ( self , message , type_num , solarwinds ):
		
		"""
		Method Name    : on_click

		Method Purpose : Slot to define the behavior of a button
		
		Parameters     :
			- message  : The message to display when the button is pressed
		"""

		# controller to represent new screen
		if   type_num == 0:
			screen = Node_View      ( solarwinds )
		elif type_num == 1:
			screen = Group_View     ( solarwinds )
		elif type_num == 2:
			screen = Interface_View ( solarwinds )
		elif type_num == 3:
			screen = Node_Add       ( solarwinds )
		elif type_num == 4:
			screen = Group_Add      ( solarwinds )

		# delete the instance
		del screen

class Sub_Screen      (   abc.ABC   ):

	'''
		Class name    : Sub_Screen
	
		Class Purpose : To provide a secondary screen for when user pushed main button
	'''

	def _launchUI ( self , screen_name=None ):

		'''
			Method name       : _setupUI
		
			Method Purpose    : To launch the User Interface w/ properties
		
			Parameters        :
				- screen_name : The name of the screen
		
			Returns           : None	
		'''

		# set properties of node screen
		self._setupComponents       (                                 )
		self._dialog = QDialog      (                                 )
		self._dialog.setWindowTitle (            screen_name          )
		self._dialog.setLayout      (         self._setupUI ( )       )
		self._dialog.setStyleSheet  (     "background-color: white"   )
		self._dialog.setMaximumSize (       QSize ( 1000 , 800 )      )
		self._dialog.setMinimumSize (       QSize ( 600  , 600 )      )
		self._dialog.resize         (            700 , 700            )

	def _executeUI ( self ):

		'''
			Method name    : _executeUI
		
			Method Purpose : To execute the dialog
		
			Parameters     : None
		
			Returns        : None
		'''
		
		self._dialog.exec_ ( )

	def _setupComponents ( self ):
		
		@pyqtSlot ( )
		def exit_button_click ( ):
			self._dialog.close ( )

		# create an exit button for each screen
		self._exit_layout = QVBoxLayout   (                                                                     )
		self._exit_button = QPushButton   (                                 "Exit"                              )
		self._exit_button.setStyleSheet   ( "background-color: green; color: white; font: 12px;"
										    "border-color: black; border-width: 3px; border-style: solid"   
										  )
		Button_Behavior.button_sizing     ( self._exit_button , QSizePolicy.Fixed , QSizePolicy.Fixed , 75 , 30 )
		self._exit_button.clicked.connect (                      lambda: exit_button_click ( )                  )
		self._exit_layout.addWidget       (                           self._exit_button                         )
		self._exit_layout.setAlignment    (                              Qt.AlignLeft                           )

		# create layout for ameren logo and add logo to it
		self._ameren_logo = QVBoxLayout (                                                                         )
		self._am_label    = QLabel      (                                                                         )
		self._pix_map     = QPixmap     (                          "Ameren_Missouri_4c.jpg"                       )
		self._am_label.resize           (                                300 , 150                                )
		self._am_label.setPixmap        ( self._pix_map.scaled ( self._am_label.size ( ) , Qt.IgnoreAspectRatio ) )
		self._ameren_logo.addWidget     (                              self._am_label                             )
		self._ameren_logo.setAlignment  (                              Qt.AlignCenter                             )

		# create layout option for every base class
		self._box        = QGroupBox   ( )
		self._box.setStyleSheet        ( "QGroupBox   { border-color: black;"
														"border-width: 6px;"
														"border-style: outset;"
														"border-radius: 10px;"    
														"background-color: green }"
										  "QPushButton { border-color: black;"
														"border-style: solid;"
														"border-width: 3px;"
														"border-radius: 7px;"
														"font: bold 15px;        }"
										  "QLabel      { font: 11px              }"
										)

	@abc.abstractmethod
	def _setupUI ( self ):

		''' To be implemented by child classes to setup the UI '''

class Node_View       (  Sub_Screen ):

	'''
		Class name    : Node_View
	
		Class Purpose : To handle the node view screen
	'''

	def __init__ ( self , solar ):

		'''
			Method name    : __init__
		
			Method Purpose : To initialize the Node View class
		
			Parameters     : None
		
			Returns        : None
		'''

		# get an instance variable of solarwinds
		self._solar = solar

		# launch the UI
		self._launchUI  ( "Node View" )
		self._executeUI (             )

	def _setupUI ( self ):

		'''
			Method name    : _setupUI ( reimplemented from Base )
		
			Method Purpose : To setup the UI
		
			Parameters     : None
		
			Returns        : None
		
		'''

		# create the main interaction for the dialog and add to horizontal box
		self.down_nodes  = QPushButton (  "View down nodes"   )
		self.detail_node = QPushButton ( "View specific node" )

		# adjust size policy of all interaction items
		interaction_list = [ self.down_nodes , self.detail_node ]
		for button in interaction_list:
			Button_Behavior.button_sizing ( button , QSizePolicy.Fixed , QSizePolicy.Fixed , 300 , 75 )

		# add the interaction elements to the box
		hbox1 = QHBoxLayout (                   )	
		hbox1.addWidget     (  self.down_nodes  )
		hbox1.addWidget     (  self.detail_node )

		# add the box to a group box
		self._box.setLayout ( hbox1 )
		
	 	# create layout to hold group box
		add_box = QHBoxLayout (           )
		add_box.addWidget     ( self._box )

	 	# create main layout for adding other layouts to and add them
		page = QVBoxLayout (                   )
		page.addLayout     ( self._exit_layout )
		page.addLayout     ( self._ameren_logo )
		page.addLayout     (      add_box      )

		# connect the buttons to functions
		self.__connect_buttons ( )

		return page

	def __connect_buttons ( self ):

		'''
			Method name    : _connect_buttons
		
			Method Purpose : To associate the buttons with a click event
		
			Parameters     : None
		
			Returns        : None
		'''

		self.down_nodes.clicked.connect  ( lambda: self.pressed_down ( ) )
		self.detail_node.clicked.connect ( lambda: self.pressed_node ( ) )

	@pyqtSlot ( )
	def pressed_down ( self ):

		'''
			Method name    : pressed_down
		
			Method Purpose : To handle when the view down nodes button is pressed
		
			Parameters     : None
		
			Returns        : None
		'''

		@pyqtSlot ( )
		def label_event ( label_text ):

			'''
				Method name      : label_event
			
				Method Purpose   : A slot to handle a clickable label
			
				Parameters       :
					- label_name : The name of the label

				Returns          : None
			'''

			# create a new object based on the node clicked
			node = Node_Properties ( self._solar , self._node_list [ label_text.row ( ) ] )
			node.getUtilization    (                                                      )

		# retrieve all nodes that are down
		down_nodes = self._solar.query ( "SELECT Nodes.Caption FROM Orion.Nodes WHERE Nodes.Status = 2" )
		nodes_len  = len               (                    down_nodes [ 'results' ]                    )

		# see if there are results
		if nodes_len > 0:

			# create a dialog of the information
			self._info_screen = QDialog     (                                               )
			self._info_screen.setSizePolicy ( QSizePolicy.Preferred , QSizePolicy.Preferred )
			
			# set properties of vertical box that holds all layout info
			vbox = QVBoxLayout (                               )
			vbox.setAlignment  ( Qt.AlignHCenter | Qt.AlignTop )

			# add label indicating to click a node
			inst_label = QLabel      ( "\nClick the node you wish to see information on:\n" )
			inst_label.setStyleSheet (                    "font: bold 18px"                 )
			vbox.addWidget           (                        inst_label                    )

			# set properties of list widget that holds all node info
			list_widget = QListWidget (                                                               )
			list_widget.setStyleSheet ( "border-color: black; border-width: 8px; border-style: solid" )

			# list to hold all node names
			self._node_list = []

			# iterate through each item and print to screen
			for list_num in range ( nodes_len ):
				label_handler = QListWidgetItem ( down_nodes [ 'results' ][ list_num ][ 'Caption' ]  )
				self._node_list.append         ( down_nodes [ 'results' ][ list_num ][ 'Caption' ]  )
				list_widget.addItem             (                   label_handler                    )

			# add the node list to the vbox and add behavior to the list
			vbox.addWidget ( list_widget )
			list_widget.doubleClicked.connect ( label_event )
			
			# set properties of dialog
			self._info_screen.setMaximumSize ( QSize ( 1000 , 1000 ) )
			self._info_screen.setMinimumSize ( QSize ( 800  , 800  ) )
			self._info_screen.setLayout      (          vbox         )
			self._info_screen.setWindowTitle (      "Down Nodes"     )
			self._info_screen.exec_          (                       )

	@pyqtSlot ( )
	def pressed_node ( self ):

		'''
			Method name    : pressed_node
		
			Method Purpose : To handle when the view node information button is pressed
		
			Parameters     : None
		
			Returns        : None
		'''

		print ( "View node button is pressed" )
		

#class Group_View      (  Sub_Screen ):

#class Interface_View  (  Sub_Screen ):

#class Node_Add        (  Sub_Screen ):

#class Group_Add       (  Sub_Screen ):

#class Info_Screen     (  Sub_Screen ):

	# '''
	# 	Class name    : Info_Screen
	
	# 	Class Purpose : To present the info from a different screen in dialog
	# '''

	#def __init__ ( self , information ):

		# '''
		# 	Method name       : __init__
		
		# 	Method Purpose    : To initialize the dialog screen for information
		
		# 	Parameters        : 
		# 		- information : The information we want to display
		
		# 	Returns           : None
		# '''
		
		# store the information into a private member variables
		# self.__information = information 

		# # launch the dialog
		# self._launchUI      ( "Information Screen" )
		# self._dialog.resize (      900 , 900       )
		# self._executeUI     (                      )

	#def _setupUI ( self ):

		# '''
		# 	Method name    : _setupUI
		
		# 	Method Purpose : To setup the User Interface for the Information screen
		
		# 	Parameters     : None
		
		# 	Returns        : None
		# '''

class Node_Properties (             ):

	'''
		Class name    : Node_Properties
	
		Class Purpose : To show the properties of an individual node
	'''

	def __init__       ( self , solar , node_name ):

		'''
			Method name     : __init__
		
			Method Purpose  : To initialize the data for a node instance
		
			Parameters      : 
				- solar     : The solarwinds creation instance
				- node_name : The name of the specific node
		
			Returns         : None
		'''

		self._node       = node_name
		self._solarwinds = solar

		self.nodeMenu ( )

	def nodeMenu       ( self ):

		'''
			Method name    : nodeMenu
		
			Method Purpose : To build the node screen menu
		
			Parameters     : None
		
			Returns        : None
		'''

		dialog = QMainWindow ( )
		dialog.resize    ( 1000 , 1000 )
		dialog.show      ( )

	def getInfo        ( self ):

		'''
			Method name    : getInfo
		
			Method Purpose : To get the information of the node
		
			Parameters     : None
		
			Returns        : The information
		'''

		info_res = self._solarwinds.query ( "SELECT n.NodeID FROM Orion.Nodes n WHERE n.Caption='{}'".format ( self._node ) )
		print ( info_res )

	def getUtilization ( self ):

		'''
			Method name    : getUtilization
		
			Method Purpose : To get the Utilization of the specific node
		
			Parameters     : None
		
			Returns        : None
		'''

		# N.Vendor as Vendor,
		# N.GroupStatus as Status,
		#N.Interfaces.Name as Int_Name

		query = """
				SELECT
					AVG(N.ResponseTimeHistory.Availability) as Availability,
					AVG(N.CPULoad) as CPULoad,
					AVG(N.PercentMemoryUsed) as MemUsed	
				FROM 
					Orion.Nodes N 
				WHERE 
					N.Caption='%s'
				""" % self._node

 
		info = self._solarwinds.query ( query )

		for num in range ( len ( info ['results']) ):
			print ( self._node )
			print ( "\tAvailability: " + str ( info ['results'][num]['Availability'] ) )
			print ( "\tCPULoad: " + str ( info ['results'][num]['CPULoad'] ) )
			#print ( "\tVendor: " + info ['results'][num]['Vendor'] )
			#print ( "\tStatus: " + info ['results'][num]['Status'] )
			print ( "\tMemUsed: " + str ( info ['results'][num]['MemUsed'] ) )
			#print ( "\tInt_Name: " + info ['results'][num]['Int_Name'] )
			#print ( info ['results'][num]['Availability'])

##############################################################################
#
#							REUSABLE CODE
#					
#				DO NOT DELETE, DO NOT DELETE, DO NOT DELETE
#
##############################################################################

		# self._action_box.addWidget ( node_label )
		# self.g


	# @pyqtSlot ( )
	# def go_button_click ( self ):

	# 	'''
	# 		Method name    : go_button_click
		
	# 		Method Purpose : To handle when the go button is pressed
		
	# 		Parameters     : None
		
	# 		Returns        : None
		
	# 	'''

	# 	# capitalize every word in their input
	# 	word_list   = self.node_text.text ( )
	# 	node_search = word_list.title ( )

	# 	# try to find the location code they are referring to
	# 	for key in node_dict:

	# 		try: 
	# 			node_found = node_dict [ key ][ node_search ]
	# 			found      = True
	# 			break

	# 		except KeyError: 
	# 			found = False
			
	# 	# reset the text box
	# 	self.node_text.setText ( "" )

	# 	# show a message saying it can't be found if it cant be
	# 	if found == False:

	# 		msg_box = QMessageBox    (                                                        )
	# 		msg_box.setText          ( "The node entered was not a valid node in SolarWinds!" )
	# 		msg_box.setIcon          (                QMessageBox.Information                 )
	# 		msg_box.setDefaultButton (                     QMessageBox.Ok                     )
	# 		msg_box.setWindowTitle   (                     "Error Message"                    )
	# 		msg_box.exec_            (                                                        )

	# 	# pull the information from Solarwinds
	# 	#else:

	# def __setupUI ( self , display_mess ):

	# 	"""
	# 	Method Name        : __setupUI

	# 	Method Purpose     : To setup the UI of the secondary screen
		
	# 	Parameters         :
	# 		- display_mess : The message to display in the sub-screen label
	# 	"""

	# 	# create layout for ameren logo and add logo to it
	# 	ameren_logo = QVBoxLayout (                                                             )
	# 	am_label    = QLabel      (                                                             )
	# 	am_label.resize           (                         300 , 150                           )
	# 	pix_map     = QPixmap     (                   "Ameren_Missouri_4c.jpg"                  )
	# 	am_label.setPixmap        ( pix_map.scaled ( am_label.size ( ) , Qt.IgnoreAspectRatio ) )
	# 	ameren_logo.addWidget     (                          am_label                           )
	# 	ameren_logo.setAlignment  (                        Qt.AlignCenter                       )

	# 	# create layout for the form that allows user to select what they wish to do
	# 	action_box     = QHBoxLayout     (                           )
		

	# 	# create Group Box widget to group everything in form layout together
	# 	box = QGroupBox   (            )
	# 	box.setLayout     ( action_box )
	# 	box.setStyleSheet ( "QGroupBox   { border-color: black;"
	# 						 		      "border-width: 6px;"
	# 								      "border-style: outset;"
	# 								      "border-radius: 10px; }"
	# 						"QPushButton { border-color: black;"
	# 									  "border-style: solid;"
	# 						 		      "border-width: 1px;   }"
	# 						"QLabel      { font: 11px           }"
	# 					   )

	# 	# create layout to hold group box
	# 	add_box = QHBoxLayout (     )
	# 	add_box.addWidget     ( box )

	# 	# create main layout for adding other layouts to and add them
	# 	page = QVBoxLayout (             )
	# 	page.addLayout     ( ameren_logo )
	# 	page.addLayout     (   add_box   )

	# 	return page

	# def __setupButtons ( self ):

	# 	"""
	# 	Method Name       : __setupButtons

	# 	Method Purpose    : To setup Button Behavior on the sub-screen
		
	# 	Parameters        : None

	# 	Returns           : The new widget created
	# 	"""

	# 	# create a box containing all the buttons
	# 	button_box = QVBoxLayout (    )
	# 	button_box.setSpacing    ( 15 )

	# 	# create three push buttons
	# 	self.clear_button = QPushButton     (      "Clear"       )
	# 	self.go_button    = QPushButton     (        "Go"        )	
	# 	self.exit_button  = QPushButton     (       "Exit"       )

	# 	# create a list for the push buttons and set the size policy
	# 	button_list = [ self.clear_button , self.go_button ,  self.exit_button ]
	# 	for button in button_list:
			
	# 		Button_Behavior.button_sizing  ( button , QSizePolicy.Preferred , QSizePolicy.Preferred , 50 , 25 )
	# 		button_box.addWidget           (                          button                        )

	# 	# create a new widget and set the layout
	# 	button_widget = QWidget      (                                               )
	# 	button_widget.setLayout      (                button_box                     )
	# 	button_widget.setMaximumSize (            button_widget.sizeHint()           )
	# 	button_widget.setMinimumSize (            button_widget.sizeHint()           )
		
	# 	# connect the buttons to the action for the button
	# 	self.go_button.clicked.connect    ( self.go_button_click    )
	# 	self.clear_button.clicked.connect ( self.clear_button_click )
	# 	self.exit_button.clicked.connect  ( self.exit_button_click  )

	# 	return button_widget

	
from Map_Solar.Entity_Migrate import *

def main ( ):

	solar = IntelligridMig ( "solarwinds.ameren.com" , "Network and Intelligrid Naming Standard.xlsx" )
	solar.readWorkbook     ( "Testing" )
	
	# group_list = solar.getGroupList ( "Testing" )
	# for group in group_list:
		# solar.deleteGroup ( group )
	# solar.deleteGroup ( "Testing" )
	
if __name__ == "__main__":
    main ( )
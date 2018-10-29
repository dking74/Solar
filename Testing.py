from Map_Solar.Entity_Migrate import *
import plotly.graph_objs as graph
import plotly.offline as offline
import plotly.plotly as py
import itertools

solar     = IntelligridMig       ( "solarwinds.ameren.com" , "Network and Intelligrid Naming Standard.xlsx" )
query_res = solar.queryGroupInfo ( 'altnmc01rwa01.ameren.com' , 7 )
print ( query_res )

num_to_str_date = {
	'1' : 'Jan.',
	'2' : 'Feb.',
	'3' : 'Mar.',
	'4' : 'Apr.',
	'5' : 'May.',
	'6' : 'Jun.',
	'7' : 'Jul.',
	'8' : 'Aug.',
	'9' : 'Sep.',
	'10': 'Oct.',
	'11': 'Nov.',
	'12': 'Dec.'
}

num_to_str_weekday = {
	'0' : 'Sun.',
	'1' : 'Mon.',
	'2' : 'Tue.',
	'3' : 'Wed.',
	'4' : 'Thu.',
	'5' : 'Fri.',
	'6' : 'Sat.'
}

day_init     = query_res['results'][0]['Day']
weekDay_init = num_to_str_weekday[str(query_res['results'][0]['WeekDay'])]
month_init   = num_to_str_date[str(query_res['results'][0]['Month'])]

#create the initial lists for creating data
infoString_init = " ".join (
						[
							weekDay_init,
							month_init,
							str(day_init)
						]
					) 
x_axis  = [infoString_init]
y_axis = [0/720, 24/720, 48/720, 72/720, 96/720, 120/720, 144/720, \
				168/720, 192/720, 216/720, 240/720, 264/720, 288/720 ]
z_axis = []
day_data = []
num_days = 1

if len ( query_res [ 'results' ] ) > 0:
	
	#create the initial values for the x-axis
	day_init     = query_res [ 'results' ][ 0 ][   'Day'   ]
	weekDay_init = num_to_str_weekday [ \
					str ( query_res [ 'results' ][ 0 ][ 'WeekDay' ] ) ]
	month_init   = num_to_str_date    [ \
					str ( query_res [ 'results' ][ 0 ][  'Month'  ] ) ]
	
	#create the initial lists for creating data
	infoString_init = " ".join (
							[
								weekDay_init,
								month_init,
								str ( day_init )
							]
						)
	x_axis          = [                       infoString_init                     ]
	y_axis          = [   0/720 ,  24/720 ,  48/720 ,  72/720 , 96/720  , 120/720 , 144/720 , \
	                    168/720 , 192/720 , 216/720 , 240/720 , 264/720 , 288/720 ]
	z_axis          = [                                                           ]
	day_data        = [                                                           ]
	num_days        = 1
	
	#get the last entry in the list
	end_entry = query_res [ 'results' ][ -1 ]
	
	#iterate through every result 
	for it_num , entry in enumerate ( query_res [ 'results' ] ):
	
		#read the values from the results item
		day        = entry              [           'Day'             ]
		hour       = entry              [           'Hour'            ]
		minute     = entry              [          'Minute'           ]
		weekDay    = num_to_str_weekday [ str ( entry [ 'WeekDay' ] ) ]
		month      = num_to_str_date    [ str ( entry [  'Month'  ] ) ]
		
		#see if the day has changed --> if so:
		# 1. add the new string to the x-axis
		# 2. add the day data to the z-axis data
		# 3. clear out the temporary list for the data for the day
		# 4. change the comparision day
		if day != day_init:
			infoString = " ".join (
								[
									weekDay,
									month,
									str ( day )
								]
							)
			x_axis.append ( infoString )
			z_axis.append (  day_data  )
			day_data   = []
			day_init   = day
			num_days   = num_days + 1
		
		#add the entry to the the data for the day
		day_data.append ( entry [ 'Available' ] )	
		
		#if we are at the last entry, append the last
		#day data to the final entry in the z-axis
		if entry == end_entry:
			z_axis.append ( day_data )
		
z_axis = list ( itertools.zip_longest ( *z_axis , fillvalue = '100.0' ) )

axes_data = [
	graph.Heatmap (
		x          = x_axis,
		y          = y_axis,
		z          = z_axis,
		colorscale = 'Bluered',
		xgap       = .6
	)
]

graph_layout = graph.Layout (
	title = 'Availability',
	plot_bgcolor = 'rgb(0,0,0)',
	xaxis = dict 	( 	
					title     = "Day of Week"     , 
					ticks     = 'outside'         ,
					nticks    = num_days          ,
					gridcolor = 'rgb(0,0,240)'    ,
					showgrid  = True
					),
	yaxis = dict 	( 	
					title  = "Time of Day (hour)" , 
					ticks  = 'outside'            , 
					nticks = 12                   , 
					dtick  = 2
					),
	width  = 1500,
	height = 800
)

figure = graph.Figure ( data=axes_data , layout=graph_layout )
#pio.write_image ( figure , 'testFig.png' )
offline.plot          ( figure , filename = 'availability_graph.html' , image='png' , image_filename='Availability' , auto_open=True )
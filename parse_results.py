import sys
import os
from collections import OrderedDict
from openpyxl import load_workbook

# Wrap list elements in tag and return as concatenated string

def tag_reduce( array, tag ):
	return reduce( lambda x, y: x + y, map( lambda x: "<{0}>{1}</{0}>".format( tag, str( x ) ), array ) )

# Very basic checking for correct arguments

if len( sys.argv ) != 3:
	print "Usage: python %s <ken's excel file> <template>" % sys.argv[ 0 ]
	sys.exit()

spreadsheet_file = sys.argv[ 1 ]
template_file = sys.argv[ 2 ]

# Abort if required files does not exist

if not os.path.isfile( spreadsheet_file ):
	print "%s not found" % spreadsheet_file
	sys.exit()
if not os.path.isfile( template_file ):
	print "%s not found" % template_file
	sys.exit()

wb = load_workbook( spreadsheet_file )
ws = wb.get_sheet_by_name( wb.get_sheet_names()[ 0 ] )

heading_filter = set( [ "PLAYER" ] + [ i for i in xrange( 1, 24 ) ] )
player_filter = set( [ "Chris Lewis", "Sarah Nicholson", "Daniel Bevan", "Kelly O'Dwyer", "Andrew Danos", "James Doolan", "Lachlan McNaughton", "Michael Lewis", "TigerLilyBet", "Adrian Maher", "Emma Nicholson", "Courtney VanTongeren", "Sarah Casey", "Tom O'Dwyer", "Tracy Embleton" ] )

headings = OrderedDict( [ ( i.value, i ) for i in ws.rows[ 0 ] if i.value in heading_filter ] )
players = OrderedDict( [ ( i[ 0 ].value, i[ 0 ].address.replace( 'A', '' ) ) for i in ws.rows[ 1: ] if i[ 0 ].value in player_filter ] )

player_vectors = [ [ headings[ j ].offset( row=( int( players[ i ] ) - 1 ) ).value for j in headings ] for i in players ]

# Cut player vectors to remove NoneTypes
cut = sum( [ 1 for i in player_vectors[0] if i != None ] )
player_vectors = [ i[0:cut] for i in player_vectors ]

# Augment player vectors to include explicit rank
for i in xrange( 0, len( player_vectors ) ):
	player_vectors[ i ].insert( 0, i + 1 )

# Calculate previous totals to determine if rank has changed

if cut - 2 <= 0: # edge case: round 1
	last_week = player_vectors
else:
	last_week = sorted( player_vectors, key=lambda x: sum( x[ 2:-1] ) )[ ::-1 ]

for i in xrange( 0, len( last_week ) ):
	curr_rank = last_week[ i ][ 0 ]
	prev_rank = i + 1
	player_index = curr_rank - 1 # by definition the index of the player in the current rankings will be their rank - 1. we can exploit this to simplify the process of checking for changes in the rankings from week to week

	if curr_rank > prev_rank:
		player_vectors[ player_index ].insert( 0, '<span class="glyphicon glyphicon-arrow-up"></span>' )
	elif curr_rank < prev_rank:
		player_vectors[ player_index ].insert( 0, '<span class="glyphicon glyphicon-arrow-down"></span>' )	
	else:
		player_vectors[ player_index ].insert( 0, '<span class="glyphicon glyphicon-minus"></span>' )

template = open( template_file, 'r' ).read()

print template.replace( "{0}", tag_reduce( [ tag_reduce( [ "", "#", "Player" ] + headings.keys()[ 1:cut ], "th" ) ], "tr" ) ).replace( "{1}", tag_reduce( [ tag_reduce( i, "td" ) for i in player_vectors ], "tr" ) ).replace( "{2}", str( 281 + cut * 29 ) )

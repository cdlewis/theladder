# -*- coding: utf-8 -*-

import sys
import os
from jinja2 import Template
from collections import OrderedDict
from openpyxl import load_workbook

reload( sys )
sys.setdefaultencoding( 'UTF-8' )

# Wrap list elements in tag and return as concatenated string

def tag_reduce( array, tag ):
	return reduce( lambda x, y: x + y, map( lambda x: "<{0}>{1}</{0}>".format( tag, str( x ) ), array ) )

# Very basic checking for correct arguments

if len( sys.argv ) != 4:
	print "Usage: python %s <ken's excel file> <players> <template>" % sys.argv[ 0 ]
	sys.exit()

spreadsheet_file = sys.argv[ 1 ]
players_file = sys.argv[ 2 ]
template_file = sys.argv[ 3 ]

# Abort if required files does not exist

for file in [ spreadsheet_file, players_file, template_file ]:
	if not os.path.isfile( file ):
		print "%s not found, exiting..." % file
		sys.exit()

wb = load_workbook( spreadsheet_file )
ws = wb.get_sheet_by_name( wb.get_sheet_names()[ 0 ] )

heading_filter = set( [ "PLAYER" ] + [ i for i in xrange( 1, 24 ) ] )

# Process player file

raw_players = open( players_file, "r" ).read().split( "\n" )
player_filter = set()
zero_players = set()
for i in raw_players:
	if i == "": # ignore empty players
		continue
	if i[ -1 ] == "0": # detect zero player
		i = i.replace( " 0", "" )
		zero_players.add( i )
	player_filter.add( i )

headings = OrderedDict( [ ( i.value, i ) for i in ws.rows[ 0 ] if i.value in heading_filter ] )
players = OrderedDict( [ ( i[ 0 ].value, i[ 0 ].address.replace( 'A', '' ) ) for i in ws.rows[ 1: ] if i[ 0 ].value in player_filter ] )

player_vectors = [ [ headings[ j ].offset( row=( int( players[ i ] ) - 1 ) ).value for j in headings ] for i in players ]

# Account for zero players

for player in player_vectors:
	if player[ 0 ] not in zero_players:
		continue
	for j in xrange( 1, len( player ) ):
		if player[ j ] is not None:
			player[ j ] = 9 - player[ j ]

player_vectors.sort( key=lambda x: -1 * sum( [ i for i in x[ 1: ] if i is not None ] ) ) # sort to correctly position zero players

# Cut player/header vectors to remove NoneTypes

cut = sum( [ 1 for i in player_vectors[0] if i != None ] )
player_vectors = [ i[ 0:cut ] for i in player_vectors ]
headings = headings.keys()[ 1:cut ]

# Augment player vectors to include explicit rank

ladder = [ { 'rank': i, \
             'name': player_vectors[ i ][ 0 ], \
             'total': sum( player_vectors[ i ][ 1: ] ), \
             'rounds': player_vectors[ i ][ 1: ] \
} for i in xrange( 0, len( player_vectors ) ) ]

# Calculate last week's rank

# Create least of (name, total) tuples sorted by name
last_week = sorted( [ ( i[ 0 ], sum( i[ 1:-1] ) ) for i in player_vectors ], key=lambda x: x[ 0 ] )

# Re-order list by score
last_week.sort( key=lambda x: -x[ 1 ] )

# Convert to map of name -> rank
last_week = dict( [ ( last_week[ i ][ 0 ], i ) for i in xrange( 0, len( last_week ) ) ] )

# Render

template = Template( open( template_file, 'r' ).read() )
print template.render( locals() )

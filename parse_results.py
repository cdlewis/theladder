# -*- coding: utf-8 -*-

import sys
import os
import itertools
from collections import OrderedDict
from jinja2 import Template
from openpyxl import load_workbook

reload( sys )
sys.setdefaultencoding( 'UTF-8' )

# Rank players, those with same group function should have same rank
def rank( players, sort_key, group, rank_name ):
    players.sort( key=sort_key ) # order players

    # create a map of scores to the number of players with that score
    score_rank = dict( [ ( i[ 0 ], len( list( i[ 1 ] ) ) ) for i in itertools.groupby( players, lambda x: x[ group ] ) ] )

    # Each player has a rank of n + 1, where n is the number of players with 
    # a higher score. We do this by starting at 1 and incrementing by the
    # number of players with a given score (people_at_rank) when the score
    # for a player is not equal to the score of the previous player.
    rank = 1
    people_at_rank = 0
    for pos, i in enumerate( players ):
        if pos is not 0 and i[ group ] != players[ pos - 1 ][ group ]: # change in rank
            rank += people_at_rank
            people_at_rank = 1
        else:
            people_at_rank += 1
        players[ pos ][ rank_name ] = rank
        players[ pos ][ rank_name + "_pretty" ] = "=" if score_rank[ i[ group ] ] > 1 else ""

    return players

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

# Cut player/header vectors to remove NoneTypes

cut = sum( [ 1 for i in player_vectors[ 0 ] if i != None ] )
player_vectors = [ i[ 0:cut ] for i in player_vectors ]
headings = headings.keys()[ 1:cut ]

# Create a list of basic player structs

ladder = [ { 'name': i[ 0 ], \
             'total': sum( i[ 1: ] ), \
             'last_round': sum( i[ 1:-1 ] ), \
             'rounds': i[ 1: ] } for i in player_vectors ]

# Augment structs with current and previous rank

ladder = rank( ladder, lambda x: ( -1 * x[ "last_round" ], x[ "name" ] ), "last_round", "last_rank" )
ladder = rank( ladder, lambda x: ( -1 * x[ "total" ], x[ "name" ] ), "total", "rank" )

# Render

template = Template( open( template_file, 'r' ).read() )
print template.render( locals() )

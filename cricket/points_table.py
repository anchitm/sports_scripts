#! /usr/bin/env python

from xlrd import open_workbook
from xlwt import Workbook, easyxf
from datetime import date
import math

# ------------
# Helper Class
# ------------

class Teams:
  def __init__(self,name):
    self.name                = name
    self.played              = 0
    self.won                 = 0
    self.lost                = 0
    self.forDict             = dict([(val,0) for val in ['Runs','Balls']])
    self.againstDict         = dict([(val,0) for val in ['Runs','Balls']])
    self.missedUmpiring      = 0
    self.points              = 0  # Represents total points, including bonus points
    self.bonusPoints         = 0
    self.nrr                 = 0.0

  def updateStats(self, scoreDictFor, scoreDictAgainst, winner, battedFirst):
    self.played               += 1
    self.forDict['Runs']      += scoreDictFor['Runs']
    self.forDict['Balls']     += scoreDictFor['Balls']
    self.againstDict['Runs']  += scoreDictAgainst['Runs']
    self.againstDict['Balls'] += scoreDictAgainst['Balls']
    if winner == True:
      self.won    += 1
      if self._gotBonusPoint(scoreDictFor, scoreDictAgainst, battedFirst):
	self.points += 5
	self.bonusPoints += 1
      else:
	self.points += 4
    else:
      self.lost   += 1

  def computeNRR(self):
    self.nrr = round(6.0 * 
                    (float(self.forDict['Runs'])/self.forDict['Balls'] - 
                     float(self.againstDict['Runs'])/self.againstDict['Balls']),
                    NRR_DEC_PLACES_ROUND)

  def applyUmpiringPenalty(self):
    self.missedUmpiring += 1
    self.points         -= MISSED_UMPIRING_PENALTY_POINTS

  # Determine whether a WINNING team gets a bonus point or not
  def _gotBonusPoint(self, scoreDictFor, scoreDictAgainst, battedFirst):
    if battedFirst:
      return scoreDictFor['Runs'] >= BONUS_POINT_RUN_RATE_MULTIPLIER * scoreDictAgainst['Runs']
    else:
      target = scoreDictAgainst['Runs'] + 1
      maxBallsForBonusPoint = math.floor(
	  target*scoreDictAgainst['Balls']/(BONUS_POINT_RUN_RATE_MULTIPLIER * scoreDictAgainst['Runs']))
      #print target
      #print maxBallsForBonusPoint
      return scoreDictFor['Balls'] <= maxBallsForBonusPoint

# ---------------
# Helper Function
# ---------------

def format_spec_float(no_of_dec_places):
  format_str = "%" + ".%d" %(no_of_dec_places) + "f"
  return format_str

# ------------
# Main Program
# ------------

if __name__ == "__main__":

  global NRR_DEC_PLACES_ROUND, MISSED_UMPIRING_PENALTY_POINTS, BONUS_POINT_RUN_RATE_MULTIPLIER

  TEAM_NAMES = ['Stallions',
                'Indians',
		'Dragons',
                'Titans']

  MAX_OVERS                       = 8
  NRR_DEC_PLACES_ROUND            = 6
  NRR_DEC_PLACES_PRINT            = 3
  START_ROW                       = 3
  START_COLUMN                    = 1 # Column B in Excel
  NO_OF_MATCHES                   = 18
  MISSED_UMPIRING_PENALTY_POINTS  = 2
  BONUS_POINT_RUN_RATE_MULTIPLIER = 1.25

  DATE_STRING = str(date.today())

  # ------------------------------------
  # Populate dictionary of match details
  # ------------------------------------

  print "Parsing NRR_Calc_2013.xls to get match details....."
  
  book = open_workbook("NRR_Calc_2013.xls")
  sheet = book.sheets()[0]

  matches = {}
  for row_idx in range(START_ROW,START_ROW+NO_OF_MATCHES):
    matchID = int(sheet.cell(row_idx,START_COLUMN).value)
    matches[matchID] = {}
    matches[matchID]['Team One'] = str(sheet.cell(row_idx,START_COLUMN+1).value)
    
    matches[matchID]['Team Two'] = str(sheet.cell(row_idx,START_COLUMN+2).value)
    
    matches[matchID]['Score One'] = {}
    matches[matchID]['Score Two'] = {}
    sheetRunsOne = sheet.cell(row_idx,START_COLUMN+3).value
    sheetRunsTwo = sheet.cell(row_idx,START_COLUMN+6).value
    
    if sheetRunsOne in ['',None] or sheetRunsTwo in ['',None]:
      matches[matchID]['Score One']['Runs']  = 0
      matches[matchID]['Score One']['Balls'] = 0
      matches[matchID]['Score Two']['Runs']  = 0
      matches[matchID]['Score Two']['Balls'] = 0
    else:
      matches[matchID]['Score One']['Runs'] = int(sheetRunsOne)
      matches[matchID]['Score Two']['Runs'] = int(sheetRunsTwo)
      
      allOutOne = str(sheet.cell(row_idx,START_COLUMN+5).value).lower()
      if allOutOne in ['yes','y']:
        ballsOne = MAX_OVERS * 6
      else:
        oversOne_str   = str(sheet.cell(row_idx,START_COLUMN+4).value)
        oversOne_tuple = oversOne_str.split('.')
        ballsOne = int(oversOne_tuple[0]) * 6 + int(oversOne_tuple[1])
      matches[matchID]['Score One']['Balls'] = ballsOne

      allOutTwo = str(sheet.cell(row_idx,START_COLUMN+8).value).lower()
      if allOutTwo in ['yes','y']:
        ballsTwo = MAX_OVERS * 6
      else:
        oversTwo_str   = str(sheet.cell(row_idx,START_COLUMN+7).value)
        oversTwo_tuple = oversTwo_str.split('.')
        ballsTwo = int(oversTwo_tuple[0]) * 6 + int(oversTwo_tuple[1])
      matches[matchID]['Score Two']['Balls'] = ballsTwo
    
    matches[matchID]['Team Batting First'] = str(sheet.cell(row_idx,START_COLUMN+9).value)
    matches[matchID]['Winner']             = str(sheet.cell(row_idx,START_COLUMN+10).value)
    matches[matchID]['Umpiring Team']      = str(sheet.cell(row_idx,START_COLUMN+11).value)
    matches[matchID]['Umpire Present']     = str(sheet.cell(row_idx,START_COLUMN+12).value)

  # --------------------------------
  # Populate list of team statistics
  # --------------------------------

  print "Computing team statistics....."

  # Initialize list of team objects
  team_list = [Teams(name) for name in TEAM_NAMES]

  # Update statistics for each team based on match results
  for key in sorted(matches.keys()):
    if matches[key]['Winner'] not in ['',None]: # To ensure only completed matches are used to update stats
      for team in team_list:
        if team.name == matches[key]['Team One']:
          team.updateStats(matches[key]['Score One'],
                           matches[key]['Score Two'],
                           matches[key]['Winner'].lower() == team.name.lower(),
			   matches[key]['Team Batting First'].lower() == team.name.lower())
        elif team.name == matches[key]['Team Two']:
          team.updateStats(matches[key]['Score Two'],
                           matches[key]['Score One'],
                           matches[key]['Winner'].lower() == team.name.lower(),
			   matches[key]['Team Batting First'].lower() == team.name.lower())
        elif team.name == matches[key]['Umpiring Team'] and matches[key]['Umpire Present'].lower() == "no":
          team.applyUmpiringPenalty()

  # Compute Net Run Rate (NRR)
  for team in team_list:
    if team.played > 0:
      team.computeNRR()

  # Sort list of team objects in descending order first by Points, then by NRR
  team_list.sort(key=lambda team: (team.points,team.won,team.bonusPoints,team.nrr), reverse=True)

  # --------------------------
  # Print to Output Excel File
  # --------------------------

  print "Generating Points Table....."

  book_op = Workbook()
  sheet = book_op.add_sheet('Points Table')
  
  # Write column headers
  COLUMNS = ['Position','Team','Played','Won','Lost','Missed Umpiring','Points (Total)','Bonus Points','NRR','For (Runs/Overs)','Against (Runs/Overs)']
  row = sheet.row(1)
  for col_idx in range(len(COLUMNS)):
    row.write(col_idx+2,COLUMNS[col_idx],easyxf(
      'font: bold True;'
      'borders: left thin, right thin, top thin, bottom thin;'
      'alignment: horizontal center;'))

  # Set custom column widths where default is not enough (empirical)
  sheet.col(3).width  = len(COLUMNS[len(COLUMNS)-1])*256
  sheet.col(7).width  = len(COLUMNS[len(COLUMNS)-1])*256
  sheet.col(8).width  = len(COLUMNS[len(COLUMNS)-1])*256
  sheet.col(9).width  = len(COLUMNS[len(COLUMNS)-1])*256
  sheet.col(11).width  = len(COLUMNS[len(COLUMNS)-1])*256
  sheet.col(12).width = len(COLUMNS[len(COLUMNS)-1])*256

  # Write points table
  format_str = format_spec_float(NRR_DEC_PLACES_PRINT)
  row_idx = 2
  pos = 1
  for team in team_list:
    for_str     = "%d/%d.%d" %(team.forDict['Runs'],team.forDict['Balls']/6,team.forDict['Balls']%6)
    against_str = "%d/%d.%d" %(team.againstDict['Runs'],team.againstDict['Balls']/6,team.againstDict['Balls']%6)
    if team.nrr > 0:
      nrr_str = "+" + format_str %(team.nrr)
    else:
      nrr_str     =  format_str %(team.nrr)
    
    row = sheet.row(row_idx)
    row.write(2,pos,easyxf(
      'borders: left thin, right thin;'
      'alignment: horizontal center;'))
    row.write(3,team.name,easyxf(
      'borders: left thin, right thin;'))
    row.write(4,team.played,easyxf(
      'borders: left thin, right thin;'
      'alignment: horizontal center;'))
    row.write(5,team.won,easyxf(
      'borders: left thin, right thin;'
      'alignment: horizontal center;'))
    row.write(6,team.lost,easyxf(
      'borders: left thin, right thin;'
      'alignment: horizontal center;'))
    row.write(7,team.missedUmpiring,easyxf(
      'borders: left thin, right thin;'
      'alignment: horizontal center;'))
    row.write(8,team.points,easyxf(
      'borders: left thin, right thin;'
      'alignment: horizontal center;'))
    row.write(9,team.bonusPoints,easyxf(
      'borders: left thin, right thin;'
      'alignment: horizontal center;'))
    row.write(10,nrr_str,easyxf(
      'borders: left thin, right thin;'
      'alignment: horizontal center;'))
    row.write(11,for_str,easyxf(
      'borders: left thin, right thin;'
      'alignment: horizontal center;'))
    row.write(12,against_str,easyxf(
      'borders: left thin, right thin;'
      'alignment: horizontal center;'))
    row_idx += 1
    pos     += 1

  # Formatting for last row of points table
  for col_idx in range(len(COLUMNS)):
    sheet.row(row_idx).write(col_idx+2,'',easyxf(
      'borders: top thin;'))

  # Save output Excel File
  op_filename = "Points_%s.xls" %(DATE_STRING)
  book_op.save(op_filename)

  print "Please open %s to see the Points Table" %(op_filename)

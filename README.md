# WeeklyPickPoolSpreadsheet

WeeklyPickPoolSpreadsheet is an Excel spreadsheet for tracking people, their picks, and outcomes of games (generally NFL, but it could easily be used for other head-to-head sports) for a Weekly Pick'em Pool. It automates some of the more mundane aspects of keeping track of a Pick'em Pool, like counting how many correct picks each person has made and highlighting the winning person/people at the end of the week based on total correct picks or, in the case of a tie, showing which person had the closest (+/-) pick for the total "points" for a final game. It also makes entering people's picks easier by highlighting missing picks and invalid/misspelled picks.

For Weekly Pick sheets, I use [PrintYourBrackets](https://www.printyourbrackets.com/nflweekly.html)

Notes:
- This spreadsheet was created using Microsoft 365 Excel.
- It does not use any VBA, only built-in functions, formulas, and conditional formatting.

Basic Features
==============
* Easily keep track of Weekly Pick'em Pools
* Display number of players in pool to calculate pot
* Show percentage of players who picked each team
* Enter winning team in the "Won" row to highlight correct and incorrect picks
* Enter final points (tie-breaker rules) in "Won" row over points to highlight winner

Conditional Formatting Features
===============================
* Enter winning team to highlight correct (soft green) and incorrect (soft red) picks for each player for each game
* Automatically updates each players "Correct" total
* Highlight "Top 1" correct player(s) counts (soft green)
* Enter total points for all Monday night games, after entering all winning teams, to highlight winning player for the week (green)
* When there is a "Correct" tie, winner(s) will be based on closest "Points" pick (above or below total Monday games points total)
* Highlight misspelled/invalid picks if pick is neither of the teams for a given column (red) 
* Highlight missing picks/points for assigned players (yellow)

Screenshots
===========
### Entering names and picks.
* Highlights misspelled/invalid entries in red and missing entries in yellow:
<p align="center">
  <img src="https://raw.githubusercontent.com/Ayitaka/WeeklyPickPoolSpreadsheet/main/Screenshots/WeeklyPickPoolSpreadStartMissingAndMisspelled.png" alt="Screenshot" />
</p>

### Halfway through a week's games.
* "Correct" picks highlighted in soft green.
* "Incorrect" pics highlighted in soft red.
* Highest "Correct" count highlighted in soft green in last column.
<p align="center">
  <img src="https://raw.githubusercontent.com/Ayitaka/WeeklyPickPoolSpreadsheet/main/Screenshots/WeeklyPickPoolSpreadSheetHalfway.png" alt="Screenshot" />
</p>

### Final entries as well as Monday night's points.
* Winner's name and points highlighted in green.
<p align="center">
  <img src="https://raw.githubusercontent.com/Ayitaka/WeeklyPickPoolSpreadsheet/main/Screenshots/WeeklyPickPoolSpreadSheetFinal.png" alt="Screenshot" />
</p>

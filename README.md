# VBA-challenge
Create new sheet called Combined_Data, loop thrugh all sheet from Q1to Q4
Create a script that loops through all the stocks for each quarter and outputs the following information:
  -The ticker symbol : Loop through ticker for each quarter in column H
  -Quarterly change: Loop through using close value from last date - open value from first date for each quarter in column J
    -if the value > 0 color in green, value < 0 color in red, value = 0 then  no color
  -Percentage change: Loop through (last date - first date)/ first date *100 for each quarter in column K
    -if the value > 0 color in green, value < 0 color in red, value = 0 then  no color
  - The total stock volume of the stock: Loop through Q1 to Q4 vol in column L
  - Add functionality to the script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". 

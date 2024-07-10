# VBA-challenge
** Create new sheet called Combined_Data, loop thrugh all sheet from Q1to Q4

** Create a script that loops through all the stocks for each quarter and outputs the following information:

  - The ticker symbol : Loop through ticker for each quarter 
  
  - Quarterly change: Loop through using close value from last date - open value from first date for each quarter 
    - if the value > 0 color in green, value < 0 color in red, value = 0 then  no color
    
  - Percentage change: Loop through (last date - first date)/ first date *100 for each quarter 
    - if the value > 0 color in green, value < 0 color in red, value = 0 then  no color
    
  - The total stock volume of the stock: Loop through Q1 to Q4 vol in column L

** The result should match the following image:
![image](https://github.com/faiyted/VBA-challenge/assets/171522014/e42054cf-7a9a-4bd3-b7ca-d3607da795fc)

** Add functionality to the script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". 
![image](https://github.com/faiyted/VBA-challenge/assets/171522014/da30f989-dcd2-4297-b31f-b8bc1d902e80)

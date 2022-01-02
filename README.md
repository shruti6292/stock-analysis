Overview of Project
Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.


Results
This assignment consists of one technical deliverable and a written report to deliver your results. You will submit the following:

 Refactor VBA code and measure performance
This deliverable will include an updated workbook and a folder with PNGs of the pop-ups with script run time



Summary
Download the challenge_starter_code.vbs file and rename it VBA_Challenge.vbs.
Create a folder called “Resources” to hold the run-time pop-up messages that you’ll screenshot after running refactored analyses for 2017 and 2018.
Rename the green_stocks.xlsm file that you used in this module as VBA_Challenge.xlsm.
Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor.
Use the steps below to add code were indicated by the numbered comments in the starter code file.
Step 1a:Create a tickerIndex variable and set it equal to zero before iterating over all the rows. You will use this tickerIndex to access the correct index across the four different arrays you’ll be using: the tickers array and the three output arrays you’ll create in Step 1b.
Step 1b:Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
The tickerVolumes array should be a long data type.
The tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.
Step 2a:Create a for loop to initialize the tickerVolumes to zero.
Step 2b:Create a for loop that will loop over all the rows in the spreadsheet.
Step 3a:Inside the for loop in Step 2b, write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
Use the tickerIndex variable as the index.
If you’d like a hint on how to increase the current tickerVolumes by using the tickerIndex variable as the index, that’s totally okay. If not, that’s great too. You can always revisit this later if you change your mind.

Step 3b: Write an if-then statement to check if the current row is the first row with the selected ticket index. If it is, then assign the current starting price to the tickerStartingPrices variable.
Step 3c: Write an if-then statement to check if the current row is the last row with the selected ticket index. If it is, then assign the current closing price to the tickerEndingPrices variable.
Step 3d: Write a script that increases the ticket index if the next row’s ticker doesn’t match the previous row’s ticker.
Step 4: Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.
Finally, run the stock analysis, then confirm that your stock analysis outputs for 2017 and 2018 are the same as they were in the module (as shown in the images below). In your Resources folder, save the pop-up messages showing elapsed run time for the refactored code as VBA_Challenge_2017.png and VBA_Challenge_2018.png. Then, save the changes to your workbook.

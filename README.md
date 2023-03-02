# VBA-challenge

### Stock Change
This project takes an excel workbook that contains different stocks and the yearly changes they go through and uses that information to determine the yearly change and total stock volume of each listed.

## Process

### Writing and Testing VBA in alphabetical_testing.xlsm

Due to the extensive excel workbook I ultimately want the code to run on, I used the testing excel workbook to ensure the code ran properly before running it on the final stock workbook.

#### Order of Process

1. The first code I wrote was simply cycling through the first sheet and printing out every Ticker in the first column exactly one time each
..* This involved setting up the Sub, my variables, and setting the titles of my new columns: including determining the end of column 1, and then beginning the first For Loop and If/Then Statement to iterate through the first column and only print each Ticker once.
2. Once I knew that the Ticker's printed out correctly, I then added to the For Loop to also save the first open and last close value for a ticker in order to calculate the yearly change. 
..* This included a separate If/Then Statement to determine and Format the Yearly Change cells to be red if negative, green if positive, and light purple for no change
3. It was a short leap from the Yearly Change to the Percent Change
..* Simply adding to the column to the right of Yearly Change to calculate the percentage and format the cell accordingly
4. Last addition to the first For Loop was a few lines of code that would continue to add the Total Stock Volume until a new Ticker was detected and then printing that total next to the Percent Change
5. With the first For Loop complete, it was time to set up and create the "Calculated Values" table
..* First I created the new variables to store the highest, lowest, and total volume and set the titles for the table columns and rows. \
..* Then another For Loop to iterate through until the end of the new list of Tickers and an If/Then to compare the values and save the Ticker and value of the highest, lowest, and total volume to use at the end
..* Lastly, out of the For Loop, updated the cells in the table to show the Ticker and percent of the Highest and Lowest Percentage Changed Stock as well as the Ticker and Total Stock Volume of the Greatest Total Volume Stock
6. The final step in the code was to set it up to run through all worksheets in the workbook.
..* This involved putting all of the code into one final For Loop that iterate through Each Worksheet and update the tables for that Worksheet
..* I then had to read through the entire code to ensure that anything that called a cell had "ws." in front

#### Challenges Along the Way

The main issue I came to was towards the end of the coding when I was transitioning everything to run through all worksheets. I had everything set up and ready to go. Everything ran smoothly except for the updating of the Greatest Percent Increase and Greatest Percent Decrease. For some reason the Ticker was updating properly but the percentage was always the percent from the first sheet. I added a message box to print out what was stored in those variables and those values were correct, but the visual output in the cells were wrong. It took me far too long to realize that the one error I had made was forgetting to add the two last "ws." to the formatting line of code. With that error fixed, it was time to run the code on the Final Excel Workbook.

### The Final Test

With the code tested and approved in the test workbook it was time to run it on the real thing. It took quite a bit of time for the code to run, but everything went smoothly.

#### The first rows of each sheet in the Workbook

2018:

![alt text](https://github.com/beccasolomon22/VBA-challenge/blob/main/Multiple_year_stock_data%20-%20Excel%203_2_2023%201_30_27%20PM.png)

2019:

![alt text](https://github.com/beccasolomon22/VBA-challenge/blob/main/Multiple_year_stock_data%20-%20Excel%203_2_2023%201_30_35%20PM.png)

2020:

![alt text](https://github.com/beccasolomon22/VBA-challenge/blob/main/Multiple_year_stock_data%20-%20Excel%203_2_2023%201_30_44%20PM.png)

###Additions
For ease of use, I added a button on the first sheet that will run the code on all sheets. I also added a separate button on each sheet that allows the user to run the code for just that sheet. This way The User does not need to open Visual Basic in order to get results.

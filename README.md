# Stock Analysis

## Overview of Project

### Purpose

The purpose of this analysis using VBA in Excel was to refactor the Module 2 solution code to loop through all of the data found in the VBA_Challenge.xlsm file and test if the refactored code allows the analysis process to be done at a faster rate and/or with less computational memory.

## Process

### The following steps were completed on the refactored code.

1. Created tickerIndex variable and set it equal to zero before iterating over all rows.

![image](https://user-images.githubusercontent.com/93107507/142784976-06c8ae25-dfd8-474f-8b00-22233849ff92.png)

2. Created three output arrays and saved them as their respective data types: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

![image](https://user-images.githubusercontent.com/93107507/142785033-de683755-6cd3-40cb-b77f-f33b80adfdcb.png)

3. Created a for loop to initialize the tickerVolumes to zero. 
![image](https://user-images.githubusercontent.com/93107507/142785183-dd0f9e26-9f05-4af6-8fb4-9d39c93a3ccb.png)

4. I created a for loop that would loop over all the rows in the selected year spreadsheet.
![image](https://user-images.githubusercontent.com/93107507/142785207-07e01900-c5fd-4a23-8094-02cd2db3f92a.png)

5. Added a cript that increases the current tickerVolumes variable and adds the ticer volume for the current stick ticker.
![image](https://user-images.githubusercontent.com/93107507/142785318-80b75040-180b-408a-985a-c49b7b6df4ac.png)
*Code copied from Challenge documentation: https://courses.bootcampspot.com/courses/976/assignments/19969?module_item_id=357762*

6. Used if-then statements to assign the current annual opening price to the tickerStartingPrices array on the codition that the current row is the first row with the selected ticker.
![image](https://user-images.githubusercontent.com/93107507/142785513-d94e5717-87d0-4d41-b193-28a250d34b5c.png)

7. Used if-then statements to assign the current annual closing price to the tickerEndingPrices array on the codition that the current row is the last row with the selected ticker, and to create a script that increases tickerIndex on the condition that the next row's ticker doesn't match the previous row's ticker.
![image](https://user-images.githubusercontent.com/93107507/142785603-729a93b6-5e72-40d4-8e66-c382d1faa94a.png)

8. Ran analysis and confirmed that the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workboot matched with the outputs in the Challenge.

![image](https://user-images.githubusercontent.com/93107507/142785757-9d56b6de-cf72-4e1d-a1f1-523c4600f911.png)

![image](https://user-images.githubusercontent.com/93107507/142785768-71ba25c4-100c-43e5-a7ee-aea349956ad8.png)


## Results




















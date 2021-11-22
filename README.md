# Stock Analysis

## Overview of Project

### Purpose

The purpose of this analysis using VBA in Excel was to refactor the Module 2 solution code to loop through all of the data found in the VBA_Challenge.xlsm file and test if the refactored code allows the analysis process to be done at a faster rate and/or with less computational memory.

## Results

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

**Examples provided in Challenge

![image](https://user-images.githubusercontent.com/93107507/142785859-e3197fd0-ab9b-4bb8-8dcd-faf187629519.png)
![image](https://user-images.githubusercontent.com/93107507/142785867-bd8f31db-33e7-4251-8830-408f2d9d9d74.png)



**Output from Refactored Code

![image](https://user-images.githubusercontent.com/93107507/142785757-9d56b6de-cf72-4e1d-a1f1-523c4600f911.png)
![image](https://user-images.githubusercontent.com/93107507/142785768-71ba25c4-100c-43e5-a7ee-aea349956ad8.png)

9. The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png

*Time on VBA_Challenge_2017.png*

<img width="555" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/93107507/142786158-8032eb55-d938-4c68-a490-e7dc01bd3846.png">

*Time on VBA_Challenge_2018.png*

<img width="611" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/93107507/142786170-4cb8b58f-ba65-4e38-887b-a6b689d52fe4.PNG">


## Summary
### Advantages and disadvantages of refactoring code
One advantage that I found of refacturing the code was the lack of repetitiveness. Through VBA code refactoring we were able to tackle multiple design patterns in one matrix, rather than creating multiple for each analysis year. By reusing code through loops and arrays, we were also able to spend more time creating succinct.

While refactoring code may save time writing code, thinking about how to imrpove the code to be more precise and straight forward takes an increased amount of time.

### Pros and cons of applying refactoring to the original VBA script.
> Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

As mentioned in the challenge, and going off of my opinions on the advantages and disadvantages of the code: Refactoring code can take an increased amount of time if it is messy and lacks documentation to understand what the original programmer meant. The pros of refacturing code is making it more efficient, especially once the logic of the code is made easier for future users to read. 



















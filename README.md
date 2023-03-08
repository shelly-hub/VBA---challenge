**VBA Challenge**

This is an exercise to utilise VBA tool in excel to process data and create into the values required by writing a series of code in the visual script.
In this case, stock market data is used for this challenge across 3 different years. 

**Description** 

Visual Basic Analysis known as VBS is a tool under Developer in Excel is used to enable user to write codes to process large data. 
- The biggest benefit here is written code can be used to process multiple worksheets in one excel file without the need to process one by one. 
- Another benefit is the code can be used in other excel files, hence writing multiple coding in different excel files are not needed. 
- The one problem that I encounter when processing this large data file is it can take some time to generate the codes. Separate worksheet with shorter data has to be used to find out the workability of the code before implementing into the original large dataset. 

**Instructions**

To open the file, as following: 

To open the VBA file, users need to download most updated Microsoft Excel version with macro-enabling function. After opening the excel file, go to Developer tab, then run the files under Macro. 

To write the code, as following: 

1. "dim statement": This is to state variables used in this script. Be mindful, as this is a real large data set, do not use integer or long, but use double

2. "For each ws in worksheets ": This is to enable the code to run across three worksheets that have stock data for 3 years separetely. 

   When this function is done, subsequent code "next ws" is used to close the code. 

3. Analyse the values needed to generate, and specify their given location. 

    The data required as following:

		"Ticker"

		"Yearly Change"

		"Percentage Change"

		"Total Stock Volume"

		"Greatest Percentage increase", 

		"Greatest Percentage decrease", 

		"Greatest total volume".


4. Identify all initial values before process the data

		"Total Stock Volume" and "Yearly Change": Need to start from 0 value, so the value in the data can be added one by one

		"Opening Price" : Need to set the initial value given, as stockmarket do not calculate yearly change from same day of openign and closing price. 

		"Number of row": This is to identify which row the summary table starts to input the data value. 

5. Then start process the data through looping method.

   Other methods used are listed below:
	
		- "IF" "ENDIF" "ELSE" "ENDIF" METHOD
		- "FORMATPERCENTAGE"
		- Conditional Formatting in the coding to colour the cells which red marked negative changes in stockmarket, and green marked positive changes in stockmarket. 

**Credits**

	- Data Analytics Bootcamp Lecturer and Tutors

  - Data Analytics Lecturer Slides under Activities 02 VBA Scripting file

  - Links below:

https://www.youtube.com/watch?v=M3OE7Z62oGM&list=PLNIs-AWhQzckr8Dgmgb3akx_gFMnpxTN5&index=6&ab_channel=WiseOwlTutorials
 
 https://www.freecodecamp.org/news/how-to-write-a-good-readme-file/
 
 https://www.exceldemy.com/vba-format-percentage-2-decimal-places/
 
https://www.google.com/search?q=how+to+find+maximum+value+in+vba&oq=how+to+find+maximum+value+in+vba&aqs=chrome..69i57j0i22i30l3j0i390l3.7183j0j7&sourceid=chrome&ie=UTF-8#fpstate=ive&vld=cid:c8ca667f,vid:TBLlSKkOujA
	
	




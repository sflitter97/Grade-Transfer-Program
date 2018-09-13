# Running Grade Transfer Program

## 1) Setup
	A. Directory
		-- After downloading this program, the Stepik Excel Sheet, and the Canvas Excel sheet make sure that all three are placed in the same directory
	B. Excel
		-- The Canvas Excel sheet needs to be exported from .csv to .xlsx.  To do this, first open the .csv file.  Then click "File", click "Save As", choose the current directory, and choose the filetype to be "Excel Workbook"
	C. Python Program
		-- There are 4 variables that need to be modified to make the program run as you desire.  When you open the python script there is a section near the top called "VARIABLES THAT NEED TO BE CHANGED" that contains all four of these variables.  Here's how you should change them:
		i. importFile - Place the name of the Stepik file (including the file extension) you will be extracting grades from in between the quotation marks.
		ii. inputGradeColumn - This variable should be an INTEGER.  The int in this variable represents the column number (1 indexed) from the Stepik file that you will be pulling grades from
		iii. exportWorkbook - Place the name of the canvas file (including the file extension) that you will be exporting grades to in between the single quotation marks.
		iv. outputGradeColumn - This variable should be an INTEGER.  The int in this variable represents the column number (1 indexed) from the Canvas file that you will be exporting grades to.

## 2) Running Program
	-- When the program runs it will copy all the selected grades over from the Stepik Excel sheet to the Canvas Excel sheet.  However, there is always names that appear in the Stepik file that don't appear in the gradebook.  All of these names will be placed in a file called "unmatched_Students.txt" that is generated in the current directory.  The user will have to go through this file by hand and match these students to accounts in Canvas.

## 3) Cleaning Up Afterwards
	A. Delete Text Files
		-- There will be 2 files generated by running this program.  The first is the "unmatched_Students.txt", and the second is a file called "double_Entries.txt" that contains all the students whose names appear more than once in the Stepik file.  Both of these will need to be deleted before the program can be run a second time.
	B. Export Excel Sheet
		-- Before the Canvas Excel sheet can be re-uploaded to Canvas it will need to be exported back to a .csv file.  This can be done by reversing the process in 1A.  Click "File", click "Save As", and then save it as a .csv file.


Happy Grading!
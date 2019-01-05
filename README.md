# PDM BOM Parser
This code will take PDM CSV BOM export and make it fancy. You can use the script standalone with a file selection GUI, or use the Drag & Drop batch script to quickly run it. 

## Requires
Requires Python 3, Pandas, EasyGui

Install modules with following command: `pip install pandas easygui`

## Instructions:

### Export the BOM from PDM as follows:

 1. Go to BOM tab of assembly in PDM
 2. In the left column of drop down menus, choose the following
	 1. DSS BOM
	 2. Indented
	 3. Show Tree 
 3. In the middle column of drop down menus, choose the following
	 1. Not Activated (doesn't matter)
	 2. Show Selected
	 3. As Built
 4. In the right column, select the correct Version and Configuration
 5. Click the **Name** column header to sort all the names
 6. Click the Save icon in the top-right, and choose **Save As**
 7. Change file type to **.CSV** but leave file name as-is
 8. Save to local computer and choose *Yes* when asked to add a **Level Column**
 
### Process BOM:

Either run `bom_creator.py` script - this will give you a file selection prompt. Or, use the `dragbomhere.bat` batch script and drag and drop your CSV file to process.

The Excel BOM will output in the same directory as the py script. 

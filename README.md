# tools
A collection of random tools for everyday analyses in biodiversity informatics

##GetRedListCategory
This macro allows you to query the IUCN red list and retrieve the red list category for a scientific name in Excel. 

To use this function:

1. Open a new or existing Excel workbook.

2. Press "ALT + F11" or go to Tools>Macro>Visual Basic Editor to open the Visual Basic Editor.

3. In the Visual Basic Editor, go to "File" > "Import File" and select the "GetRedListCategory.bas" file that you downloaded earlier.

4. Replace "your key" with your IUCN API key

5. Save your VBA project and return to the Excel workbook.

6. In a cell in a column, enter a species scientific name.

7. In the cell to the right of the species name, enter the following formula: =GetRedListCategory(A2)
Replace "A2" with the cell reference of the first cell containing the species name.

Press enter to calculate the formula. This may crash or take a long time if it's run over many names at once. I've only tested this script in Excel for Mac. It may not work on windows or other OS. You can try Chat GPT to adjust.

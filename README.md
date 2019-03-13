# Excel VBA-demos: copy-paste, etc

VBA code for Excel:

This repo contains various VBA code to impleeent various commonplace Excel functions and features, 
such as copying and pasting data from one worksheet to another. 

Each of the uploaded text files contain different programs/demos of VBA code. The text files were created using Vim via a Unix terminal. For reference, the file "Vim- create text file to save VBA copy-paste code" from this repo shows how to use Vim to create and save a new text file.

One of the VBA scripts/demos copies and pastes data from a specified range of cells from one worksheet to another. While this example may seem trivial, this could potentially improve efficiency when using copy-paste functions in Excel. For example, for a worksheet with many rows or columns of data, this code can be used--with some small adjustments, of course--to copy-paste a particular subset of the rows or columns, and then paste the data into another worksheet. This way, for instance, you can create a new worksheet containing specific subsections, time periods, or other aspects of the data that you want to analyze or visualize separately from the worksheet containing the full dataset, without having to use a filter. 

The following VBA script (see the Excel_VBA_code/VBA_copy_paste_to_other_worksheet.txt file) for this copy-paste program is as follows:

Sub copy_paste_to_worksheet():
        'Paste data from Rounds Worksheet, with range of cells B1 to P83871.
        Worksheets("Rounds").Range("B1:P83871").Copy Worksheets("Pasted_data").Range("A1:P83871")
End Sub


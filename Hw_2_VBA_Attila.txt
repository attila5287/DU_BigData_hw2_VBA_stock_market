Sub Homework2_Easy()

Dim ws As Worksheet
Dim Bottom_ROW As Long
Dim Ticker As String
Dim Sum_ROW As Long
Dim Active_ROW As Long
Dim Stock_Volume_Total As Double

'below is a FOR LOOP that will go through each worksheet in the workbook
For Each ws In ThisWorkbook.Sheets
    
'below will enable the code repeat same operations for each worksheet with less rows in the code 
    With ws
        
        'below assigns the bottom row to a variable so data will be limited to active cells within the worksheet
		Bottom_ROW = .Range("A1").End(xlDown).Row
        
        'header cells will be named as below
        .Range("L1").Value = "Ticker"
		'title for total stock vol
        .Range("M1").Value = "Total Stock Value"
        
		'macro will begin writing from the second row after header and will drag down 
		Sum_ROW = 2
        
        'loop will continue until the bottom cell taken into account
        For Active_ROW = 2 To Bottom_ROW
			
			'below assigns a variable in order to compare with the next row
             Ticker = .Cells(Active_ROW, 1)
			 'below places a conditional statement to check if the first column value is same with the next below
             If .Cells(Active_ROW + 1, 1) = Ticker Then
             
                'below will be creating a pivot sum under the same variable 
                Stock_Volume_Total = Stock_Volume_Total + .Cells(Active_ROW, 7)
            
            Else
                
                'accumulator for total stock volume will add up for each cell
                Stock_Volume_Total = Stock_Volume_Total + .Cells(Active_ROW, 7)
                
				'macro prints new values
				.Cells(Sum_ROW, 12).Value = Ticker
                .Cells(Sum_ROW, 13).Value = Stock_Volume_Total
                
                'macro moves to the next row
                Sum_ROW = Sum_ROW + 1
                
                'variable will be reset to zero for the next loop
                Stock_Volume_Total = 0
            
            End If
        
		'next loop will continue until all rows are gone through loop
        Next Active_ROW
    
    End With

	'macro will move to the next worksheet when all rows within the current ws completed
Next ws

End Sub





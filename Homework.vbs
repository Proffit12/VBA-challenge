Attribute VB_Name = "Module1"
Sub homework()

    'Loop through worksheets
    'For Each ws In Worksheet (I couldn't get my loop through worksheets to work)
    
        'Denote all of the vairiables that will be used
        Dim tick As String
        tick = " "
        Dim tot As Double
        tot = 0
        Dim Opn As Double
        Opn = 0
        Dim clse As Double
        clse = 0
        Dim chnge As Double
        chnge = 0
        Dim prct As Double
        prct = 0
        Dim row As Long
        row = 2
        Dim lastrow As Long
        Dim i As Long
        
        'denote lastrow to count how many active rows are in the worksheet
        lastrow = Cells(Rows.Count, 1).End(xlUp).row
        
        'Add the headers for the summary table
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volumee"
        
        'Set opn equal to the first open stock value
        Opn = Cells(2, 3)
        
        'Loop through all the rows starting with the first stock
        For i = 2 To lastrow
        'Set if statement for when the stock changes in the worksheet
        If Cells(i + 1, 1) <> Cells(i, 1) Then
            'Set tick equal to the ticker name as it changes
            tick = Cells(i, 1)
            'set clse equal to the closing price of the last stock of the year
            clse = Cells(i, 6)
            'Find the yearly difference by subtracting opn from clse
            chnge = clse - Opn
            'Find the percent change with this formula
            prct = (chnge / Opn) * 100
            'Add all of the previous stock volumes to the last one of the year
            tot = tot + Cells(i, 7)
            'Insert the ticker into the summary table
            Range("I" & row) = tick
            'Insert the yearly change into the summary table
            Range("J" & row) = chnge
            'If statement to highlight the percent change as green if it's positive anded if it's negative
            If (chnge > 0) Then
                Range("J" & row).Interior.ColorIndex = 4
            ElseIf (chnge < 0) Then
                Range("J" & row).Interior.ColorIndex = 3
            End If
            'Add a percent sign to the percent changes
            Range("K" & row) = (CStr(prct) & "%")
            'Insert the stock total into the row
            Range("L" & row) = tot
            'Move to the next row in the summary table for the next stock
            row = row + 1
            'reset the yearly change
            chnge = 0
            'reset the year end close amount
            clse = 0
            'set the next year open stock amount
            Opn = Cells(i + 1, 3)
            Else
            'add all the stock volumes together if you're in the same stock
            tot = tot + Cells(i, 7)
        End If
        
        Next i
    'Next ws
    
End Sub



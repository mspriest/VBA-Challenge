Attribute VB_Name = "Module1"
Sub stock()

'Declare worksheet variable
Dim ws As Worksheet

'Loop through worksheet & activate
For Each ws In Worksheets
    ws.Activate

'declare ticker variable
Dim ticker As String

'declare yearlychange & set initial value
Dim yearlychange As Double

'declare percentage change & set intial value
Dim percentage As Double

'declare ticker volume & set initial value
Dim volume As Double
volume = 0

'create summary table
Dim summarytable As Integer
summarytable = 2

'determine the last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'add headers
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    
'declare counter for yearly change
    Dim counter As Integer
    counter = 0

    'declare open date
    Dim opendate As Double

    'declare close date
    Dim closedate As Double

    
'loop through ticker
For i = 2 To lastrow

    'If counter = 0 Then...
    If counter = 0 Then
    
        opendate = Cells(i, 3).Value
    
    End If
   
   'determine yearly change
        counter = counter + 1
          
    'Check if ticker name is the same, if not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'set ticker value
    ticker = Cells(i, 1).Value
      
    'close date
    closedate = Cells(i, 6).Value
    
    'calculate yearly change
    yearlychange = closedate - opendate
        
    'calculate percentage
    If opendate <> 0 Then
        percentage = yearlychange / opendate
    
    End If
    
    'add volume to total
    volume = volume + Cells(i, 7).Value
    
    'Print ticker to summary table
    Range("I" & summarytable).Value = ticker
    
    'Print yearlychange to summary table
    Range("J" & summarytable).Value = yearlychange
    
        'Add formatting
        If Range("J" & summarytable).Value < 0 Then
            Range("J" & summarytable).Interior.ColorIndex = 3
            
        Else
            Range("J" & summarytable).Interior.ColorIndex = 4
        
        End If
    
    'Print percentage to summary table
    Range("K" & summarytable).Value = percentage
    
    'Print volume to summary table
    Range("L" & summarytable).Value = volume
    
    'Add one to the summary table row
    summarytable = summarytable + 1
        
    'Reset yearly change
    counter = 0
      
    'Reset volume
    volume = 0
    
'If cell immediately following is the same ticker name

    Else
    
'Add to the volume
volume = volume + Cells(i, 7).Value
    
    End If

Next i
    
  Next ws
    
End Sub

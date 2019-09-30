Sub wrap()

Dim ws As Worksheet



For Each ws In ThisWorkbook.Worksheets


'Declare variables
Dim ticker As String
Dim lastrow As Long
Dim i As Long
Dim a As Long
Dim tickrow As Long
Dim tickercount As Long
Dim cpri As Long
Dim opri As Long
Dim yearlychange As Long
Dim percentchange As Double
Dim maxper As Double
Dim minper As Double
Dim maxtotal As Double
Dim maxtick As String
Dim maxtotaltick
Dim mintick As String
Dim total As Double


Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage Change"
Cells(1, 12).Value = "Total Volume"

Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"


'Set variables
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
tickercount = 1
total = 0
      
       
For i = 2 To lastrow
       
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    
            ticker = Cells(i, 1).Value
            
            tickercount = tickercount + 1
       
            opri = Cells(i, 3).Value + 1E-07

            Cells(tickercount, 9).Value = ticker
        
            'Cells(tickercount, 10).Value = opri
       
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
            ticker = Cells(i + 1, 1).Value

            cpri = Cells(i, 6).Value + 1E-07
         
            'Cells(tickercount, 11).Value = cpri
            
            yearlychange = cpri - opri
            
            Cells(tickercount, 10).Value = yearlychange
            
            Cells(tickercount, 11).Value = yearlychange / (opri + 1E-07)
            
            Cells(tickercount, 12).Value = total
            
            total = 0
            
    Else
    
            total = total + Cells(i, 7).Value
            
                            
    End If

Next i

maxper = WorksheetFunction.Max(ws.Range("K2:K" & tickercount))
minper = WorksheetFunction.Min(ws.Range("K2:K" & tickercount))
maxtotal = WorksheetFunction.Max(ws.Range("L2:L" & tickercount))

'MsgBox maxper
'MsgBox minper


For i = 2 To tickercount
    
        If Cells(i, 11).Value = maxper Then
        
            maxtick = Cells(i, 9).Value
            
        ElseIf Cells(i, 11).Value = minper Then
        
            mintick = Cells(i, 9).Value
            
        ElseIf Cells(i, 12).Value = maxtotal Then
        
            maxtotaltick = Cells(i, 9).Value
          
        End If
        
Next i

maxper = WorksheetFunction.Max(ws.Range("K2:K" & tickercount))
minper = WorksheetFunction.Min(ws.Range("K2:K" & tickercount))
maxtotal = WorksheetFunction.Max(ws.Range("L2:L" & tickercount))

Cells(2, 16).Value = maxper
Cells(3, 16).Value = minper
Cells(4, 16).Value = maxtotal

Cells(2, 15).Value = maxtick
Cells(3, 15).Value = mintick
Cells(4, 15).Value = maxtotaltick

Cells(2, 14).Value = "Greatest Percentage Increase"
Cells(3, 14).Value = "Greatest Percentage Decrease"
Cells(4, 14).Value = "Greatest Total Volume"


    For i = 2 To tickercount
    
        If Cells(i, 10).Value > 0 Then
            
            Cells(i, 10).Interior.ColorIndex = 4
                
        ElseIf Cells(i, 10).Value <= 0 Then
        
            Cells(i, 10).Interior.ColorIndex = 3
            
        End If
    
    Next i
    
ws.Range("K2:K" & tickercount).Style = "Percent"
Cells(2, 16).Style = "Percent"
Cells(3, 16).Style = "Percent"


Next


End Sub



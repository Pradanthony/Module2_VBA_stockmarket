Attribute VB_Name = "Module1"
Sub worksheetloop():

Dim ws_count As Integer
Dim n As Integer
ws_count = ActiveWorkbook.Worksheets.Count

For n = 1 To (ws_count)

    ActiveWorkbook.Worksheets(n).Activate


    Dim ticker As String
    Dim volume As LongLong
    Dim daterow As Integer
    Dim rowcount As Long
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    rowcount = ActiveSheet.Rows.Count
    daterow = 2
    
    For I = 2 To (rowcount - 1)
        If (Cells(I + 1, 1).Value <> Cells(I, 1).Value) Then
        ticker = Cells(I, 1).Value
        volume = volume + Cells(I, 7).Value
        Cells(daterow, 9).Value = ticker
        Cells(daterow, 12).Value = volume
        daterow = daterow + 1
        
        volume = 0
        
        Else
        volume = volume + Cells(I, 7).Value
        
        End If
    Next I
    

    Dim firstRow As Long
    Dim lastRow As Long
    Dim searchValue As String
    Dim tickerRow As Integer
    Dim openprice As Double
    Dim closeprice As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    
    tickerRow = Cells(Rows.Count, "I").End(xlUp).Row
    
    For j = 1 To (tickerRow - 1)
        searchValue = Cells(j + 1, 9).Value
        
    With ActiveSheet.Range("A:A")
        Set c = .Find(searchValue, LookIn:=xlValues, lookat:=xlWhole)
        If Not c Is Nothing Then
        firstRow = c.Row
            
        Set c = .Find(searchValue, LookIn:=xlValues, lookat:=xlWhole, searchDirection:=xlPrevious)
        lastRow = c.Row
            
        openprice = Cells(firstRow, 3).Value
        closeprice = Cells(lastRow, 6).Value
        yearlychange = closeprice - openprice
        Cells(j + 1, 10).Value = yearlychange
        If Cells(j + 1, 10).Value >= 0 Then
            Cells(j + 1, 10).Interior.ColorIndex = 4
            Else: Cells(j + 1, 10).Interior.ColorIndex = 3
        End If
            
        percentchange = yearlychange / openprice
        Cells(j + 1, 11).Value = percentchange
        Cells(j + 1, 11).NumberFormat = "0.00%"
               
        Else
            
        End If
    End With
    Next j
    

    Dim maxvalue As Double
    Dim minvalue As Double
    Dim maxtotal As LongLong
    
    Range("o2").Value = "Greatest % Increase"
    Range("o3").Value = "Greatest % Decrease"
    Range("o4").Value = "Greatest Total Volume"
    Range("p1").Value = "Ticker"
    Range("q1").Value = "Value"
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
    
    maxvalue = Cells(2, 11).Value
    tickerRow = Cells(Rows.Count, "k").End(xlUp).Row
    
    For I = 2 To (tickerRow - 1)
        If Cells(I, 11).Value > maxvalue Then
        maxvalue = Cells(I, 11).Value
        Cells(2, 16).Value = Cells(I, 9).Value
        End If
        
    Next I
    Cells(2, 17).Value = maxvalue
    
    minvalue = Cells(2, 11).Value
    
    For j = 2 To (tickerRow - 1)
        If Cells(j, 11).Value < minvalue Then
        minvalue = Cells(j, 11).Value
        Cells(3, 16).Value = Cells(j, 9).Value
        End If
        
    Next j
    Cells(3, 17).Value = minvalue
    
    maxtotal = Cells(2, 12).Value
    
    For k = 2 To (tickerRow - 1)
        If Cells(k, 12).Value > maxtotal Then
        maxtotal = Cells(k, 12).Value
        Cells(4, 16).Value = Cells(k, 9).Value
        End If
        
    Next k
    Cells(4, 17).Value = maxtotal
    
 MsgBox ActiveWorkbook.Worksheets(n).Name

Next n

End Sub




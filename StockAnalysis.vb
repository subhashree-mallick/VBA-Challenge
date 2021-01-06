Attribute VB_Name = "Module1"
Sub wallStreet():
 
 ' Created a Variable to Hold File Name, Last Row, Last Column, and Year
  Dim i As Long
  Dim change As Single
  Dim j As Integer
  Dim start As Long
  Dim total As Double
  Dim lastrow As Long
  Dim percentchange As Single
  Dim ws As Worksheet
  
  For Each ws In Worksheets
  
   'Set Initial Value for each worksheet
  j = 0
  total = 0
  change = 0
  start = 2
  
  
  
  
  'set the title row
  ws.Range("i1").Value = "Ticker"
  ws.Range("j1").Value = "Yearly Change"
  ws.Range("k1").Value = "Percent Change"
  ws.Range("l1").Value = "Total Stock Volume Of Stock"
  
 
  ' Determine the Last Row
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  
  For i = 2 To lastrow
    
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      
        'stores results in variables
        total = total + ws.Cells(i, 7).Value
        If total = 0 Then
            ws.Range("i" & 2 + j).Value = Cells(i, 1).Value
            ws.Range("j" & 2 + j).Value = 0
            ws.Range("k" & 2 + j).Value = "%" & 0
            ws.Range("l" & 2 + j).Value = 0
        Else
            If ws.Cells(start, 3) = 0 Then
             For find_value = start To i
              If ws.Cells(find_value, 3).Value <> 0 Then
               start = find_value
               Exit For
              End If
             Next find_value
        End If
            
            
        change = (ws.Cells(i, 6) - ws.Cells(start, 3))
        If Cells(start, 3) = 0 Then
                'percentchange = Round((change / Cells(start, 3) * 100), 2)
                percentchange = 0
        Else
                percentchange = Round((change / Cells(start, 3) * 100), 2)
        End If
        'start the next stock ticker
        start = i + 1
        
        'print the results
        ws.Range("i" & 2 + j).Value = ws.Cells(i, 1).Value
        ws.Range("j" & 2 + j).Value = Round(change, 2)
        ws.Range("k" & 2 + j).Value = "%" & percentchange
        ws.Range("l" & 2 + j).Value = total
    
    
       ' color positive green and negetive red
        Select Case change
         Case Is > 0
          ws.Range("j" & 2 + j).Interior.ColorIndex = 4
         Case Is < 0
          ws.Range("j" & 2 + j).Interior.ColorIndex = 3
         Case Else
          ws.Range("j" & 2 + j).Interior.ColorIndex = 0
        End Select
        
      End If
      
       'reset variable for new stock ticker
         
         total = 0
         change = 0
         j = j + 1
         
        
        
   Else
         
      total = total + ws.Cells(i, 7).Value
      
  End If
     
Next i

Next ws
 
End Sub

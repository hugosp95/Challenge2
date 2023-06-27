VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stocks()
    Dim lastRows As Long
    lastRows = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim sumRow As Long
    sumRow = 2
    
    Dim ticker As String
    Dim volumeTotal As LongLong
    Dim closing As Single
    Dim opening As Single
    Dim Ychange As Single
    Dim Pchange As Single
    
    Dim GPChange As Single
    Dim GPChangeRow As Long
    
    Dim LPChange As Single
    Dim LPChangeRow As Long
    
    Dim GVolTotal As LongLong
    Dim GVolTotalRow As Long
    
    
    

    
    Range("I1,O1") = "Ticker"
    Range("J1") = "Yearly change"
    Range("K1") = "Percentage change"
    Range("L1") = "Total Stock Volume"
    Range("N2") = "Greatest % Increase"
    Range("N3") = "Greatest % Decrease"
    Range("N4") = "Greatest Total Volume"
    Range("O2") = "Greatest Percentage Change"
    Range("P1") = "Value"
    
    GPChange = 0
    GPChangeRow = 0
    
    LPChange = 0
    LPChangeRow = 0
    
    GVolTotal = 0
    GVolTotalRow = 0
    
    
    For i = 2 To lastRows
        volumeTotal = volumeTotal + Cells(i, 7).Value
        
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            ticker = Cells(i, 1).Value
            closing = Cells(i, 6).Value
            Ychange = closing - opening
            'Pchange = ((closing * 100) / opening) - 100
            Pchange = ((closing - opening) / opening) * 100
            Cells(sumRow, 9).Value = ticker
            Cells(sumRow, 10).Value = Ychange
            Cells(sumRow, 11).Value = Pchange & "%"
            Cells(sumRow, 12).Value = volumeTotal
             If Ychange > 0 Then
                Cells(sumRow, 10).Interior.ColorIndex = 4
            ElseIf Ychange < 0 Then
                Cells(sumRow, 10).Interior.ColorIndex = 3
            End If
            
            If Pchange > GPChangeRow Then
                GPChange = Pchange
                GPChangeRow = sumRow
                Range("O2") = GPChange
            End If
            
            
            If volumeTotal > GVolTotal Then
                GVolTotal = volumeTotal
                GVolTotalRow = sumRow
                Range("O4") = Cells(GVolTotalRow, 9).Value
            End If
            
            If LPChange = 0 Then
            LPChange = Pchange
                LPChangeRow = sumRow
            ElseIf Pchange < LPChange Then
                LPChange = Pchange
                LPChangeRow = sumRow
            End If
            
            
            sumRow = sumRow + 1
            volumeTotal = 0
            Ychange = 0
        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            opening = Cells(i, 3).Value
        End If
    Next i
    Range("P2") = GPChange
    Range("O2") = Cells(GPChangeRow, 9).Value
    
    Range("P3") = LPChange
    Range("O3") = Cells(LPChangeRow, 9).Value
    
    Range("P4") = GVolTotal

   
End Sub


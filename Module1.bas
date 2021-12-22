Attribute VB_Name = "Module1"
Option Explicit
Sub summarize_all():
    Dim numworksheets As Integer
    Dim i As Integer
    
' Loops through worksheets summarizing one at a time

    numworksheets = ActiveWorkbook.Worksheets.Count
    
    For i = 1 To numworksheets
        ThisWorkbook.Worksheets(i).Activate
        summarize_worksheet
    Next i
    
End Sub
Sub summarize_worksheet():
    Dim ticker As String
    Dim row As Long
    Dim openprice As Double
    Dim closeprice As Double
    Dim newrow As Long
    Dim vol As LongLong

' Do initial Formatting
    format
    
' Set initial values
    row = 2
    ticker = Cells(row, 1).Value
    openprice = Cells(row, 3).Value
    newrow = 2
    vol = 0
    
    
'Run until first blank cell
    While Not Cells(row, 1).Value = ""
' only change value of ticker if it does not match the previous rows value and
' make a new row of data, then reset values
        If Cells(row, 1).Value <> ticker Then
            closeprice = Cells(row - 1, 6)
            Call makerow(newrow, ticker, openprice, closeprice, vol)
            newrow = newrow + 1
            ticker = Cells(row, 1).Value
            openprice = Cells(row, 3).Value
            vol = 0
        End If
' accumulate stock volume and increment row
        vol = vol + Cells(row, 7)
        row = row + 1
    Wend
' make last row
    closeprice = Cells(row - 1, 4)
    Call makerow(newrow, ticker, openprice, closeprice, vol)
' call subroutine for bonus
    greatest
        
End Sub
Sub format():
    Range("I1, P1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
' Edit column widths to fit titles and data
    Columns("O:O").ColumnWidth = 18.21
    Columns("K:K").ColumnWidth = 13
    Columns("L:L").ColumnWidth = 15.8
    Columns("Q:Q").ColumnWidth = 12
End Sub
Sub makerow(newrow, ticker, openprice, closeprice, vol):

' cell values from left to right (columns I to L) : ticker symbol, net change,
' percent change, and total volume
    Cells(newrow, 9) = ticker
    Cells(newrow, 10) = closeprice - openprice

' Red if delta is negative, green if poitive, blue if zero
    If Cells(newrow, 10).Value < 0 Then
        Cells(newrow, 10).Interior.Color = RGB(255, 0, 0)
    ElseIf Cells(newrow, 10).Value > 0 Then
        Cells(newrow, 10).Interior.Color = RGB(0, 255, 0)
    Else
        Cells(newrow, 10).Interior.Color = RGB(0, 0, 255)
    End If

' check if the initial open price is 0 and if so set cell to NA
    If openprice = 0 Then
        Cells(newrow, 11) = "NA"
    Else
        Cells(newrow, 11) = (closeprice - openprice) / openprice
' format result as 2 decimal percentage
        Cells(newrow, 11).NumberFormat = "0.00%"
    End If
    Cells(newrow, 12).Value = vol
    
End Sub
Sub greatest():
    Dim row As Long
    Dim bigp As Double
    Dim rowbigp As Long
    Dim smallp As Double
    Dim rowsmallp As Long
    Dim vol As LongLong
    Dim rowvol As Long
    
    row = 2
    bigp = 0
    smallp = 0
    vol = 0
    
' loop through results, keeping track of largest stockvolume and percent change
' as well as largest negative percent change, track row for easy formatting
    While Not Cells(row, 9).Value = ""
        If Not Cells(row, 11).Value = "NA" Then
            If Cells(row, 11).Value > bigp Then
                bigp = Cells(row, 11).Value
                rowbigp = row
        End If
            If Cells(row, 11).Value < smallp Then
                smallp = Cells(row, 11).Value
                rowsmallp = row
            End If
        End If
        If Cells(row, 12).Value > vol Then
            vol = Cells(row, 12).Value
            rowvol = row
        End If
        row = row + 1
    Wend
    
' format bonus data
    Range("P2").Value = Cells(rowbigp, 9).Value
    Range("Q2").Value = Cells(rowbigp, 11).Value
    Range("Q2").NumberFormat = "0.00%"
    Range("P3").Value = Cells(rowsmallp, 9).Value
    Range("Q3").Value = Cells(rowsmallp, 11).Value
    Range("Q3").NumberFormat = "0.00%"
    Range("P4").Value = Cells(rowvol, 9).Value
    Range("Q4").Value = Cells(rowvol, 12).Value
End Sub

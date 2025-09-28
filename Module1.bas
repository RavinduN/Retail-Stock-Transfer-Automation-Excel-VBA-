Attribute VB_Name = "Module1"
Sub CreateTransferPlan_Simplified()
    Dim targetDays As Double: targetDays = 26
    Dim safeDays As Double: safeDays = 14
    Dim minTransfer As Long: minTransfer = 1
    
    Dim wsStock As Worksheet, wsOut As Worksheet
    Set wsStock = ThisWorkbook.Sheets("Stock")
    On Error Resume Next
    Set wsOut = ThisWorkbook.Sheets("TransferPlan")
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Sheets.Add(After:=wsStock)
        wsOut.Name = "TransferPlan"
    End If
    On Error GoTo 0
    
    wsOut.Cells.Clear
    wsOut.Range("A1:E1").Value = Array("ITEM", "DESCRIPTION", "From LOC", "To LOC", "Transfer Qty")
    
    Dim lastRow As Long: lastRow = wsStock.Cells(wsStock.Rows.count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    Dim dictItems As Object: Set dictItems = CreateObject("Scripting.Dictionary")
    Dim i As Long, itm As String
    For i = 2 To lastRow
        itm = Trim(CStr(wsStock.Cells(i, 1).Value))
        If Len(itm) > 0 Then
            If Not dictItems.Exists(itm) Then
                dictItems.Add itm, CStr(wsStock.Cells(i, 2).Value)
            End If
        End If
    Next i
    
    Dim outRow As Long: outRow = 2
    Dim itemCode As Variant
    For Each itemCode In dictItems.Keys
        Dim desc As String: desc = dictItems(itemCode)
        
        Dim locs() As String, stocks() As Double, sales() As Double
        Dim count As Long: count = 0
        For i = 2 To lastRow
            If CStr(wsStock.Cells(i, 1).Value) = itemCode Then
                count = count + 1
                ReDim Preserve locs(1 To count), stocks(1 To count), sales(1 To count)
                locs(count) = CStr(wsStock.Cells(i, 3).Value)
                stocks(count) = Val(wsStock.Cells(i, 4).Value)
                sales(count) = Application.Max(1, Val(wsStock.Cells(i, 5).Value))
            End If
        Next i
        
        If count = 0 Then GoTo NextItem
        
        Dim surplus() As Double, deficit() As Double, holding() As Double
        ReDim surplus(1 To count), deficit(1 To count), holding(1 To count)
        
        For i = 1 To count
            holding(i) = stocks(i) / sales(i)
            surplus(i) = stocks(i) - (safeDays * sales(i))
            If surplus(i) < 0 Then surplus(i) = 0
            deficit(i) = (targetDays * sales(i)) - stocks(i)
            If deficit(i) < 0 Then deficit(i) = 0
        Next i
        
        Dim allBalanced As Boolean
        Do
            allBalanced = True
            
            Dim donor As Long, receiver As Long
            donor = 0: receiver = 0
            
            Dim maxSurplus As Double: maxSurplus = 0
            Dim minHolding As Double: minHolding = 999999
            For i = 1 To count
                If surplus(i) > maxSurplus Then
                    maxSurplus = surplus(i)
                    donor = i
                End If
                If deficit(i) > 0 And holding(i) < minHolding Then
                    minHolding = holding(i)
                    receiver = i
                End If
            Next i
            
            If donor = 0 Or receiver = 0 Then Exit Do
            
            Dim transferQty As Long
            transferQty = Application.Min(surplus(donor), deficit(receiver))
            transferQty = Application.RoundDown(transferQty, 0)
            If transferQty < minTransfer Then Exit Do
            
            wsOut.Cells(outRow, 1).Value = itemCode
            wsOut.Cells(outRow, 2).Value = desc
            wsOut.Cells(outRow, 3).Value = locs(donor)
            wsOut.Cells(outRow, 4).Value = locs(receiver)
            wsOut.Cells(outRow, 5).Value = transferQty
            outRow = outRow + 1
            
            stocks(donor) = stocks(donor) - transferQty
            stocks(receiver) = stocks(receiver) + transferQty
            
            surplus(donor) = stocks(donor) - (safeDays * sales(donor))
            If surplus(donor) < 0 Then surplus(donor) = 0
            deficit(receiver) = (targetDays * sales(receiver)) - stocks(receiver)
            If deficit(receiver) < 0 Then deficit(receiver) = 0
            
            For i = 1 To count
                holding(i) = stocks(i) / sales(i)
            Next i
            
            Dim maxHold As Double, minHold As Double
            maxHold = holding(1): minHold = holding(1)
            For i = 2 To count
                If holding(i) > maxHold Then maxHold = holding(i)
                If holding(i) < minHold Then minHold = holding(i)
            Next i
            If maxHold - minHold > 1 Then allBalanced = False
        Loop Until allBalanced
NextItem:
    Next itemCode
    
    MsgBox "Transfer plan created", vbInformation
End Sub


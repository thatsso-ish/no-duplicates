Attribute VB_Name = "Module1"
'DataTransfer Class Module
Public SrcData As Variant
Public DestCell As Range

Sub RefreshForm()
    ' Unload the current instance of the form
    Unload UserForm2 ' Replace "UserForm1" with the name of your form
    
    ' Load a new instance of the form
    UserForm2.Show ' Replace "UserForm1" with the name of your form
End Sub

Sub DisplayColumnADataOnListBox()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim myUserForm As UserForm2 ' Change to your UserForm name
    
    ' Specify the worksheet where the data is located
    Set ws = ThisWorkbook.Sheets("DOH") ' Change "Sheet1" to your sheet name
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Clear the existing items in the listbox
    UserForm2.ListBox1.Clear ' Change to your listbox name
    
    ' Populate the listbox with data from column A
    For i = 2 To lastRow
        UserForm2.ListBox1.AddItem ws.Cells(i, 1).Value ' Change to your listbox name
    Next i
    
    ' Show the userform
    Set myUserForm = UserForm2 ' Change to your UserForm name
    
End Sub

Sub DisplayColumnADataOnListBox1()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim myUserForm As UserForm2 ' Change to your UserForm name
    
    ' Specify the worksheet where the data is located
    Set ws = ThisWorkbook.Sheets("ADB") ' Change "Sheet1" to your sheet name
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Clear the existing items in the listbox
    UserForm2.ListBox2.Clear ' Change to your listbox name
    
    ' Populate the listbox with data from column A
    For i = 2 To lastRow
        UserForm2.ListBox2.AddItem ws.Cells(i, 1).Value ' Change to your listbox name
    Next i
    
    ' Show the userform
    Set myUserForm = UserForm2 ' Change to your UserForm name
    
End Sub

Sub DisplayColumnADataOnListBox2()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim myUserForm As UserForm2 ' Change to your UserForm name
    
    ' Specify the worksheet where the data is located
    Set ws = ThisWorkbook.Sheets("Re-Write") ' Change "Sheet1" to your sheet name
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Clear the existing items in the listbox
    UserForm2.ListBox3.Clear ' Change to your listbox name
    
    ' Populate the listbox with data from column A
    For i = 2 To lastRow
        UserForm2.ListBox3.AddItem ws.Cells(i, 1).Value ' Change to your listbox name
    Next i
    
    ' Show the userform
    Set myUserForm = UserForm2 ' Change to your UserForm name
    
End Sub

Function DetermineDestinationSheet(result As Double, hivStatus As String, oldRisk As String, newRisk As String) As Worksheet
    Dim wsDestination As Worksheet
    
    If result > 365 And hivStatus <> "UNK" Then
        Set DetermineDestinationSheet = ThisWorkbook.Sheets("Re-Write")
    ElseIf result >= 183 And result < 365 And hivStatus <> "UNK" Then
        Set DetermineDestinationSheet = ThisWorkbook.Sheets("DOH")
    ElseIf oldRisk = "Life" And newRisk = "ADB" And hivStatus = "UNK" Then
        Set DetermineDestinationSheet = ThisWorkbook.Sheets("ADB")
    End If
End Function

Sub ReadDataFromWorkbookAndSendToSheets()
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim columnMappings As Variant
    Dim columnIndex As Variant
    Dim destColumn As Integer
    Dim emptyRow As Boolean
    Dim oldRisk As String
    Dim newRisk As String
    Dim hivStatus As String
    Dim result As Double
    Dim wsDestination As Worksheet
    
    ' Open the source workbook
    Set wbSource = Workbooks.Open("C:\Users\ismae\Downloads\life to adb drop 2.xlsx")
    
    ' Set the source worksheet
    Set wsSource = wbSource.Sheets("life to adb drop 2")
    
    ' Define the column mappings (source column to destination column)
    columnMappings = Array(Array(2, 1), Array(3, 2), Array(5, 3), Array(7, 4), Array(11, 5), Array(12, 6), Array(14, 7), Array(15, 8), Array(16, 10), Array(17, 11), Array(18, 9))
    
    ' Loop through each column mapping
    For Each columnIndex In columnMappings
        ' Extract source and destination column indices
        Dim srcColumn As Integer
        srcColumn = columnIndex(0)
        destColumn = columnIndex(1)
        
        ' Find the last row with data in the selected column
        lastRow = wsSource.Cells(wsSource.Rows.Count, srcColumn).End(xlUp).Row
        
        ' Loop through each row in the selected column and send data to the destination sheet
        For i = 2 To lastRow
            ' Get the values of old risk, new risk, and HIV status
            oldRisk = Trim(wsSource.Cells(i, 11).Value)
            newRisk = Trim(wsSource.Cells(i, 12).Value)
            hivStatus = Trim(wsSource.Cells(i, 14).Value)
            
            ' Check if any of the critical cells are empty
           ' If oldRisk <> "" And newRisk <> "" And hivStatus <> "" Then
                ' Perform the calculation
                result = Abs(wsSource.Cells(i, 16).Value - IIf(IsEmpty(wsSource.Cells(i, 15).Value), Date, wsSource.Cells(i, 15).Value))
                
                ' Determine the relevant destination sheet based on conditions
                Set wsDestination = DetermineDestinationSheet(result, hivStatus, oldRisk, newRisk)
                
                ' If a valid destination sheet is found, transfer the data
                If Not wsDestination Is Nothing Then
                    wsDestination.Cells(i, destColumn).Value = wsSource.Cells(i, srcColumn).Value
                End If
            'End If
        Next i
    Next columnIndex
    
    ' Close the source workbook without saving changes
    wbSource.Close SaveChanges:=False
End Sub

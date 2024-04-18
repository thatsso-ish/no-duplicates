VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   13275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   25395
   OleObjectBlob   =   "search.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click()

End Sub

Sub SearchData(searchValue As String, ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet)
    Dim foundCell As Range
    Dim ws As Worksheet
    
    ' Search for the value in the first sheet
    Set foundCell = ws1.UsedRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    Set ws = ws1
    
    ' If not found in the first sheet, search in the second sheet
    If foundCell Is Nothing Then
        Set foundCell = ws2.UsedRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            Set ws = ws2
        End If
    End If
    
    ' If not found in the first or second sheet, search in the third sheet
    If foundCell Is Nothing Then
        Set foundCell = ws3.UsedRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            Set ws = ws3
        End If
    End If
    
    ' Check if the value was found
    If Not foundCell Is Nothing Then
        ' Update labels, textboxes, etc. with the found data
        With UserForm1
            .Label35.Caption = ws.Cells(foundCell.Row, 1).Value
            .Label33.Caption = ws.Cells(foundCell.Row, 2).Value
            .Label31.Caption = ws.Cells(foundCell.Row, 3).Value
            .Label29.Caption = ws.Cells(foundCell.Row, 4).Value
            .Label27.Caption = ws.Cells(foundCell.Row, 5).Value
            .Label25.Caption = ws.Cells(foundCell.Row, 6).Value
            .Label23.Caption = ws.Cells(foundCell.Row, 7).Value
            .Label21.Caption = ws.Cells(foundCell.Row, 8).Value
            .Label19.Caption = ws.Cells(foundCell.Row, 9).Value
            .Label17.Caption = ws.Cells(foundCell.Row, 10).Value
            .Label15.Caption = ws.Cells(foundCell.Row, 11).Value
            .Label13.Caption = ws.Cells(foundCell.Row, 12).Value
            .ComboBox3.Value = ws.Cells(foundCell.Row, 13).Value
            .Label6.Caption = ws.Cells(foundCell.Row, 14).Value
            .ComboBox1.Value = ws.Cells(foundCell.Row, 15).Value
            .Label4.Caption = ws.Cells(foundCell.Row, 16).Value
            .ComboBox2.Value = ws.Cells(foundCell.Row, 17).Value
            .Label11.Caption = ws.Cells(foundCell.Row, 18).Value
            .TextBox2.Value = ws.Cells(foundCell.Row, 19).Value
            .CheckBox1.Value = ws.Cells(foundCell.Row, 20).Value
        End With
    Else
        ' Clear labels, textboxes, etc. if value not found
        With UserForm1
            .Label35.Caption = " "
            .Label33.Caption = " "
            .Label31.Caption = " "
            .Label29.Caption = " "
            .Label27.Caption = " "
            .Label25.Caption = " "
            .Label23.Caption = " "
            .Label21.Caption = " "
            .Label19.Caption = " "
            .Label17.Caption = " "
            .Label15.Caption = " "
            .Label13.Caption = " "
            .ComboBox3.Value = " "
            .Label6.Caption = " "
            .ComboBox1.Value = " "
            .Label4.Caption = " "
            .ComboBox2.Value = " "
            .Label11.Caption = " "
            .TextBox2.Value = " "
            .CheckBox1.Value = " "
        End With
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim searchTerm As String
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    
    ' Get the search term from the user or from some other source
    searchTerm = TextBox1.Value
    
    ' Set the worksheet objects to the appropriate sheets
    Set ws1 = ThisWorkbook.Sheets("DOH")
    Set ws2 = ThisWorkbook.Sheets("Re-Write")
    Set ws3 = ThisWorkbook.Sheets("ADB")
    
    ' Call the SearchData function with the search term and worksheet objects
    Call UpdateData(TextBox1.Value, ws1, ws2, ws3)
    
    
     Dim ctrl As Control
    
    ' Check if CheckBox1 is checked (value is true)
    If CheckBox1.Value = True Then
        ' Loop through all controls on the UserForm
        For Each ctrl In Me.Controls
            ' Check if the control is a ComboBox
            If TypeName(ctrl) = "ComboBox" Then
                ' Disable the ComboBox
                ctrl.Enabled = False
            End If
        Next ctrl
    Else
        ' If CheckBox1 is unchecked, enable all ComboBoxes
        For Each ctrl In Me.Controls
            If TypeName(ctrl) = "ComboBox" Then
                ctrl.Enabled = True
            End If
        Next ctrl
        
    End If
    

End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Sub UpdateData(searchValue As String, ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet)
    Dim foundCell As Range
    Dim ws As Worksheet
    
    ' Search for the value in the first sheet
    Set foundCell = ws1.UsedRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' If not found in the first sheet, search in the second sheet
    If foundCell Is Nothing Then
        Set foundCell = ws2.UsedRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            Set ws = ws2
        End If
    End If
    
    ' If not found in the first or second sheet, search in the third sheet
    If foundCell Is Nothing Then
        Set foundCell = ws3.UsedRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            Set ws = ws3
        End If
    End If
    
    ' Check if the value was found
    If Not foundCell Is Nothing Then
        ' Update labels, textboxes, etc. with the found data
        With UserForm1
            foundCell.Offset(0, 0).Value = .Label35.Caption
            foundCell.Offset(0, 1).Value = .Label33.Caption
            foundCell.Offset(0, 2).Value = .Label31.Caption
            foundCell.Offset(0, 3).Value = .Label29.Caption
            foundCell.Offset(0, 4).Value = .Label27.Caption
            foundCell.Offset(0, 5).Value = .Label25.Caption
            foundCell.Offset(0, 6).Value = .Label23.Caption
            foundCell.Offset(0, 7).Value = .Label21.Caption
            foundCell.Offset(0, 8).Value = .Label19.Caption
            foundCell.Offset(0, 9).Value = .Label17.Caption
            foundCell.Offset(0, 10).Value = .Label15.Caption
            foundCell.Offset(0, 11).Value = .Label13.Caption
            foundCell.Offset(0, 12).Value = .ComboBox3.Value
            foundCell.Offset(0, 13).Value = .Label6.Caption
            foundCell.Offset(0, 14).Value = .ComboBox1.Value
            foundCell.Offset(0, 15).Value = .Label4.Caption
            foundCell.Offset(0, 16).Value = .ComboBox2.Value
            foundCell.Offset(0, 17).Value = .Label11.Caption
            foundCell.Offset(0, 18).Value = .TextBox2.Value
            foundCell.Offset(0, 19).Value = .CheckBox1.Value
        End With
    End If
End Sub

Private Sub CommandButton3_Click()

    ' Hide Form 2
    Me.Hide
    
    ' Show Form 1
    UserForm2.Show

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub TextBox1_Change()
    Dim searchTerm As String
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    
    ' Set the worksheet objects to the appropriate sheets
    Set ws1 = ThisWorkbook.Sheets("DOH")
    Set ws2 = ThisWorkbook.Sheets("Re-Write")
    Set ws3 = ThisWorkbook.Sheets("ADB")
    
    ' Get the search term from the TextBox
    searchTerm = TextBox1.Value
    
    ' Call the SearchData subroutine with the search term and worksheet objects
     If TextBox1.Value = "" Then
        ClearLabelsAndTextboxes
    Else
        Call SearchData(TextBox1.Value, ws1, ws2, ws3)
    End If
End Sub

Sub ClearLabelsAndTextboxes()
    With UserForm1
        .Label35.Caption = " "
        .Label33.Caption = " "
        .Label31.Caption = " "
        .Label29.Caption = " "
        .Label27.Caption = " "
        .Label25.Caption = " "
        .Label23.Caption = " "
        .Label21.Caption = " "
        .Label19.Caption = " "
        .Label17.Caption = " "
        .Label15.Caption = " "
        .Label13.Caption = " "
        .ComboBox3.Value = " "
        .Label6.Caption = " "
        .ComboBox1.Value = " "
        .Label4.Caption = " "
        .ComboBox2.Value = " "
        .Label11.Caption = " "
        .TextBox2.Value = " "
        .CheckBox1.Value = " "
    End With
End Sub

Sub ExampleUsage()
    Dim searchTerm As String
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    
    ' Get the search term from the user or from some other source
    searchTerm = TextBox1.Value
    
    ' Set the worksheet objects to the appropriate sheets
    Set ws1 = ThisWorkbook.Sheets("DOH")
    Set ws2 = ThisWorkbook.Sheets("Re-Write")
    Set ws3 = ThisWorkbook.Sheets("ADB")
    
    ' Call the SearchData subroutine with the search term and worksheet objects
    Call SearchData(searchTerm, ws1, ws2, ws3)
End Sub


Private Sub UserForm_Activate()

ComboBox1.Style = fmStyleDropDownList
ComboBox2.Style = fmStyleDropDownList
ComboBox3.Style = fmStyleDropDownList

ComboBox1.Clear ' Clear existing items before adding new ones
Me.ComboBox1.AddItem "Call"
Me.ComboBox1.AddItem "Email"
Me.ComboBox1.AddItem "SMS"
Me.ComboBox1.AddItem "Call & Email"
Me.ComboBox1.AddItem "Call & SMS"

ComboBox2.Clear ' Clear existing items before adding new ones
Me.ComboBox2.AddItem "Call"
Me.ComboBox2.AddItem "Email"
Me.ComboBox2.AddItem "SMS"
Me.ComboBox2.AddItem "Call & Email"
Me.ComboBox2.AddItem "Call & SMS"

ComboBox3.Clear ' Clear existing items before adding new ones
Me.ComboBox3.AddItem "Call"
Me.ComboBox3.AddItem "Email"
Me.ComboBox3.AddItem "SMS"
Me.ComboBox3.AddItem "Call & Email"
Me.ComboBox3.AddItem "Call & SMS"

End Sub

Private Sub UserForm_Click()

        DisableAllTextboxes
        TextBox1.Enabled = True

        Label35.Caption = " "
        Label33.Caption = " "
        Label31.Caption = " "
        Label29.Caption = " "
        Label27.Caption = " "
        Label25.Caption = " "
        Label23.Caption = " "
        Label21.Caption = " "
        Label19.Caption = " "
        Label17.Caption = " "
        Label15.Caption = " "
        Label13.Caption = " "
        ComboBox3.Value = " "
        Label6.Caption = " "
        ComboBox1.Value = " "
        Label4.Caption = " "
        ComboBox2.Value = " "
        Label11.Caption = " "
        TextBox2.Value = " "
        CheckBox1.Value = " "

        TextBox2.MultiLine = True
        TextBox2.WordWrap = True
        TextBox2.EnterKeyBehavior = True

End Sub

Private Sub ComboBox1_Change()

' Check if ComboBox1 has a value
    If ComboBox1.Value <> "" Then
        ' Disable ComboBox1
        ComboBox1.Enabled = False
    Else
        ' Enable ComboBox1 if it doesn't have a value
        ComboBox1.Enabled = True
    End If

    ' Check if ComboBox1 is not empty
    If ComboBox1.Value <> "" Then
        ' Execute the action related to ComboBox1
        Label4.Caption = Format(Date, "dd-mmm-yyyy") ' Update Label4 with current date stamp
    Else
        ' Clear Label4 if ComboBox1 is empty
        Label4.Caption = ""
    End If
End Sub

Private Sub ComboBox2_Change()

' Check if ComboBox1 has a value
    If ComboBox2.Value <> "" Then
        ' Disable ComboBox1
        ComboBox2.Enabled = False
    Else
        ' Enable ComboBox1 if it doesn't have a value
        ComboBox2.Enabled = True
    End If

    ' Check if ComboBox2 is not empty
    If ComboBox2.Value <> "" Then
        ' Execute the action related to ComboBox2
        Label11.Caption = Format(Date, "dd-mmm-yyyy") ' Update Label2 with current date stamp
    Else
        ' Clear Label2 if ComboBox2 is empty
        Label11.Caption = ""
    End If
End Sub

Private Sub ComboBox3_Change()

' Check if ComboBox1 has a value
    If ComboBox3.Value <> "" Then
        ' Disable ComboBox1
        ComboBox3.Enabled = False
    Else
        ' Enable ComboBox1 if it doesn't have a value
        ComboBox3.Enabled = True
    End If

' Check if ComboBox2 is not empty
    If ComboBox3.Value <> "" Then
        ' Execute the action related to ComboBox2
        Label6.Caption = Format(Date, "dd-mmm-yyyy") ' Update Label2 with current date stamp
    Else
        ' Clear Label2 if ComboBox2 is empty
        Label6.Caption = ""
    End If
End Sub

Private Sub CheckBox1_Change()
   
End Sub


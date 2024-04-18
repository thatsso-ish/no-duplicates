VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   13215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   25275
   OleObjectBlob   =   "populate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    ' Hide Form 2
    Me.Hide
    
    ' Show Form 1
    UserForm1.Show

End Sub

Private Sub UserForm_Activate()

Call DisplayColumnADataOnListBox
Call DisplayColumnADataOnListBox1
Call DisplayColumnADataOnListBox2

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Check if an item is selected in the listbox
    If Me.ListBox1.ListIndex <> -1 Then
        ' Get the selected item from the listbox
        Dim selectedItem As String
        selectedItem = Me.ListBox1.Value
        
        UserForm1.TextBox1.Value = selectedItem
        ' Hide UserForm2
        Me.Hide
        
        ' Show Form1
        UserForm1.Show
        
        ' Populate TextBox1 in Form1 with the selected item
        
    End If
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Check if an item is selected in the listbox
    If Me.ListBox2.ListIndex <> -1 Then
        ' Get the selected item from the listbox
        Dim selectedItem As String
        selectedItem = Me.ListBox2.Value
        
        UserForm1.TextBox1.Value = selectedItem
        ' Hide UserForm2
        Me.Hide
        
        ' Show Form1
        UserForm1.Show
        
        ' Populate TextBox1 in Form1 with the selected item
        
    End If
End Sub

Private Sub ListBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Check if an item is selected in the listbox
    If Me.ListBox3.ListIndex <> -1 Then
        ' Get the selected item from the listbox
        Dim selectedItem As String
        selectedItem = Me.ListBox3.Value
        
        UserForm1.TextBox1.Value = selectedItem
        ' Hide UserForm2
        Me.Hide
        
        ' Show Form1
        UserForm1.Show
        
        ' Populate TextBox1 in Form1 with the selected item
        
    End If
End Sub

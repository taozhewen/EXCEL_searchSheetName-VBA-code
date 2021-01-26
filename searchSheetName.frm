VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} searchSheetName 
   Caption         =   "表名查询 "
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7650
   OleObjectBlob   =   "searchSheetName.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "searchSheetName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
    TextBox1.Value = ""
End Sub



Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim sht_name As String
    If ListBox1.Value <> "" Then
    sht_name = ListBox1.Value
    Sheets(sht_name).Select
    Else
        MsgBox "请选择表名"
    End If
End Sub

Private Sub TextBox1_Change()
    Dim shet_name, valu(), similar() As String
    Dim i, m, k As Integer
    Dim x As Object
    shet_name = TextBox1.Value
    
    i = 0
    k = 0
    For Each x In Worksheets
        ReDim Preserve valu(i)
        valu(i) = x.Name
        i = i + 1
    Next

    
    
    For m = 0 To i - 1
        If valu(m) Like "*" & shet_name & "*" Then
            ReDim Preserve similar(k)
            similar(k) = valu(m)
            k = k + 1
            
        End If
    Next
    
    If k = 0 Then
        ListBox1.List = valu()
    Else
        ListBox1.List = similar()
    
    End If
End Sub

Private Sub UserForm_Activate()

    Dim valu()
    Dim i As Integer
    Dim x As Object
    
    
    
    
    ListBox1.ColumnCount = 1

    i = 0
    For Each x In Worksheets
        ReDim Preserve valu(i)
        valu(i) = x.Name
        i = i + 1
    Next
    
    
    ListBox1.List = valu()
 
  
End Sub

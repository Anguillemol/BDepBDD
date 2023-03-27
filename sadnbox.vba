Option Explicit

Global motperdu As String


Sub Button1_Click()



Dim onglet_fichier_base As String
Dim workbook_fichier_base As String


workbook_fichier_base = ActiveWorkbook.Name


onglet_fichier_base = "DATA"

Dim onglet_info As String
onglet_info = "Info"



Sheets(onglet_fichier_base).Select

Dim ligne As Long
Dim newrange As Integer
Dim i As Integer
Dim col_codes As Long
newrange = Sheets(onglet_fichier_base).Range("A1").End(xlToRight).Column
For i = 1 To newrange
    If LCase(Sheets(onglet_fichier_base).Cells(1, i).Value) = "codes" Then
        col_codes = i
    End If
Next i

ligne = Sheets(onglet_fichier_base).Cells(Application.Rows.Count, col_codes).End(xlUp).Row + 1

motperdu = "mdptopsecret"

Dim base_data As String
base_data = "Base data.xlsx"

Dim Message, Title, Default, MyValue, mdp

  
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.Item(base_data)
    Dim xRet
    xRet = (Not xWb Is Nothing)
    
    mdp = CStr(Workbooks(workbook_fichier_base).Sheets(onglet_info).Range("AB4").Value)
    
    If xRet = True Then
  
        
        Message = "Entrez le mot de passe Dépôt"    ' Set prompt.
        Title = "Mot de passe Formulaire Région"    ' Set title.

        
'        mdp = "mdp"
        

        
        MyValue = InputBox(Message, Title)
        If MyValue = mdp Then

        
                
            Windows("Base data.xlsx").Activate
            
                Sheets("transco").Select
            Columns("D:D").Select
            Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
                :=Array(1, 1), TrailingMinusNumbers:=True
            Selection.NumberFormat = "0.00"
            Selection.NumberFormat = "0.0"
            Selection.NumberFormat = "0"
            
            Sheets("BI ").Select
            Columns("C:C").Select
            Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
                :=Array(1, 1), TrailingMinusNumbers:=True
            Selection.NumberFormat = "0.00"
            Selection.NumberFormat = "0.0"
            Selection.NumberFormat = "0"
        
            
                
            Windows(workbook_fichier_base).Activate
            Sheets(onglet_fichier_base).Select
            Sheets(onglet_fichier_base).Unprotect Password:="mdptopsecret"
            
            UserForm1.Show
            
            
        ElseIf StrPtr(MyValue) = 0 Then
            Windows(workbook_fichier_base).Activate
            Sheets(1).Select
        Else
            MsgBox "Mot de passe renseigné inccorrect"
            Windows(workbook_fichier_base).Activate
            Sheets(1).Select
        End If
        
    Else


        Message = "Entrez le mot de passe Dépôt"    ' Set prompt.
        Title = "Mot de passe Formulaire Région"    ' Set title.
        'Default = "1"    ' Set default.
        
'        mdp = "mdp"
        

        
        MyValue = InputBox(Message, Title)
        If MyValue = mdp Then
            
            Dim annee As String
            annee = CStr(Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 3).Value)
            
            Workbooks.Open Filename:= _
                "https://kfplc.sharepoint.com/:x:/r/teams/OGRP-Marketingdesventes/Shared%20Documents/Animation%20commerciale%20de%20gamme/" & annee & "/Donn%C3%A9es/Base%20data.xlsx" _
                , UpdateLinks:=0
            
'            Workbooks.Open Filename:= _
'                "https://kfplc-my.sharepoint.com/personal/capelle_d_frbd_kfplc_com/Documents/Documents/Partage/Animco%20F&R/Base%20data.xlsx" _
'                , UpdateLinks:=0
        
                
            Windows("Base data.xlsx").Activate
            
                Sheets("transco").Select
            Columns("D:D").Select
            Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
                :=Array(1, 1), TrailingMinusNumbers:=True
            Selection.NumberFormat = "0.00"
            Selection.NumberFormat = "0.0"
            Selection.NumberFormat = "0"
            
            Sheets("BI ").Select
            Columns("C:C").Select
            Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
                :=Array(1, 1), TrailingMinusNumbers:=True
            Selection.NumberFormat = "0.00"
            Selection.NumberFormat = "0.0"
            Selection.NumberFormat = "0"
        
            
                
            Windows(workbook_fichier_base).Activate
            Sheets(onglet_fichier_base).Select
            Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Unprotect Password:=motperdu
            
            UserForm1.Show
            
            
        ElseIf StrPtr(MyValue) = 0 Then
            Windows(workbook_fichier_base).Activate
            Sheets(1).Select
        Else
            MsgBox "Mot de passe renseigné inccorrect"
            Windows(workbook_fichier_base).Activate
            Sheets(1).Select
        End If


        
    End If




End Sub

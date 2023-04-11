
Option Explicit


Public workbook_compil As String
Public onglet_compil As String

Public emplacement_fichier_suivi As String
Public fichier_suivi As String

Public base_data As String
Public onglet_transco As String
Public onglet_lots As String
Public onglet_derefs As String
Public onglet_souscat As String
Public onglet_pcb As String
Public onglet_rattachement As String
Public onglet_marge As String
Public onglet_code_region As String
Public onglet_suivi As String
Public onglet_BI As String
Public onglet_multifourn As String

Public mois_lettre

Public ligne As Integer
Public nbre_ligne_multifourn As Integer

Public newrange As Integer
Dim col_codes As Integer
Dim col_mois As Integer
Dim col_region As Integer
Dim col_typo As Integer
Dim col_easier As Integer
Dim col_easier_conc As Integer
Dim col_ean As Integer
Dim col_lib As Integer
Dim col_marche As Integer
Dim col_cat As Integer
Dim col_souscat As Integer
Dim col_pcb As Integer
Dim col_minsouh As Integer
Dim col_mintot As Integer
Dim col_nbrefac As Integer
Dim col_minfac As Integer
Dim col_ca As Integer
Dim col_couv As Integer
Dim col_entrepot As Integer

Public ButtonOneClick As Boolean 'Make sure this is before all subs
Public ButtonTwoClick As Boolean


Public lettre_col_codes As String


Private Sub UserForm_Initialize()




workbook_compil = ActiveWorkbook.Name
onglet_compil = "COMPIL"

emplacement_fichier_suivi = "C:\Users\dsi\Documents\"
'filename_fichier_suivi = "C:\Users\dsi\Documents\Fichier_suivi.xlsx"
fichier_suivi = "Fichier_suivi.xlsx"

base_data = "Base data.xlsx"
onglet_transco = "transco"
onglet_lots = "lots"
onglet_derefs = "derefs"
onglet_souscat = "sous cat"
onglet_pcb = "PCB"
onglet_rattachement = "rattachement"
onglet_code_region = "code region"
onglet_suivi = "Suivi Delphine"
onglet_BI = "BI "
onglet_marge = "PV et marge"
onglet_multifourn = "Multifournisseurs"

Dim date_fichier As Date
date_fichier = CDate("01" & "/" & Left(Right(Replace(workbook_compil, ".xlsm", ""), 6), 2) & "/" & Right(Replace(workbook_compil, ".xlsm", ""), 4))
mois_lettre = StrConv(MonthName(Left(Right(Replace(workbook_compil, ".xlsm", ""), 6), 2)), vbProperCase)



newrange = Workbooks(workbook_compil).Sheets(onglet_compil).Range("A1").End(xlToRight).Column

Dim i As Integer
For i = 1 To newrange
    If LCase(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value) = "codes" Or LCase(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value) = "code" Then
        col_codes = i
    ElseIf LCase(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value) = "mois" Then
        col_mois = i
    ElseIf LCase(Replace(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value, "é", "e")) = "region" Then
        col_region = i
    ElseIf LCase(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value) = "typologie" Then
        col_typo = i
    ElseIf LCase(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value) = "easier" Then
        col_easier = i
    ElseIf LCase(Replace(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value, " ", "")) = "easierconcatener" Then
        col_easier_conc = i
    ElseIf LCase(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value) = "ean" Then
        col_ean = i
    ElseIf LCase(Replace(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value, "é", "e")) = "libelle" Then
        col_lib = i
    ElseIf LCase(Replace(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value, "é", "e")) = "marche" Then
        col_marche = i
    ElseIf LCase(Replace(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value, "é", "e")) = "categorie" Then
        col_cat = i
    ElseIf LCase(Replace(Replace(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value, "é", "e"), "-", "")) = "souscategorie" Or LCase(Replace(Replace(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value, "é", "e"), " ", "")) = "souscategorie" Then
        col_souscat = i
    ElseIf LCase(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value) = "pcb" Then
        col_pcb = i
    ElseIf LCase(Replace(Replace(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value, "é", "e"), " ", "")) = "minsouhaite" Then
        col_minsouh = i
    ElseIf LCase(Replace(Replace(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value, "é", "e"), " ", "")) = "mintotregion" Then
        col_mintot = i
    ElseIf LCase(Replace(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value, " ", "")) = "nombrefacingmerch" Then
        col_nbrefac = i
    ElseIf LCase(Replace(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value, " ", "")) = "minfacing" Then
        col_minfac = i
    ElseIf LCase(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value) = "ca" Then
        col_ca = i
    ElseIf LCase(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value) = "couverture" Then
        col_couv = i
    ElseIf LCase(Replace(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(1, i).Value, "ô", "o")) = "entrepots" Then
        col_entrepot = i
    End If
Next i

lettre_col_codes = Split(Cells(1, col_codes).Address, "$")(1)

ligne = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(Application.Rows.Count, col_codes).End(xlUp).Row + 1
nbre_ligne_multifourn = Workbooks(base_data).Sheets(onglet_multifourn).Cells(Application.Rows.Count, col_codes).End(xlUp).Row

Label1.Font.Size = 15
Label2.Font.Size = 10
Label3.Font.Size = 10
Label4.Font.Size = 10
Label5.Font.Size = 10
Label6.Font.Size = 10
Label7.Font.Size = 10

TextBox1.Font.Size = 18
TextBox2.Font.Size = 10
TextBox3.Font.Size = 10
TextBox4.Font.Size = 18
TextBox5.Font.Size = 10

ListBox1.Font.Size = 12
ListBox2.Font.Size = 10



End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    Windows(fichier_suivi).Activate
'    ActiveWorkbook.Save
'    ActiveWindow.Close

Unload Me


End Sub

Private Sub fermer_Click()

'Dim xWb_FS As Workbook
'On Error Resume Next
'Set xWb_FS = Application.Workbooks.Item(fichier_suivi)
'Dim xRet_FS As Boolean
'xRet_FS = (Not xWb_FS Is Nothing)
'
'If xRet_FS = True Then
'        Windows(fichier_suivi).Close SaveChanges:=True
'End If
'
'Dim xWb_NW As Workbook
'On Error Resume Next
'Set xWb_NW = Application.Workbooks.Item(new_workbook)
'Dim xRet_NW As Boolean
'xRet_NW = (Not xWb_NW Is Nothing)
'If xRet_NW = True Then
'    If Workbooks(new_workbook).Sheets(1).Cells(2, 1).Value = "" Then
'        Windows(new_workbook).Close SaveChanges:=False
'    ElseIf Workbooks(new_workbook).Sheets(1).Cells(2, 1).Value = mois_lettre Then
'        Workbooks(new_workbook).SaveAs Filename:=emplacement_fichier_suivi & fichier_suivi, _
'            FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
'        Windows(fichier_suivi).Close
'    End If
'End If
'    Windows(fichier_suivi).Activate
'    ActiveWorkbook.Save
'    ActiveWindow.Close

Unload Me
End Sub
Private Sub TextBox1_Change()
If Len(TextBox1.Value) = 6 And IsNumeric(TextBox1.Value) = True Then
    
    'ligne ou apprait le code
    If Application.WorksheetFunction.CountIf(Range(lettre_col_codes & ":" & lettre_col_codes), TextBox1.Value) > 0 Then
        Dim i As Integer
        For i = 2 To ligne - 1
            If CLng(TextBox1.Value) = Cells(i, col_codes).Value Then
                ListBox1.AddItem i & " - " & Cells(i, col_typo).Values & " - " & Cells(i, col_region).Values
            End If
        Next i
    End If
    
    'n'affiche rien si le code n'est pas présent dans la compil
    If ListBox1.ListCount > 0 Then
            
        'alerte onglet suivi
        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_suivi).Range("H:H"), TextBox1.Value) > 0 Then
            Dim num_ligne As Long
            num_ligne = Workbooks(base_data).Sheets(onglet_suivi).Range("H:H").Find(What:=CLng(TextBox1.Value)).Row
            If Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 8).Value = CLng(TextBox1.Value) And LCase(Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 7).Value) = "oui" Then
                TextBox6.Font.Size = 10
                TextBox6.Value = "Le code " & Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 8).Value & " était le remplacement du code " & Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 3).Value
                TextBox6.BackColor = vbRed
            Else
                TextBox6.Value = ""
            End If
        End If
    
    
        'deref
        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_derefs).Range("G:G"), TextBox1.Value) > 0 Then
            Dim ligne_derefs As Long
            ligne_derefs = Workbooks(base_data).Sheets(onglet_derefs).Cells(Application.Rows.Count, 1).End(xlUp).Row
            Dim var_date As Date
            var_date = CDate("01" & "/" & Left(Right(Replace(workbook_compil, ".xlsm", ""), 6), 2) & "/" & Right(Replace(workbook_compil, ".xlsm", ""), 4))
            Dim var_annee As Long
            var_annee = Year(var_date)
            Dim var_mois_num As Integer
            var_mois_num = Month(var_date)
    
            Dim n As Long
            For n = 2 To ligne_derefs
                If Workbooks(base_data).Sheets(onglet_derefs).Cells(n, 7).Value = CLng(TextBox1.Value) Then
                    If var_annee < Year(Workbooks(base_data).Sheets(onglet_derefs).Cells(n, 5).Value) Then
                        TextBox6.Value = "Article DEREF"
                        TextBox6.Font.Size = 15
                        TextBox6.BackColor = vbRed
                    ElseIf var_annee = Year(Workbooks(base_data).Sheets(onglet_derefs).Cells(n, 5).Value) Then
                        If var_mois_num <= Month(Workbooks(base_data).Sheets(onglet_derefs).Cells(n, 5).Value) Then
                            TextBox6.Value = "Article DEREF"
                            TextBox6.Font.Size = 15
                            TextBox6.BackColor = vbRed
                        End If
                    End If
                End If
            Next n
        End If
        
        'lot/composant
        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_lots).Range("A:A"), TextBox1.Value) > 0 Then
            Dim num_ligne_lot_comp
            num_ligne_lot_comp = Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=CLng(TextBox1.Value)).Row
            If Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_lot_comp, 5).Value = "Lot" Then
                TextBox6.Value = "LOT"
                TextBox6.Font.Size = 15
                TextBox6.BackColor = vbRed
            ElseIf Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_lot_comp, 5).Value = "Composant" Then
                TextBox6.Value = "COMPOSANT"
                TextBox6.Font.Size = 15
                TextBox6.BackColor = vbRed
            End If
        End If
        
        'multifournisseurs
        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_multifourn).Range("A:A"), TextBox1.Value) > 0 Then
            Dim num_ligne_multifourn As Long
            num_ligne_multifourn = Workbooks(base_data).Sheets(onglet_multifourn).Range("A:A").Find(What:=CLng(TextBox1.Value)).Row
            TextBox6.Value = "MULTIFOURNISSEURS"
            '& "    Région : " & Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn, 6).Value
            TextBox6.Font.Size = 10
            TextBox6.BackColor = vbRed
        End If
    End If
    
    
Else
    ListBox1.Clear
    TextBox6.BackColor = vbWhite
    TextBox6.Value = ""
    TextBox6.Font.Size = 10
End If
End Sub


Private Sub TextBox4_Change()
If Len(TextBox4.Value) = 6 And IsNumeric(TextBox4.Value) = True Then
    Dim i As Integer
    'ligne ou apprait le code
    If Application.WorksheetFunction.CountIf(Range(lettre_col_codes & ":" & lettre_col_codes), TextBox4.Value) > 0 Then
        
            
            
        
        For i = 2 To ligne - 1
            
            If CLng(TextBox4.Value) = Cells(i, col_codes).Value Then

                ListBox2.AddItem Cells(i, col_region) & Chr(9) & i & Chr(9) & Cells(i, col_typo)

                
            End If
        Next i
    ElseIf Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_lots).Range("A:A"), TextBox4.Value) > 0 Then
        Dim code_comp_lot As Long
        code_comp_lot = Workbooks(base_data).Sheets(onglet_lots).Cells(Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=CLng(TextBox4.Value)).Row + 1, 1).Value
        
        For i = 2 To ligne - 1
            If code_comp_lot = Cells(i, col_codes).Value Then
                ListBox2.AddItem Cells(i, col_region) & Chr(9) & i & Chr(9) & Cells(i, col_typo)
            End If
        Next i
    
    
    End If


    'alerte onglet suivi
    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_suivi).Range("C:C"), TextBox4.Value) > 0 Then
        Dim num_ligne As Long
        num_ligne = Workbooks(base_data).Sheets(onglet_suivi).Range("C:C").Find(What:=CLng(TextBox4.Value)).Row
        If Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 3).Value = CLng(TextBox4.Value) And LCase(Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 7).Value) = "alerte" Then
            TextBox5.Value = Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 5).Value
        ElseIf Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 3).Value = CLng(TextBox4.Value) And LCase(Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 7).Value) = "oui" Then
            TextBox5.Font.Size = 10
            TextBox5.Value = "Ce code " & Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 3).Value & " va être remplacé par le code " & Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 8).Value
            TextBox5.BackColor = vbRed
        Else
            TextBox5.Value = ""
        End If
    End If


    'deref
    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_derefs).Range("G:G"), TextBox4.Value) > 0 Then
        Dim ligne_derefs As Long
        ligne_derefs = Workbooks(base_data).Sheets(onglet_derefs).Cells(Application.Rows.Count, 1).End(xlUp).Row
        Dim var_date As Date
        var_date = CDate("01" & "/" & Left(Right(Replace(workbook_compil, ".xlsm", ""), 6), 2) & "/" & Right(Replace(workbook_compil, ".xlsm", ""), 4))
        Dim var_annee As Long
        var_annee = Year(var_date)
        Dim var_mois_num As Integer
        var_mois_num = Month(var_date)

        Dim n As Long
        For n = 2 To ligne_derefs
            If Workbooks(base_data).Sheets(onglet_derefs).Cells(n, 7).Value = CLng(TextBox4.Value) Then
                If var_annee < Year(Workbooks(base_data).Sheets(onglet_derefs).Cells(n, 5).Value) Then
                    TextBox5.Value = "Article DEREF"
                    TextBox5.Font.Size = 15
                    TextBox5.BackColor = vbRed
                ElseIf var_annee = Year(Workbooks(base_data).Sheets(onglet_derefs).Cells(n, 5).Value) Then
                    If var_mois_num <= Month(Workbooks(base_data).Sheets(onglet_derefs).Cells(n, 5).Value) Then
                        TextBox5.Value = "Article DEREF"
                        TextBox5.Font.Size = 15
                        TextBox5.BackColor = vbRed
                    End If
                End If
            End If
        Next n
    End If
    
    'lot/composant
    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_lots).Range("A:A"), TextBox4.Value) > 0 Then
        Dim num_ligne_lot_comp
        num_ligne_lot_comp = Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=CLng(TextBox4.Value)).Row
        If Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_lot_comp, 5).Value = "Lot" Then
            TextBox5.Value = "LOT"
            TextBox5.Font.Size = 15
            TextBox5.BackColor = vbRed
        ElseIf Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_lot_comp, 5).Value = "Composant" Then
            TextBox5.Value = "COMPOSANT"
            TextBox5.Font.Size = 15
            TextBox5.BackColor = vbRed
        End If
    End If
    
    'multifournisseurs
    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_multifourn).Range("A:A"), TextBox4.Value) > 0 Then
        Dim num_ligne_multifourn As Long
        num_ligne_multifourn = Workbooks(base_data).Sheets(onglet_multifourn).Range("A:A").Find(What:=CLng(TextBox4.Value)).Row
        TextBox5.Value = "MULTIFOURNISSEURS"
        '& "    " Région : " & Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn, 6).Value
        TextBox5.Font.Size = 10
        TextBox5.BackColor = vbRed
    End If
    
Else
    TextBox5.BackColor = vbWhite
    TextBox5.Value = ""
    TextBox5.Font.Size = 10
    ListBox2.Clear

End If
End Sub

Sub ajout_data(ligne_select, old_typo_sub, old_region_sub, old_minsouh_sub, old_mintot_sub, old_nbrefac_sub, old_minfac_sub, old_entrepot_sub)

    'variable ca engage + couverture
    Dim code_ean As Double
    Dim prmp As Double
    Dim nbre_depot As Double
    Dim min_souh As Double
    Dim somme As Double
    Dim prix_vente As Double




                'mois
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_mois) = mois_lettre
                'region
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_region) = old_region_sub
                'typo
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_typo) = old_typo_sub
                 'easier/ean/lib
                If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_transco).Range("B:B"), Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_codes)) > 0 Then
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_easier) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_codes)), Workbooks(base_data).Sheets(onglet_transco).Range("B:C"), 2, 0)
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_ean) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_codes)), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_lib) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_codes)), Workbooks(base_data).Sheets(onglet_transco).Range("B:N"), 13, 0)
                Else
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_easier) = "-"
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_ean) = "-"
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_lib) = "-"
                End If
                'easier concatener
                If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_easier) = "-" Then
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_easier_conc) = "-"
                Else
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_easier_conc) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_easier) & "EA"
                End If
                 'marche/cat/souscat
                If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_souscat).Range("A:A"), Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_codes)) > 0 Then
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_marche) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_codes)), Workbooks(base_data).Sheets(onglet_souscat).Range("A:D"), 4, 0)
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_cat) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_codes)), Workbooks(base_data).Sheets(onglet_souscat).Range("A:E"), 5, 0)
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_souscat) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_codes)), Workbooks(base_data).Sheets(onglet_souscat).Range("A:F"), 5, 0)
                Else
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_marche) = "-"
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_cat) = "-"
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_souscat) = "-"
                End If
                 'PCB
                If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_pcb).Range("B:B"), Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_codes)) > 0 Then
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_pcb) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_codes)), Workbooks(base_data).Sheets(onglet_pcb).Range("B:J"), 9, 0)
                Else
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_pcb) = "-"
                End If
                'minsouh
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_minsouh) = old_minsouh_sub
                'mintot
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_mintot) = old_mintot_sub
                'nbrefac
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_nbrefac) = old_nbrefac_sub
                'minfac
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_minfac) = old_minfac_sub
                'CA engagé
                If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_marge).Range("C:C"), Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_codes)) > 0 Then
                    prix_vente = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_codes)), Workbooks(base_data).Sheets(onglet_marge).Range("C:D"), 2, 0)
                    If prix_vente <> 0 Then
                        Cells(ligne_select, col_ca) = prix_vente * min_souh
                    Else
                        Cells(ligne_select, col_ca) = "-"
                    End If
                End If
                'couverture
                If Application.WorksheetFunction.IfError(Application.VLookup(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_codes), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0), 0) Then
                    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_BI).Range("C:C"), Application.WorksheetFunction.VLookup(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_codes), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)) > 0 Then

                        code_ean = Application.WorksheetFunction.VLookup(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_codes), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)
                        min_souh = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_minsouh)
                        prmp = Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("C:F"), 4, 0)
                        nbre_depot = Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("C:G"), 5, 0)
                        somme = Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("C:T"), 18, 0)
                        If prmp <> 0 Then
                            If somme <> 0 Or nbre_depot <> 0 Then
                                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_couv) = Round((min_souh * prmp) / (somme * prmp) / (nbre_depot) * 12, 2)
                            Else
                                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_couv) = "-"
                            End If
                        Else
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_couv) = "-"
                        End If
                    Else
                        Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_couv) = "-"
                    End If
                Else
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_couv) = "-"
                End If
                'entrepot
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_select, col_entrepot) = old_entrepot_sub



End Sub



Sub ajout_data_remplacer_suivi(file_suivi_remplacer, ligne_suivi_remplacer, ligne_modif_remplacer, old_lib_remplacer, old_code_remplacer, nom_remplacer, com_remplacer)

    'alimentation fichier suivi
'    ligne_suivi_remplacer = Workbooks(file_suivi_remplacer).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
    Workbooks(file_suivi_remplacer).Sheets(1).Cells(ligne_suivi_remplacer, 1).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_remplacer, col_mois).Value
    Workbooks(file_suivi_remplacer).Sheets(1).Cells(ligne_suivi_remplacer, 2).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_remplacer, col_region).Value
    Workbooks(file_suivi_remplacer).Sheets(1).Cells(ligne_suivi_remplacer, 3).Value = old_code_remplacer
    Workbooks(file_suivi_remplacer).Sheets(1).Cells(ligne_suivi_remplacer, 4).Value = old_lib_remplacer
    Workbooks(file_suivi_remplacer).Sheets(1).Cells(ligne_suivi_remplacer, 5).Value = com_remplacer
    Workbooks(file_suivi_remplacer).Sheets(1).Cells(ligne_suivi_remplacer, 6).Value = "remplacé"
    Workbooks(file_suivi_remplacer).Sheets(1).Cells(ligne_suivi_remplacer, 7).Value = "oui"
    Workbooks(file_suivi_remplacer).Sheets(1).Cells(ligne_suivi_remplacer, 8).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_remplacer, col_codes).Value
    Workbooks(file_suivi_remplacer).Sheets(1).Cells(ligne_suivi_remplacer, 9).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_remplacer, col_lib)
    Workbooks(file_suivi_remplacer).Sheets(1).Cells(ligne_suivi_remplacer, 10).Value = nom_remplacer


End Sub


Sub new_code_suivi(code As Long, ligne_modif_usf As Long, file_suivi As String)



    Dim num_ligne
            
    Dim ligne_suivi As Long
            
            
            
    Dim old_lib As String
    old_lib = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_lib)
    Dim old_typo As String
    old_typo = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_typo)
    Dim old_region As String
    old_region = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_region)
    Dim old_minsouh As Long
    old_minsouh = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_minsouh)
    Dim old_mintot As Long
    old_mintot = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_mintot)
    Dim old_nbrefac As Long
    old_nbrefac = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_nbrefac)
    Dim old_minfac As Long
    old_minfac = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_minfac)
    Dim old_entrepot As String
    old_entrepot = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_entrepot)


Dim flag_remplacement As Integer
Dim ligne_remplacement As Long
If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_suivi).Range("C:C"), code) > 0 Then
    Dim num_code_suivi As Long
    num_code_suivi = Workbooks(base_data).Sheets(onglet_suivi).Range("C:C").Find(What:=code).Row
    If Workbooks(base_data).Sheets(onglet_suivi).Cells(num_code_suivi, 3).Value = CLng(code) And LCase(Workbooks(base_data).Sheets(onglet_suivi).Cells(num_code_suivi, 7).Value) = "oui" Then
        flag_remplacement = 1
        ligne_remplacement = num_code_suivi
    End If
End If


Dim flag_fourn As Integer
If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_multifourn).Range("A:A"), code) > 0 Then
    Dim i_multifourn As Long
    For i_multifourn = 2 To nbre_ligne_multifourn
        If Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 1).Value = CLng(code) And Replace(UCase(CStr(Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 6).Value)), " ", "") = Replace(old_region, " ", "") Then
            Dim ligne_multifourn
            ligne_multifourn = i_multifourn
            flag_fourn = 1
        End If
    Next i_multifourn
End If


    'FOURNISSEUR
    If flag_fourn = 1 Then
        MsgBox "MULTIFOURNISSEUR"

            Dim num_ligne_multifourn As Long
            num_ligne_multifourn = ligne_multifourn

            Rows(ligne_modif_usf & ":" & ligne_modif_usf).delete

            'multifournisseur - 1
            Dim num_ligne_multifourn_1 As Long
            num_ligne_multifourn_1 = num_ligne_multifourn
            While Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_1, 5).Value = "fournisseur"

               'Modification codes onglet compil
                Rows(ligne_modif_usf & ":" & ligne_modif_usf).Insert Shift:=xlDown
                'codes
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes) = Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_1, 1).Value
                Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ligne_modif_usf).Font.Color = -65434 'violet
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).AddComment
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).Comment.Text Text:="Fournisseur"
                
                'alimentation fichier compil
                Call ajout_data(ligne_modif_usf, old_typo, old_region, old_minsouh, old_mintot, old_nbrefac, old_minfac, old_entrepot)


                'alimentation fichier suivi
                ligne_suivi = Workbooks(file_suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
                Call ajout_data_remplacer_suivi(file_suivi, ligne_suivi, ligne_modif_usf, old_lib, TextBox1.Value, TextBox2.Value, TextBox3.Value)
                
                Windows(workbook_compil).Activate


                ligne_modif_usf = ligne_modif_usf + 1
                num_ligne_multifourn_1 = num_ligne_multifourn_1 - 1
            Wend


            'NOUVEAU Fournisseur + 1
            Dim num_ligne_multifourn_plus_1 As Long
            num_ligne_multifourn_plus_1 = num_ligne_multifourn
            While Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_plus_1 + 1, 5).Value = "fournisseur"

               'Modification codes onglet compil
                Rows(ligne_modif_usf & ":" & ligne_modif_usf).Insert Shift:=xlDown
                'codes
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes) = Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_plus_1 + 1, 1).Value
                Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ligne_modif_usf).Font.Color = -65434 'violet
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).AddComment
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).Comment.Text Text:="Fournisseur"
                
                'alimentation fichier compil
                Call ajout_data(ligne_modif_usf, old_typo, old_region, old_minsouh, old_mintot, old_nbrefac, old_minfac, old_entrepot)

                'alimentation fichier suivi
                ligne_suivi = Workbooks(file_suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
                Call ajout_data_remplacer_suivi(file_suivi, ligne_suivi, ligne_modif_usf, old_lib, TextBox1.Value, TextBox2.Value, TextBox3.Value)

                Windows(workbook_compil).Activate


                ligne_modif_usf = ligne_modif_usf + 1
                num_ligne_multifourn_plus_1 = num_ligne_multifourn_plus_1 + 1
            Wend


    'check lot/composant
    ElseIf Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_lots).Range("A:A"), code) > 0 Then

        num_ligne = Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=code).Row

        'NOUVEAU CODE LOT
        If Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne, 5).Value = "Lot" Then

            MsgBox "LOT"

            Rows(ligne_modif_usf & ":" & ligne_modif_usf).delete
            While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne + 1, 5).Value = "Composant"

               'Modification codes onglet compil
                Rows(ligne_modif_usf & ":" & ligne_modif_usf).Insert Shift:=xlDown
                'codes
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne + 1, 1).Value
                Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ligne_modif_usf).Font.Color = -65281 'pink
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).AddComment
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).Comment.Text Text:="Code Composant"
                
                'alimentation fichier compil
                Call ajout_data(ligne_modif_usf, old_typo, old_region, old_minsouh, old_mintot, old_nbrefac, old_minfac, old_entrepot)

                'alimentation fichier suivi
                ligne_suivi = Workbooks(file_suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
                Call ajout_data_remplacer_suivi(file_suivi, ligne_suivi, ligne_modif_usf, old_lib, TextBox1.Value, TextBox2.Value, TextBox3.Value)

                Windows(workbook_compil).Activate



                ligne_modif_usf = ligne_modif_usf + 1
                num_ligne = num_ligne + 1
            Wend

        'NOUVEAU CODE COMPOSANT
        ElseIf Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne, 5).Value = "Composant" Then
            MsgBox "Composant"

            Rows(ligne_modif_usf & ":" & ligne_modif_usf).delete

            'composant - 1
            Dim num_ligne_1 As Long
            num_ligne_1 = num_ligne
            While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_1, 5).Value = "Composant"

               'Modification codes onglet compil
                Rows(ligne_modif_usf & ":" & ligne_modif_usf).Insert Shift:=xlDown
                'codes
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_1, 1).Value
                Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ligne_modif_usf).Font.Color = -65281 'pink
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).AddComment
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).Comment.Text Text:="Code Composant"
                
                'alimentation fichier compil
                Call ajout_data(ligne_modif_usf, old_typo, old_region, old_minsouh, old_mintot, old_nbrefac, old_minfac, old_entrepot)

                'alimentation fichier suivi
                ligne_suivi = Workbooks(file_suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
                Call ajout_data_remplacer_suivi(file_suivi, ligne_suivi, ligne_modif_usf, old_lib, TextBox1.Value, TextBox2.Value, TextBox3.Value)

                Windows(workbook_compil).Activate


                ligne_modif_usf = ligne_modif_usf + 1
                num_ligne_1 = num_ligne_1 - 1
            Wend


            'NOUVEAU CODE composant + 1
            Dim num_ligne_plus_1 As Long
            num_ligne_plus_1 = num_ligne
            While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_plus_1 + 1, 5).Value = "Composant"

               'Modification codes onglet compil
                Rows(ligne_modif_usf & ":" & ligne_modif_usf).Insert Shift:=xlDown
                'codes
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_plus_1 + 1, 1).Value
                Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ligne_modif_usf).Font.Color = -65281 'pink
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).AddComment
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).Comment.Text Text:="Code Composant"
                
                
                'alimentation fichier compil
                Call ajout_data(ligne_modif_usf, old_typo, old_region, old_minsouh, old_mintot, old_nbrefac, old_minfac, old_entrepot)

                'alimentation fichier suivi
                ligne_suivi = Workbooks(file_suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
                Call ajout_data_remplacer_suivi(file_suivi, ligne_suivi, ligne_modif_usf, old_lib, TextBox1.Value, TextBox2.Value, TextBox3.Value)

                Windows(workbook_compil).Activate


                ligne_modif_usf = ligne_modif_usf + 1
                num_ligne_plus_1 = num_ligne_plus_1 + 1
            Wend

        End If



    Else
        'check suivi delphine
        If flag_remplacement >= 1 Then

            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes) = Workbooks(base_data).Sheets(onglet_suivi).Cells(ligne_remplacement, 8).Value


            Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ligne_modif_usf).Font.Color = -65281 'pink
            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).AddComment
            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).Comment.Text Text:="Le code " & code & " a été remplacé par le code " & Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).Value


        'nouveau code deref
        ElseIf LCase(TextBox5.Value) = "article deref" Then


            If MsgBox("Le nouveau code " & code & " est un article Deref. Souhaitez-vous l'ajouter quand même?", vbYesNo, "confirmation") = vbYes Then
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes) = code
                Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ligne_modif_usf).Font.Color = -65281 'pink
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).AddComment
                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).Comment.Text Text:="Article DEREF"
            Else
                Unload UserForm1
'                UserForm1.Show

            End If
        'code "normal"
        Else
            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes) = code
            Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ligne_modif_usf).Font.ColorIndex = xlAutomatic
            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_usf, col_codes).ClearComments
        End If

        'alimentation fichier compil
        Call ajout_data(ligne_modif_usf, old_typo, old_region, old_minsouh, old_mintot, old_nbrefac, old_minfac, old_entrepot)

        'alimentation fichier suivi
        ligne_suivi = Workbooks(file_suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
        Call ajout_data_remplacer_suivi(file_suivi, ligne_suivi, ligne_modif_usf, old_lib, TextBox1.Value, TextBox2.Value, TextBox3.Value)

        Windows(workbook_compil).Activate

        
        
    End If


End Sub


Private Sub remplacer_ligne_Click()

'Dim xWb As Workbook
'On Error Resume Next
'Set xWb = Application.Workbooks.Item(fichier_suivi)
'Dim xret As Boolean
'xret = (Not xWb Is Nothing)
Dim suivi As String
'If xret = True Then
'    suivi = fichier_suivi
'Else
'    suivi = new_workbook
'End If
suivi = fichier_suivi
        

        If WorksheetFunction.CountIf(Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes), Me.TextBox1.Value) = 0 Then
            MsgBox "Ce code ne peut pas être modifié car n'est pas encore renseigné"
        ElseIf IsNumeric(TextBox4.Value) = False And Len(TextBox4.Value) < 6 Then
            MsgBox "Veuillez rentrez un code convenable"
        ElseIf TextBox2.Value = "" Then
            MsgBox "Veuillez saisir votre nom dans la case Nom"
        ElseIf TextBox3.Value = "" Then
            MsgBox "Veuillez saisir votre commentaire dans la case Commentaire"
        Else

            Dim j As Integer
            Dim cpt As Integer
            cpt = 0


            
        
            For j = 0 To ListBox1.ListCount - 1
                If ListBox1.Selected(j) = True Then
                
                        Dim ligne_modif As Long
                        ligne_modif = ListBox1.List(j, 0)
                        pos = InStr(ligne_modif, " ")
                        ligne_modif = Left(ligne_modif, pos - 1)
                        ligne_modif = CInt(ligne_modif)
                        
                        'cas code flag siege
                        If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, newrange) = 1 Then
                            Dim choix

                            choix = MsgBox("Le code " & TextBox1.Value & " a été ajouté par le siège à la ligne " & ligne_modif & " et pour région " & Cells(ligne_modif, col_region) & " . Souhaitez vous le remplacer?", 36, "Confirmation")
                            If choix = vbNo Then
                                Exit Sub
                            End If
                        End If
                        
                        
                        If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_codes).Font.ColorIndex <> xlAutomatic Then
                            
                            'COMPOSANT
                            If LCase(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_codes).Comment.Text) = "code composant" Then
                            
                                Dim var_region As String
                                var_region = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_region)
    
                                Dim var_typo As String
                                var_typo = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_typo)
                                
                                Dim var_flag
                                var_flag = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, newrange)
    
                                Dim num_ligne_comp As Long
                                num_ligne_comp = Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=CLng(TextBox1.Value)).Row
                                
                                Dim ligne_tot As Long
                                
    
                                Dim x_1 As Long
    
                                Dim num_ligne_comp_1 As Long
                                num_ligne_comp_1 = num_ligne_comp
                                While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_1 - 1, 5).Value = "Composant"
                                    ligne_tot = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                                    For x_1 = 2 To ligne_tot
                                        If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_1, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_1 - 1, 1).Value And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_1, col_region) = var_region And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_1, col_typo) = var_typo And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_1, newrange) = var_flag Then
    
                                            Rows(x_1 & ":" & x_1).EntireRow.delete
    
                                        End If
                                    Next x_1
                                    num_ligne_comp_1 = num_ligne_comp_1 - 1
                                Wend
                                
    
                                
                                Dim x_plus_1 As Long
    
                                Dim num_ligne_comp_plus_1 As Long
                                num_ligne_comp_plus_1 = num_ligne_comp
                                While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_plus_1 + 1, 5).Value = "Composant"
                                    ligne_tot = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                                    For x_plus_1 = 2 To ligne_tot
                                        If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_plus_1, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_plus_1 + 1, 1).Value And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_plus_1, col_region) = var_region And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_plus_1, col_typo) = var_typo Then
    
                                            Rows(x_plus_1 & ":" & x_plus_1).EntireRow.delete
                                            
                                        End If
                                    Next x_plus_1
                                    num_ligne_comp_plus_1 = num_ligne_comp_plus_1 + 1
                                Wend
                                
                                Dim last_ligne_cas_comp As Long
                                last_ligne_cas_comp = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                                Dim ligne_modif_cas_comp As Long
                                Dim x As Long
                                For x = 2 To last_ligne_cas_comp
                                    If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x, col_codes) = CLng(TextBox1.Value) And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x, col_region) = var_region And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x, col_typo) = var_typo Then
                                       ligne_modif_cas_comp = x
                                    End If
                                Next x
                                
    '                            MsgBox ligne_modif_cas_comp
                                Call new_code_suivi(TextBox4.Value, ligne_modif_cas_comp, suivi)
                            
                            'MULTIFOURNISSEUR
                            ElseIf LCase(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_codes).Comment.Text) = "fournisseur" Then
                                Dim var_region_multifourn As String
                                var_region_multifourn = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_region)
    
                                Dim var_typo_multifourn As String
                                var_typo_multifourn = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_typo)
    
                                Dim num_ligne_multifourn As Long
                                Dim i_multifourn As Long
                                For i_multifourn = 2 To nbre_ligne_multifourn
                                    If Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 1) = CLng(TextBox1.Value) And Replace(UCase(CStr(Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 6).Value)), " ", "") = Replace(var_region_multifourn, " ", "") Then
                                        num_ligne_multifourn = i_multifourn
                                    End If
                                Next i_multifourn

                                Dim ligne_tot_multifourn As Long
    
                                Dim x_1_multifourn As Long

                                Dim num_ligne_multifourn_1 As Long
                                num_ligne_multifourn_1 = num_ligne_multifourn
                                While LCase(Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_1 - 1, 5).Value) = "fournisseur"
                                    ligne_tot_multifourn = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                                    For x_1_multifourn = 2 To ligne_tot_multifourn
                                        If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_1_multifourn, col_codes) = Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_1 - 1, 1).Value And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_1_multifourn, col_region) = var_region_multifourn And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_1_multifourn, col_typo) = var_typo_multifourn Then

                                            Rows(x_1_multifourn & ":" & x_1_multifourn).EntireRow.delete

                                        End If
                                    Next x_1_multifourn
                                    num_ligne_multifourn_1 = num_ligne_multifourn_1 - 1
                                Wend



                                Dim x_plus_1_multifourn As Long

                                Dim num_ligne_multifourn_plus_1 As Long
                                num_ligne_multifourn_plus_1 = num_ligne_multifourn
                                While LCase(Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_plus_1 + 1, 5).Value) = "fournisseur"
                                    ligne_tot_multifourn = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                                    For x_plus_1_multifourn = 2 To ligne_tot_multifourn
                                        If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_plus_1_multifourn, col_codes) = Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_plus_1 + 1, 1).Value And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_plus_1_multifourn, col_region) = var_region_multifourn And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_plus_1_multifourn, col_typo) = var_typo_multifourn Then

                                            Rows(x_plus_1_multifourn & ":" & x_plus_1_multifourn).EntireRow.delete

                                        End If
                                    Next x_plus_1_multifourn
                                    num_ligne_multifourn_plus_1 = num_ligne_multifourn_plus_1 + 1
                                Wend

                                Dim last_ligne_multifourn As Long
                                last_ligne_multifourn = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                                Dim ligne_modif_cas_multifourn As Long
                                Dim x_multifourn As Long
                                For x_multifourn = 2 To last_ligne_multifourn
                                    If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_multifourn, col_codes) = CLng(TextBox1.Value) And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_multifourn, col_region) = var_region_multifourn And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_multifourn, col_typo) = var_typo_multifourn Then
                                       ligne_modif_cas_multifourn = x_multifourn
                                    End If
                                Next x_multifourn

                                Call new_code_suivi(TextBox4.Value, ligne_modif_cas_multifourn, suivi)
                            
                            ElseIf LCase(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_codes).Comment.Text) = "article deref" Then
                                MsgBox "deref"
                                Call new_code_suivi(TextBox4.Value, ligne_modif, suivi)
                                
                            
                            Else
                                MsgBox "suivi"
                                Call new_code_suivi(TextBox4.Value, ligne_modif, suivi)
                            
                            End If
                        'ancien code NOIR
                        Else
                            
                            
                            Call new_code_suivi(TextBox4.Value, ligne_modif, suivi)
                                
                        End If
            
                        'Remplacement du code

                        'ligne_modif et nouveau code à remplacer

                        num_ligne = Workbooks(base_data).Sheets(onglet_suivi).Range("H:H").Find(What:=CLng(TextBox1.Value)).Row
                        ancien_code = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_codes).Value
                        code_remplacement = Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 8).Value

                        'Modif du code
                        Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_codes).AddComment "Le code " + ancien_code + "a été remplacé"
                        Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_codes).Value = code_remplacement
'                    End If
                Else
                    cpt = cpt + 1
                End If

            Next j
            
            If cpt = ListBox1.ListCount Then
                MsgBox "veuillez selectionné un item"
            Else
                Windows(workbook_compil).Activate
                Unload UserForm1
                UserForm1.Show
            
            End If
            

        End If

      
End Sub


Public Sub check_flag_siege_tout(code)
Dim i As Long
For i = 2 To ligne - 1
    If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_codes) = CLng(code) And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, newrange) = 1 Then
        Dim choix

        choix = MsgBox("Le code a été ajouté par le siège. " & i & " Souhaitez vous le supprimer?", 36, "Confirmation")
        If choix = vbNo Then
            Exit Sub
        End If
    End If
    
Next i
End Sub

Private Sub remplacer_tout_Click()

'Dim xWb As Workbook
'On Error Resume Next
'Set xWb = Application.Workbooks.Item(fichier_suivi)
'Dim xret As Boolean
'xret = (Not xWb Is Nothing)
Dim suivi As String
'If xret = True Then
'    suivi = fichier_suivi
'Else
'    suivi = new_workbook
'End If
suivi = fichier_suivi




        If WorksheetFunction.CountIf(Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes), Me.TextBox1.Value) = 0 Then
            MsgBox "Ce code ne peut pas être modifié car n'est pas encore renseigné"
        ElseIf IsNumeric(TextBox4.Value) = False And Len(TextBox4.Value) < 6 Then
            MsgBox "Veuillez rentrez un code convenable"
        ElseIf TextBox2.Value = "" Then
            MsgBox "Veuillez saisir votre nom dans la case Nom"
        ElseIf TextBox3.Value = "" Then
            MsgBox "Veuillez saisir votre commentaire dans la case Commentaire"
        ElseIf TextBox1.Value = TextBox4.Value Then
            MsgBox "Veuillez mettre un nouveau code différent de l'ancien de code"
        Else
            
            Dim i As Long
            For i = 2 To ligne - 1
                If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_codes) = CLng(TextBox1.Value) And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, newrange) = 1 Then
                    Dim choix
            
                    choix = MsgBox("Le code a été ajouté par le siège à la ligne " & i & " et pour région " & Cells(i, col_region) & " . Souhaitez vous le supprimer?", 36, "Confirmation")
                    If choix = vbNo Then
                        Exit Sub
                    End If
                End If
                
            Next i
        
            If WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_lots).Range("A:A"), Me.TextBox1.Value) > 0 Then
                MsgBox "COMP"
                
                Dim num_ligne_comp As Long
                num_ligne_comp = Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=CLng(TextBox1.Value)).Row
    
                Dim ligne_tot As Long
    
                Dim x_1 As Long
    
                Dim num_ligne_comp_1 As Long
                num_ligne_comp_1 = num_ligne_comp
                While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_1 - 1, 5).Value = "Composant"
                    While WorksheetFunction.CountIf(Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes), Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_1 - 1, 1).Value) > 0
                            x_1 = Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes).Find(What:=CLng(Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_1 - 1, 1).Value)).Row
                            Rows(x_1 & ":" & x_1).EntireRow.delete
                    Wend
                    num_ligne_comp_1 = num_ligne_comp_1 - 1
                Wend
    
    
                Dim x_plus_1 As Long
    
                Dim num_ligne_comp_plus_1 As Long
                num_ligne_comp_plus_1 = num_ligne_comp
                While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_plus_1 + 1, 5).Value = "Composant"
                    While WorksheetFunction.CountIf(Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes), Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_plus_1 + 1, 1).Value) > 0
                            x_plus_1 = Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes).Find(What:=CLng(Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_plus_1 + 1, 1).Value)).Row
                            Rows(x_plus_1 & ":" & x_plus_1).EntireRow.delete
                    Wend
                    num_ligne_comp_plus_1 = num_ligne_comp_plus_1 + 1
                Wend
            
                
            ElseIf WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_multifourn).Range("A:A"), Me.TextBox1.Value) > 0 Then
                MsgBox "FOURNISSEUR" & Chr(13) & Chr(10) & "Veuillez utiliser le bouton remplacer ligne pour un code Fournisseur"
                
                
                
            End If
            






            While WorksheetFunction.CountIf(Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes), Me.TextBox1.Value) > 0
                Dim num_ligne_compil As Long
                num_ligne_compil = Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes).Find(What:=CLng(TextBox1.Value)).Row

                Call new_code_suivi(TextBox4.Value, num_ligne_compil, suivi)
            Wend
            
            Unload UserForm1
            UserForm1.Show



    End If











End Sub


Private Sub supprimer_ligne_Click()

'Dim xWb As Workbook
'On Error Resume Next
'Set xWb = Application.Workbooks.Item(fichier_suivi)
'Dim xret As Boolean
'xret = (Not xWb Is Nothing)
Dim suivi As String
'If xret = True Then
'    suivi = fichier_suivi
'Else
'    suivi = new_workbook
'End If
suivi = fichier_suivi

        If WorksheetFunction.CountIf(Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes), Me.TextBox1.Value) = 0 Then
            MsgBox "Ce code ne peut pas être modifié car n'est pas encore renseigné"
        ElseIf TextBox2.Value = "" Then
            MsgBox "Veuillez saisir votre nom dans la case Nom"
        ElseIf TextBox3.Value = "" Then
            MsgBox "Veuillez saisir votre commentaire dans la case Commentaire"
        Else

            Dim j As Integer
            Dim cpt As Integer
            cpt = 0
            
            Dim ligne_suivi As Long


            
        
            For j = 0 To ListBox1.ListCount - 1
                If ListBox1.Selected(j) = True Then
                
                    Dim ligne_modif As Long
                    ligne_modif = ListBox1.List(j, 0)
                    pos = InStr(ligne_modif, " ")
                    ligne_modif = Left(ligne_modif, pos - 1)
                    ligne_modif = CInt(ligne_modif)
                
                    'cas code flag siege
                    If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, newrange) = 1 Then
                        Dim choix

                        choix = MsgBox("Le code " & TextBox1.Value & " a été ajouté par le siège à la ligne " & ligne_modif & " et pour région " & Cells(ligne_modif, col_region) & " . Souhaitez vous le remplacer?", 36, "Confirmation")
                        If choix = vbNo Then
                            Exit Sub
                        End If
                    End If
                
                    
                    
                    Dim old_lib As String
                    old_lib = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_lib)
                    
                    
                    
                    If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_codes).Font.ColorIndex <> xlAutomatic Then
                        If LCase(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_codes).Comment.Text) = "code composant" Then

                            Dim var_region As String
                            var_region = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_region)

                            Dim var_typo As String
                            var_typo = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_typo)

                            Dim num_ligne_comp As Long
                            num_ligne_comp = Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=CLng(TextBox1.Value)).Row

                            Dim ligne_tot As Long


                            Dim x_1 As Long

                            Dim num_ligne_comp_1 As Long
                            num_ligne_comp_1 = num_ligne_comp
                            While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_1 - 1, 5).Value = "Composant"
                                ligne_tot = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                                For x_1 = 2 To ligne_tot
                                    If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_1, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_1 - 1, 1).Value And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_1, col_region) = var_region And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_1, col_typo) = var_typo Then

                                        Rows(x_1 & ":" & x_1).EntireRow.delete

                                    End If
                                Next x_1
                                num_ligne_comp_1 = num_ligne_comp_1 - 1
                            Wend



                            Dim x_plus_1 As Long

                            Dim num_ligne_comp_plus_1 As Long
                            num_ligne_comp_plus_1 = num_ligne_comp
                            While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_plus_1 + 1, 5).Value = "Composant"
                                ligne_tot = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                                For x_plus_1 = 2 To ligne_tot
                                    If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_plus_1, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_plus_1 + 1, 1).Value And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_plus_1, col_region) = var_region And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_plus_1, col_typo) = var_typo Then

                                        Rows(x_plus_1 & ":" & x_plus_1).EntireRow.delete

                                    End If
                                Next x_plus_1
                                num_ligne_comp_plus_1 = num_ligne_comp_plus_1 + 1
                            Wend

                            Dim last_ligne_cas_comp As Long
                            last_ligne_cas_comp = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                            Dim ligne_modif_cas_comp As Long
                            Dim x As Long
                            For x = 2 To last_ligne_cas_comp
                                If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x, col_codes) = CLng(TextBox1.Value) And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x, col_region) = var_region And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x, col_typo) = var_typo Then
                                    ligne_modif_cas_comp = x
                                End If
                            Next x
                            
                            
                            
                            'alimentation fichier suivi
                            ligne_suivi = Workbooks(suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 1).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_cas_comp, col_mois).Value
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 2).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_cas_comp, col_region).Value
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 3).Value = TextBox1.Value
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 4).Value = old_lib
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 5).Value = TextBox3.Value
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 6).Value = "Suppression"
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 10).Value = TextBox2.Value
    
                            Workbooks(workbook_compil).Sheets(onglet_compil).Rows(ligne_modif_cas_comp & ":" & ligne_modif_cas_comp).EntireRow.delete
                        
                        'cas fournisseur
                        ElseIf LCase(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_codes).Comment.Text) = "fournisseur" Then
                            
                            Dim var_region_multifourn As String
                            var_region_multifourn = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_region)

                            Dim var_typo_multifourn As String
                            var_typo_multifourn = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_typo)

                            Dim num_ligne_multifourn As Long
                            Dim i_multifourn As Long
                            For i_multifourn = 2 To nbre_ligne_multifourn
                                If Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 1) = CLng(TextBox1.Value) And Replace(UCase(CStr(Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 6).Value)), " ", "") = Replace(var_region_multifourn, " ", "") Then
                                    num_ligne_multifourn = i_multifourn
                                End If
                            Next i_multifourn
                                
                            Dim ligne_tot_multifourn As Long


                            Dim x_multifourn_1 As Long

                            Dim num_ligne_multifourn_1 As Long
                            num_ligne_multifourn_1 = num_ligne_multifourn
                            While LCase(Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_1 - 1, 5).Value) = "fournisseur"
                                ligne_tot_multifourn = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                                For x_multifourn_1 = 2 To ligne_tot_multifourn
                                    If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_multifourn_1, col_codes) = Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_1 - 1, 1).Value And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_multifourn_1, col_region) = var_region_multifourn And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_multifourn_1, col_typo) = var_typo_multifourn Then

                                        Rows(x_multifourn_1 & ":" & x_multifourn_1).EntireRow.delete

                                    End If
                                Next x_multifourn_1
                                num_ligne_multifourn_1 = num_ligne_multifourn_1 - 1
                            Wend



                            Dim x_multifourn_plus_1 As Long

                            Dim num_ligne_multifourn_plus_1 As Long
                            num_ligne_multifourn_plus_1 = num_ligne_multifourn
                            While LCase(Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_plus_1 + 1, 5).Value) = "fournisseur"
                                ligne_tot_multifourn = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                                For x_multifourn_plus_1 = 2 To ligne_tot_multifourn
                                    If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_multifourn_plus_1, col_codes) = Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_plus_1 + 1, 1).Value And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_multifourn_plus_1, col_region) = var_region_multifourn And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_multifourn_plus_1, col_typo) = var_typo_multifourn Then

                                        Rows(x_multifourn_plus_1 & ":" & x_multifourn_plus_1).EntireRow.delete

                                    End If
                                Next x_multifourn_plus_1
                                num_ligne_multifourn_plus_1 = num_ligne_multifourn_plus_1 + 1
                            Wend

                            Dim last_ligne_cas_multifourn As Long
                            last_ligne_cas_multifourn = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                            Dim ligne_modif_cas_multifourn As Long
                            Dim x_multifourn As Long
                            For x_multifourn = 2 To last_ligne_cas_multifourn
                                If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_multifourn, col_codes) = CLng(TextBox1.Value) And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_multifourn, col_region) = var_region_multifourn And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x_multifourn, col_typo) = var_typo_multifourn Then
                                    ligne_modif_cas_multifourn = x_multifourn
                                End If
                            Next x_multifourn
                            
                            
                            
                            'alimentation fichier suivi
                            ligne_suivi = Workbooks(suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 1).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_cas_multifourn, col_mois).Value
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 2).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif_cas_multifourn, col_region).Value
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 3).Value = TextBox1.Value
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 4).Value = old_lib
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 5).Value = TextBox3.Value
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 6).Value = "Suppression"
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 10).Value = TextBox2.Value
    
                            Workbooks(workbook_compil).Sheets(onglet_compil).Rows(ligne_modif_cas_multifourn & ":" & ligne_modif_cas_multifourn).EntireRow.delete
                        
                        Else
                            'alimentation fichier suivi
                            ligne_suivi = Workbooks(suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 1).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_mois).Value
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 2).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_region).Value
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 3).Value = TextBox1.Value
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 4).Value = old_lib
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 5).Value = TextBox3.Value
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 6).Value = "Suppression"
                            Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 10).Value = TextBox2.Value
    
                            Workbooks(workbook_compil).Sheets(onglet_compil).Rows(ligne_modif & ":" & ligne_modif).EntireRow.delete
                        
                        End If
                        
                        
                    Else
                        'alimentation fichier suivi
                        ligne_suivi = Workbooks(suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
                        Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 1).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_mois).Value
                        Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 2).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_region).Value
                        Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 3).Value = TextBox1.Value
                        Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 4).Value = old_lib
                        Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 5).Value = TextBox3.Value
                        Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 6).Value = "Suppression"
                        Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 10).Value = TextBox2.Value

                        Workbooks(workbook_compil).Sheets(onglet_compil).Rows(ligne_modif & ":" & ligne_modif).EntireRow.delete
                        

                    End If
                            
                Else
                    cpt = cpt + 1
                End If

            Next j
            
            If cpt = ListBox1.ListCount Then
                MsgBox "veuillez selectionné un item"
            Else
                Windows(workbook_compil).Activate
                Unload UserForm1
                UserForm1.Show
            
            End If
            

        End If
                        

End Sub




Private Sub supprimer_tout_Click()

'Dim xWb As Workbook
'On Error Resume Next
'Set xWb = Application.Workbooks.Item(fichier_suivi)
'Dim xret As Boolean
'xret = (Not xWb Is Nothing)
Dim suivi As String
'If xret = True Then
'    suivi = fichier_suivi
'Else
'    suivi = new_workbook
'End If
suivi = fichier_suivi



        If WorksheetFunction.CountIf(Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes), Me.TextBox1.Value) = 0 Then
            MsgBox "Ce code ne peut pas être modifié car n'est pas encore renseigné"
        ElseIf TextBox2.Value = "" Then
            MsgBox "Veuillez saisir votre nom dans la case Nom"
        ElseIf TextBox3.Value = "" Then
            MsgBox "Veuillez saisir votre commentaire dans la case Commentaire"

        Else
        
            Dim ligne_suivi As Long
            
            Dim i As Long
            For i = 2 To ligne - 1
                If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_codes) = CLng(TextBox1.Value) And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, newrange) = 1 Then
                    Dim choix
            
                    choix = MsgBox("Le code a été ajouté par le siège à la ligne " & i & " et pour région " & Cells(i, col_region) & " . Souhaitez vous le supprimer?", 36, "Confirmation")
                    If choix = vbNo Then
                        Exit Sub
                    End If
                End If
                
            Next i
            
            
            If WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_lots).Range("A:A"), Me.TextBox1.Value) > 0 Then
                MsgBox "COMP"
                
                Dim num_ligne_comp As Long
                num_ligne_comp = Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=CLng(TextBox1.Value)).Row
    
                Dim ligne_tot As Long
    
                Dim x_1 As Long
    
                Dim num_ligne_comp_1 As Long
                num_ligne_comp_1 = num_ligne_comp
                While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_1 - 1, 5).Value = "Composant"
                    While WorksheetFunction.CountIf(Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes), Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_1 - 1, 1).Value) > 0
                            x_1 = Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes).Find(What:=CLng(Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_1 - 1, 1).Value)).Row
                            Rows(x_1 & ":" & x_1).EntireRow.delete
                    Wend
                    num_ligne_comp_1 = num_ligne_comp_1 - 1
                Wend
    
    
                Dim x_plus_1 As Long
    
                Dim num_ligne_comp_plus_1 As Long
                num_ligne_comp_plus_1 = num_ligne_comp
                While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_plus_1 + 1, 5).Value = "Composant"
                    While WorksheetFunction.CountIf(Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes), Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_plus_1 + 1, 1).Value) > 0
                            x_plus_1 = Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes).Find(What:=CLng(Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_plus_1 + 1, 1).Value)).Row
                            Rows(x_plus_1 & ":" & x_plus_1).EntireRow.delete
                    Wend
                    num_ligne_comp_plus_1 = num_ligne_comp_plus_1 + 1
                Wend
            
                
            ElseIf WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_multifourn).Range("A:A"), Me.TextBox1.Value) > 0 Then
                MsgBox "FOURNISSEUR" & Chr(13) & Chr(10) & "Veuillez utiliser le bouton remplacer ligne pour un code Fournisseur"
                
                
                
            End If
            
           
            While WorksheetFunction.CountIf(Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes), Me.TextBox1.Value) > 0
                Dim num_ligne_compil As Long
                num_ligne_compil = Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes).Find(What:=CLng(TextBox1.Value)).Row

                Dim old_lib As String
                old_lib = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(num_ligne_compil, col_lib)

                'alimentation fichier suivi
                ligne_suivi = Workbooks(suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
                Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 1).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(num_ligne_compil, col_mois).Value
                Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 2).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(num_ligne_compil, col_region).Value
                Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 3).Value = TextBox1.Value
                Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 4).Value = old_lib
                Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 5).Value = TextBox3.Value
                Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 6).Value = "Suppression"
                Workbooks(suivi).Sheets(1).Cells(ligne_suivi, 10).Value = TextBox2.Value

                Workbooks(workbook_compil).Sheets(onglet_compil).Rows(num_ligne_compil & ":" & num_ligne_compil).EntireRow.delete
            Wend
            
            Unload UserForm1
            UserForm1.Show
           
           
            
        End If




End Sub

Private Sub remplacer_min_Click()
    'Dim xWb As Workbook
    'On Error Resume Next
    'Set xWb = Application.Workbooks.Item(fichier_suivi)
    'Dim xret As Boolean
    'xret = (Not xWb Is Nothing)
    Dim suivi As String
    'If xret = True Then
    '    suivi = fichier_suivi
    'Else
    '    suivi = new_workbook
    'End If
    suivi = fichier_suivi
        
        
        If WorksheetFunction.CountIf(Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes), Me.TextBox1.Value) = 0 Then
            MsgBox "Ce code ne peut pas être modifié car n'est pas encore renseigné"
        ElseIf IsNumeric(TextBox4.Value) = False And Len(TextBox4.Value) < 6 Then
            MsgBox "Veuillez rentrez un code convenable"
        ElseIf TextBox2.Value = "" Then
            MsgBox "Veuillez saisir votre nom dans la case Nom"
        ElseIf TextBox3.Value = "" Then
            MsgBox "Veuillez saisir votre commentaire dans la case Commentaire"
        Else

            Dim j As Integer
            Dim cpt As Integer
            cpt = 0


            
        
            For j = 0 To ListBox1.ListCount - 1
                If ListBox1.Selected(j) = True Then
                
                        Dim ligne_modif As Long
                        ligne_modif = ListBox1.List(j, 0)
                        pos = InStr(ligne_modif, " ")
                        ligne_modif = Left(ligne_modif, pos - 1)
                        ligne_modif = CInt(ligne_modif)
                        'Numéro de ligne récupéré

                        Dim new_min As Long
                        new_min = CLong(TextBox8.Value)

                        Dim region_min As String
                        region_min = Workbook(workbook_compil).Sheets(onglet_compil).Cell(ligne_modif, col_region)
                        'TODO:: Trouver la région
                        'A partir d'ici faire le remplacement des mins.
                        Workbook(workbook_compil).Sheets(onglet_compil).Cell(ligne_modif, col_minsouh).Value = new_min

                        'Actualiser les autres MINS
                        Dim nbre_depot_procedure_region As Long
                        nbre_depot_procedure_region = Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_code_region).Range("A:A"), region_min)
                        
                        Dim new_min_tot As Long
                        new_min_tot = new_min * nbre_depot_procedure_region

                        Workbook(workbook_compil).Sheets(onglet_compil).Cell(ligne_modif, col_mintot).Value = new_min_tot
                        TextBox7.Value = new_min
                
                        
                Else
                    cpt = cpt + 1
                End If

            Next j
            
            If cpt = ListBox1.ListCount Then
                MsgBox "veuillez selectionné un item"
            Else
                Windows(workbook_compil).Activate
                Unload UserForm1
                UserForm1.Show
            
            End If
            

        End If

End Sub


Private Sub ListBox1_Click()

    'TODO: Faire le remplissage de la textbox
    Dim j As Integer
    Dim cpt As Integer
    cpt = 0

    For j = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(j) = True Then
            Dim ligne_modif As Long
            ligne_modif = ListBox1.List(j, 0)
            pos = InStr(ligne_modif, " ")
            ligne_modif = Left(ligne_modif, pos - 1)
            ligne_modif = CInt(ligne_modif)
            'Numéro de ligne récupéré

            Dim min_select As Long
            min_select = Workbook(workbook_compil).Sheets(onglet_compil).Cell(ligne_modif, col_minsouh).Value

            TextBox7.Value = min_select

        End If
    Next

End Sub

'
'Private Sub remplacer_tout_Click()
'
'    If WorksheetFunction.CountIf(Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes), Me.TextBox1.Value) = 0 Then
'        MsgBox "Ce code ne peut pas être modifié car n'est pas encore renseigné"
'    ElseIf IsNumeric(TextBox4.Value) = False And Len(TextBox4.Value) < 6 Then
'        MsgBox "Veuillez rentrez un nouveau code convenable"
'    ElseIf TextBox2.Value = "" Then
'        MsgBox "Veuillez saisir votre nom dans la case Nom"
'    ElseIf TextBox3.Value = "" Then
'        MsgBox "Veuillez saisir votre commentaire dans la case Commentaire"
'    Else
'
'        Dim xWb As Workbook
'        On Error Resume Next
'        Set xWb = Application.Workbooks.Item(fichier_suivi)
'        xRet = (Not xWb Is Nothing)
'
'        Dim i As Integer
'        Dim ligne_suivi As Long
'
'
'        If xRet = True Then
'
'            If MsgBox("Confirmez-vous la modification du code " & TextBox1.Value & " par le code " & TextBox4.Value & " sur toute les lignes où il est renseigné?", vbYesNo, "confirmation") = vbYes Then
'
'                For i = 2 To ligne
'                    If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_codes) = CLng(TextBox1.Value) Then
'                        Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_codes) = TextBox4.Value
'
'                        ligne_suivi = Workbooks(fichier_suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
'
'                        Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 1).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_mois).Value
'                        Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 2).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_region).Value
'                        Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 3).Value = TextBox1.Value
'                        Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 4).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_lib).Value
'                        Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 5).Value = TextBox3.Value
'
'                        ligne_suivi = ligne_suivi + 1
'                    End If
'                Next i
'
'            End If
'
'        Else
'
'            strFileName = emplacement_fichier_suivi & fichier_suivi
'            strFileExists = Dir(strFileName)
'                If strFileExists = "" Then
'                    Workbooks.Add
'                    ActiveWorkbook.SaveAs Filename:=emplacement_fichier_suivi & fichier_suivi, _
'                        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
'                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 1).Value = "Mois"
'                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 2).Value = "Région"
'                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 3).Value = "Codes"
'                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 4).Value = "Libellé"
'                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 5).Value = "Retour RM"
'                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 6).Value = "Action paramétrage"
'                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 7).Value = "Systématisation"
'                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 8).Value = "Nouveau code"
'                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 9).Value = "Libellé"
'                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 10).Value = "Nom"
'                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 11).Value = "Commentaire"
'
'                    If MsgBox("Confirmez-vous la modification du code " & TextBox1.Value & " par le code " & TextBox4.Value & " sur toute les lignes où il est renseigné?", vbYesNo, "confirmation") = vbYes Then
'
'                        For i = 2 To ligne
'                            If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_codes) = CLng(TextBox1.Value) Then
'                                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_codes) = TextBox4.Value
'
'                                ligne_suivi = Workbooks(fichier_suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
'
'                                Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 1).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_mois).Value
'                                Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 2).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_region).Value
'                                Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 3).Value = TextBox1.Value
'                                Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 4).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_lib).Value
'                                Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 5).Value = TextBox3.Value
'
'                                ligne_suivi = ligne_suivi + 1
'                            End If
'                        Next i
'
'                    End If
'
'
'
'                Else
'                    Workbooks.Open Filename:=emplacement_fichier_suivi & fichier_suivi
'
'                    If MsgBox("Confirmez-vous la modification du code " & TextBox1.Value & " par le code " & TextBox4.Value & " sur toute les lignes où il est renseigné?", vbYesNo, "confirmation") = vbYes Then
'
'                        For i = 2 To ligne
'                            If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_codes) = CLng(TextBox1.Value) Then
'                                Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_codes) = TextBox4.Value
'
'                                ligne_suivi = Workbooks(fichier_suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
'
'                                Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 1).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_mois).Value
'                                Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 2).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_region).Value
'                                Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 3).Value = TextBox1.Value
'                                Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 4).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i, col_lib).Value
'                                Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 5).Value = TextBox3.Value
'
'                                ligne_suivi = ligne_suivi + 1
'
'                            End If
'                        Next i
'
'                    End If
'                End If
'
'        End If
'
'
''    If WorksheetFunction.CountIf(Workbooks(fichier_suivi).Sheets(1).Range(lettre_col_codes & ":" & lettre_col_codes), Me.TextBox1.Value) = 0 Then
''        MsgBox ("ce code n'existe pas")
''    Else
''            Columns(lettre_col_codes & ":" & lettre_col_codes).Select
''            Selection.Replace What:=TextBox1.Value, Replacement:=TextBox4.Value, LookAt:=xlPart, _
''                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
''                ReplaceFormat:=False
''            Workbooks(fichier_suivi).Sheets(1).Range("A1").Select
''            MsgBox ("remplacement effectué avec succès")
''            Unload UserForm1
''            UserForm1.Show
''    End If
'    End If
'
'End Sub

Public Sub comp(code As Long)

If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_lots).Range("A:A"), code) > 0 Then
'    if Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_compil, col_codes).Comment.Text Text:="Article DEREF" then
End If
End Sub

'Sub new_code_suivi(code As Long)
'If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_lots).Range("A:A"), code) > 0 Then
'    MsgBox "lots/comp"
'End If
'End Sub



'Sub Button1_Click()
'ButtonOneClick = True
'End Sub
'
'Sub Button2_Click()
'If ButtonOneClick Then
'    MsgBox "Button 1 Was Clicked"
'Else
'    MsgBox "Button 1 Was NOT Clicked"
'End If
'
'ButtonOneClick = False
'End Sub



Sub test_clique()

If ButtonOneClick Then
    MsgBox "1"
ElseIf ButtonTwoClick Then
    MsgBox "test"
End If

End Sub

Sub test_click()
ButtonTwoClick = True
Call test_clique
ButtonTwoClick = False
End Sub



Private Sub CommandButton1_Click()
ButtonOneClick = True
Call test_clique
ButtonOneClick = False
End Sub













'Private Sub remplacer_ligne_Click()
'        Dim j As Integer
'        Dim cpt As Integer
'        Dim ligne_suivi As Long
'        cpt = 0
'        If WorksheetFunction.CountIf(Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes), Me.TextBox1.Value) = 0 Then
'            MsgBox "Ce code ne peut pas être modifié car n'est pas encore renseigné"
'        ElseIf IsNumeric(TextBox4.Value) = False And Len(TextBox4.Value) < 6 Then
'            MsgBox "Veuillez rentrez un code convenable"
'        ElseIf TextBox2.Value = "" Then
'            MsgBox "Veuillez saisir votre nom dans la case Nom"
'        ElseIf TextBox3.Value = "" Then
'            MsgBox "Veuillez saisir votre commentaire dans la case Commentaire"
'        Else
'            For j = 0 To ListBox1.ListCount - 1
'                If ListBox1.Selected(j) = True Then
'
'                    If IsNumeric(TextBox4.Value) = True And Len(TextBox4.Value) = 6 And TextBox2.Value <> "" And TextBox3.Value <> "" Then
'                        Dim ligne_modif As Integer
'                        ligne_modif = ListBox1.List(j, 0)
'                        If MsgBox("Confirmez-vous la modification du code " & TextBox1.Value & " par le code " & TextBox4.Value & " à la ligne " & ligne_modif & "?", vbYesNo, "confirmation") = vbYes Then
'                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_codes).Value = TextBox4.Value
'                            MsgBox "Modification effectué sur le Code " & TextBox1.Value & " remplacé par le code " & TextBox4.Value & " à la ligne " & ligne_modif
'
'                            Dim xWb As Workbook
'                            On Error Resume Next
'                            Set xWb = Application.Workbooks.Item(fichier_suivi)
'                            xRet = (Not xWb Is Nothing)
'
'                            If xRet = True Then
'                                ligne_suivi = Workbooks(fichier_suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
'
'                                Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 1).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_mois).Value
'                                Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 2).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_region).Value
'                                Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 3).Value = TextBox1.Value
'                                Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 4).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_lib).Value
'                                Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 5).Value = TextBox3.Value
'                            Else
'                                strFileName = emplacement_fichier_suivi & fichier_suivi
'                                strFileExists = Dir(strFileName)
'                                If strFileExists = "" Then
'                                    Workbooks.Add
'                                    ActiveWorkbook.SaveAs Filename:=emplacement_fichier_suivi & fichier_suivi, _
'                                        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 1).Value = "Mois"
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 2).Value = "Région"
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 3).Value = "Codes"
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 4).Value = "Libellé"
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 5).Value = "Retour RM"
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 6).Value = "Action paramétrage"
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 7).Value = "Systématisation"
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 8).Value = "Nouveau code"
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 9).Value = "Libellé"
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 10).Value = "Nom"
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(1, 11).Value = "Commentaire"
'
'                                    ligne_suivi = Workbooks(fichier_suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
'
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 1).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_mois).Value
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 2).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_region).Value
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 3).Value = TextBox1.Value
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 4).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_lib).Value
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 5).Value = TextBox3.Value
'                                Else
'                                    Workbooks.Open Filename:=emplacement_fichier_suivi & fichier_suivi
'                                    ligne_suivi = Workbooks(fichier_suivi).Sheets(1).Cells(Application.Rows.Count, 3).End(xlUp).Row + 1
'
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 1).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_mois).Value
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 2).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_region).Value
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 3).Value = TextBox1.Value
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 4).Value = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_lib).Value
'                                    Workbooks(fichier_suivi).Sheets(1).Cells(ligne_suivi, 5).Value = TextBox3.Value
'                                End If
'                            End If
'
'
'                        End If
'
'                    End If
'                Else
'                    cpt = cpt + 1
'                End If
'
'            Next j
'            If cpt = ListBox1.ListCount Then
'                MsgBox "veuillez selectionné un item"
'            End If
'
'            Windows(workbook_compil).Activate
'            Unload UserForm1
'            UserForm1.Show
'        End If
'
'
'End Sub




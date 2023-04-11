Option Explicit




Public ligne As Integer

Public col_codes As Integer
Public lettre_col_codes As String

Public col_mois As Integer

Public col_region As Integer

Public col_min As Integer

Public col_typo As Integer

Public col_lib As Integer

Public col_ean As Integer

Public col_rayon As Integer

Public col_occ As Integer
Public lettre_col_occ As String

Public col_pcb As Integer

Public col_ca As Integer

Public col_marge As Integer

Public col_rattachement As Integer

Public col_couv As Integer

Public col_flag_siege As Integer


Public dif As Long

Public onglet_fichier_base As String
Public onglet_info As String
Public workbook_fichier_base As String


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

Public code_region As String

Public region As String

Public nbre_ligne_multifourn As Long

Public ligne_region As Long
Public ligne_rattachement As Long
Public nbre_col_ratt As Long

Public ligne_suivi As Long
Public flag_remplacement As Integer
Public ligne_remplacement As Long
Public col_codes_ratt As Long











Private Sub UserForm_initialize()




    workbook_fichier_base = ActiveWorkbook.Name

    'Dim j As Integer
    'Dim num_onglet As Integer
    'For j = 1 To Sheets.Count
    '    If Sheets(j).Name = CStr(Replace(Replace(ActiveWorkbook.Name, ".xlsm", ""), "Userform_", "")) Then
    '        num_onglet = j
    '    End If
    'Next j
    '
    'Sheets(num_onglet).Name = CStr(Replace(Replace(ActiveWorkbook.Name, ".xlsm", ""), "Userform_", ""))
    'onglet_fichier_base = Sheets(num_onglet).Name
    Dim prefix As String
    prefix = Left(workbook_fichier_base, 3)
    Sheets(2).Name = CStr(Replace(Replace(ActiveWorkbook.Name, ".xlsm", ""), prefix, ""))
    onglet_fichier_base = Sheets(2).Name


    Windows(workbook_fichier_base).Activate
    Sheets(onglet_fichier_base).Select

    onglet_info = "Info"


    'region = UCase(CStr(Replace(Replace(ActiveWorkbook.Name, ".xlsm", ""), "Userform_", "")))
    region = UCase(onglet_fichier_base)



    base_data = "Base data.xlsx"
    'code_region = "code_region.xlsx"
    onglet_transco = "transco"
    onglet_lots = "lots"
    onglet_derefs = "derefs"
    onglet_souscat = "sous cat"
    onglet_pcb = "PCB"
    onglet_rattachement = "rattachement"
    onglet_marge = "PV et marge"
    onglet_code_region = "code region"
    onglet_suivi = "Suivi Delphine"
    onglet_BI = "BI "
    onglet_multifourn = "Multifournisseurs"


    Dim newrange As Integer
    Dim i As Integer
    newrange = Range("A1").End(xlToRight).Column
    For i = 1 To newrange
        If LCase(Cells(1, i).Value) = "codes" Then
            col_codes = i
        ElseIf LCase(Cells(1, i).Value) = "mois" Then
            col_mois = i
        ElseIf Replace(LCase(Cells(1, i).Value), "é", "e") = "region" Then
            col_region = i
        ElseIf Replace(Replace(LCase(Cells(1, i).Value), "é", "e"), " ", "") = "minsouhaite" Or Replace(LCase(Cells(1, i).Value), " ", "") = "min" Then
            col_min = i
        ElseIf LCase(Cells(1, i).Value) = "typo" Or LCase(Cells(1, i).Value) = "typologie" Then
            col_typo = i
        ElseIf Replace(LCase(Cells(1, i).Value), "é", "e") = "libelle" Then
            col_lib = i
        ElseIf LCase(Cells(1, i).Value) = "ean" Then
            col_ean = i
        ElseIf LCase(Cells(1, i).Value) = "rayon" Then
            col_rayon = i
        ElseIf LCase(Cells(1, i).Value) = "occurrence" Then
            col_occ = i
        ElseIf LCase(Cells(1, i).Value) = "pcb" Then
            col_pcb = i
        ElseIf LCase(Cells(1, i).Value) = "ca engage" Then
            col_ca = i
        ElseIf LCase(Cells(1, i).Value) = "marge" Or Replace(LCase(Cells(1, i).Value), " ", "") = "tauxdemarge" Then
            col_marge = i
        ElseIf Replace(LCase(Cells(1, i).Value), " ", "") = "tauxderattachement" Then
            col_rattachement = i
        ElseIf LCase(Cells(1, i).Value) = "couverture" Then
            col_couv = i
        ElseIf LCase(Cells(1, i).Value) = "flag" Or Replace(LCase(Cells(1, i).Value), " ", "") = "flagcodesiege" Then
            col_flag_siege = i
        
        End If
    Next i

    lettre_col_codes = Split(Cells(1, col_codes).Address, "$")(1)
    lettre_col_occ = Split(Cells(1, col_occ).Address, "$")(1)



    'multifourn
    nbre_ligne_multifourn = Workbooks(base_data).Sheets(onglet_multifourn).Cells(Application.Rows.Count, 1).End(xlUp).Row


    'rattachement
    ligne_region = Workbooks(base_data).Sheets(onglet_code_region).Cells(Application.Rows.Count, 1).End(xlUp).Row
    ligne_rattachement = Workbooks(base_data).Sheets(onglet_rattachement).Cells(Application.Rows.Count, 1).End(xlUp).Row
    nbre_col_ratt = Workbooks(base_data).Sheets(onglet_rattachement).Range("A7").End(xlToRight).Column
    Dim k As Long
    For k = 1 To nbre_col_ratt
        If LCase(Workbooks(base_data).Sheets(onglet_rattachement).Cells(1, k).Value) = "article" Then
            col_codes_ratt = k
        End If
    Next k

    'suivi
    ligne_suivi = Workbooks(base_data).Sheets(onglet_suivi).Cells(Application.Rows.Count, 1).End(xlUp).Row


    'ligne = Sheets("Feuil1").Range("F45678").End(xlUp).Row + 1
    ligne = Sheets(onglet_fichier_base).Cells(Application.Rows.Count, col_codes).End(xlUp).Row + 1
    'If Cells(ligne, 1) = "" Then
    '    MsgBox ("tableau entièrement complété")
    '    Unload Me
    'Else

    'Label1.Font.Size = 15
    'Label2.Font.Size = 15
    'Label3.Font.Size = 15
    'Label4.Font.Size = 15
    '
    '
    'Occurrence.Font.Size = 12
    'Label6.Font.Size = 10
    '
    TextBox1.Font.Size = 12
    TextBox14.Font.Size = 12
    TextBox7.Font.Size = 12
    TextBox13.Font.Size = 12
    TextBox15.Font.Size = 12
    TextBox17.Font.Size = 10
    TextBox20.Font.Size = 12
    TextBox16.Font.Size = 9
    TextBox19.Font.Size = 12
    TextBox8.Font.Size = 12
    ListBox4.Font.Size = 12


    TextBox9.Value = ligne
    TextBox9.Font.Size = 12




    If ligne - 2 > 60 Then
        Me.TextBox16.BackColor = vbRed
        MsgBox "Limite de données renseignées dépassée"
        TextBox16.Value = ligne - 2 & " / 60" & vbCrLf & "Limite dépassée"
    Else
        TextBox16.Value = ligne - 2 & " / 60"
    End If
        


    dif = col_codes - col_occ


    CheckBox1.Value = True
    CheckBox1.Enabled = False




    'End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Sheets(onglet_fichier_base).Protect Password:=motperdu, DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
    
    Windows(workbook_fichier_base).Activate
    Sheets("Graphique").PivotTables("repartition_rayon").PivotCache.Refresh
    
    Sheets("Graphique").PivotTables("ca_engage").PivotCache.Refresh

End Sub



Private Sub Fermer_Click()

    Sheets(onglet_fichier_base).Protect Password:=motperdu, DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
    
    '    Windows(base_data).Activate
    '    ActiveWindow.Close SaveChanges:=False
        
    Windows(workbook_fichier_base).Activate
    Sheets("Graphique").PivotTables("repartition_rayon").PivotCache.Refresh
    
    Sheets("Graphique").PivotTables("ca_engage").PivotCache.Refresh
        
    Unload Me
End Sub

Sub modif_ca_couv(ligne_ajout)

        'variable ca engage + couverture
        Dim code_ean As Double
        Dim prmp As Double
        Dim nbre_depot As Double
        Dim min As Double
        Dim somme As Double
        Dim tx_marge As Double
        Dim prix_vente As Double
                       'couverture
                        If Application.WorksheetFunction.IfError(Application.VLookup(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0), 0) Then
                            If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_BI).Range("C:C"), Application.WorksheetFunction.VLookup(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)) > 0 Then
                                'Dim code_ean As Double
                                code_ean = Application.WorksheetFunction.VLookup(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)
                                'Dim prmp As Double
                                prmp = Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("C:F"), 4, 0)
                                'Dim nbre_depot As Double
                                nbre_depot = Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("C:G"), 5, 0)
                                'Dim min As Double
                                min = 600 * nbre_depot
                                'Dim somme As Double
                                somme = Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("C:T"), 18, 0)
                                If prmp <> 0 Then
                                    If somme <> 0 Or nbre_depot <> 0 Then
                                        Cells(ligne_ajout, col_couv) = Round((TextBox14.Value * prmp) / ((somme / nbre_depot) * prmp) * 12, 2)
                                    Else
                                        Cells(ligne_ajout, col_couv) = "-"
                                    End If
                                Else
                                    Cells(ligne_ajout, col_couv) = "-"
                                End If
                            Else
                                Cells(ligne_ajout, col_couv) = "-"
                            End If
                        Else
                            Cells(ligne_ajout, col_couv) = "-"
                        End If

                
                        
                        'CA engagé + Marge
                        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_marge).Range("C:C"), Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes)) > 0 Then
                
                            tx_marge = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes)), Workbooks(base_data).Sheets(onglet_marge).Range("C:E"), 3, 0)
                            If tx_marge <> 0 Then
                                Cells(ligne_ajout, col_marge) = Round(tx_marge, 2)
                            Else
                                Cells(ligne_ajout, col_marge) = "-"
                            End If
                
                            prix_vente = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes)), Workbooks(base_data).Sheets(onglet_marge).Range("C:D"), 2, 0)
                            If prix_vente <> 0 Then
                                Cells(ligne_ajout, col_ca) = prix_vente * TextBox14.Value
                            Else
                                Cells(ligne_ajout, col_ca) = "-"
                            End If
                
                        End If


End Sub


Private Sub Modifier_Click()
    Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Unprotect Password:=motperdu

        Dim j As Integer
        Dim cpt As Integer
        cpt = 0
        If WorksheetFunction.CountIf(Sheets(onglet_fichier_base).Range(lettre_col_codes & ":" & lettre_col_codes), Me.TextBox1.Value) = 0 Then
            MsgBox "Ce code ne peut pas être modifié car n'est pas encore renseigné"
        Else
            
            For j = 0 To ListBox4.ListCount - 1
                If ListBox4.Selected(j) = True Then
                    Dim ligne_modif As Integer
                    ligne_modif = ListBox4.List(j, 0)
                    If Cells(ligne_modif, col_flag_siege) = 1 Then
                        MsgBox ("Impossible de modifier" & Chr(13) & Chr(10) & "Ce code correspond à un code rentré par le SIEGE")
                    Else
                        If IsNumeric(TextBox14.Value) = True And TextBox14.Value <> "" Then
                            If MsgBox("Confirmez-vous la modification du min sur le code " & TextBox1.Value & " à la ligne " & ligne_modif & "?", vbYesNo, "confirmation") = vbYes Then
                                
                                If Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_modif, col_codes).Font.ColorIndex <> xlAutomatic Then
                                    
                                    'cas multifourn
                                    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_multifourn).Range("A:A"), CLng(TextBox1.Value)) > 0 Then
                                                                   
                                        Dim var_typo_multifourn As String
                                        var_typo_multifourn = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_modif, col_typo)
                                        
                                        Dim i_multifourn As Long
                                        For i_multifourn = 2 To nbre_ligne_multifourn
                                            If Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 1).Value = CLng(TextBox1.Value) And Replace(UCase(CStr(Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 6).Value)), " ", "") = Replace(region, " ", "") Then
                                                Dim ligne_multifourn
                                                ligne_multifourn = i_multifourn
                                            End If
                                        Next i_multifourn
                                    
                                    
                                        Dim x_1_multifourn As Long
            
                                        Dim num_ligne_multifourn_1 As Long
                                        num_ligne_multifourn_1 = ligne_multifourn
                                        While Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_1, 5).Value = "fournisseur"
                                            
                                            For x_1_multifourn = 2 To ligne
                                                If Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_1_multifourn, col_codes) = Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_1, 1).Value And Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_1_multifourn, col_typo) = var_typo_multifourn Then
                                                    
                                                    Cells(x_1_multifourn, col_min).Value = TextBox14.Value
                                                    Cells(x_1_multifourn, col_min + 1).Value = Round(TextBox14.Value / 6, 0)
                                                    Cells(x_1_multifourn, col_min + 2).Value = Cells(x_1_multifourn, col_min + 1) * 2
                                                    Cells(x_1_multifourn, col_min + 3).Value = Cells(x_1_multifourn, col_min + 1) * 3
                                                    Cells(x_1_multifourn, col_min + 4).Value = Cells(x_1_multifourn, col_min + 1) * 4
                                                    Cells(x_1_multifourn, col_min + 5).Value = Cells(x_1_multifourn, col_min + 1) * 5
                                                    Cells(x_1_multifourn, col_min + 6).Value = Cells(x_1_multifourn, col_min + 1) * 6
                                                    Cells(x_1_multifourn, col_min + 7).Value = Cells(x_1_multifourn, col_min + 1) * 7
                                                    Cells(x_1_multifourn, col_min + 8).Value = Cells(x_1_multifourn, col_min + 1) * 8
                                                    Cells(x_1_multifourn, col_min + 9).Value = Cells(x_1_multifourn, col_min + 1) * 9
                                                    Cells(x_1_multifourn, col_min + 10).Value = Cells(x_1_multifourn, col_min + 1) * 10
                                                    Cells(x_1_multifourn, col_min + 11).Value = Cells(x_1_multifourn, col_min + 1) * 11
                                                    
                                                    Call modif_ca_couv(x_1_multifourn)
                                                    
            
                                                End If
                                            Next x_1_multifourn
                                            num_ligne_multifourn_1 = num_ligne_multifourn_1 - 1
                                        Wend
            
            
            
                                        Dim x_plus_1_multifourn As Long
            
                                        Dim num_ligne_multifourn_plus_1 As Long
                                        num_ligne_multifourn_plus_1 = ligne_multifourn
                                        While Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_plus_1 + 1, 5).Value = "fournisseur"
                                            
                                            For x_plus_1_multifourn = 2 To ligne
                                                If Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_plus_1_multifourn, col_codes) = Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_plus_1 + 1, 1).Value And Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_plus_1_multifourn, col_typo) = var_typo_multifourn Then
                                                    
                                                    Cells(x_plus_1_multifourn, col_min).Value = TextBox14.Value
                                                    Cells(x_plus_1_multifourn, col_min + 1).Value = Round(TextBox14.Value / 6, 0)
                                                    Cells(x_plus_1_multifourn, col_min + 2).Value = Cells(x_plus_1_multifourn, col_min + 1) * 2
                                                    Cells(x_plus_1_multifourn, col_min + 3).Value = Cells(x_plus_1_multifourn, col_min + 1) * 3
                                                    Cells(x_plus_1_multifourn, col_min + 4).Value = Cells(x_plus_1_multifourn, col_min + 1) * 4
                                                    Cells(x_plus_1_multifourn, col_min + 5).Value = Cells(x_plus_1_multifourn, col_min + 1) * 5
                                                    Cells(x_plus_1_multifourn, col_min + 6).Value = Cells(x_plus_1_multifourn, col_min + 1) * 6
                                                    Cells(x_plus_1_multifourn, col_min + 7).Value = Cells(x_plus_1_multifourn, col_min + 1) * 7
                                                    Cells(x_plus_1_multifourn, col_min + 8).Value = Cells(x_plus_1_multifourn, col_min + 1) * 8
                                                    Cells(x_plus_1_multifourn, col_min + 9).Value = Cells(x_plus_1_multifourn, col_min + 1) * 9
                                                    Cells(x_plus_1_multifourn, col_min + 10).Value = Cells(x_plus_1_multifourn, col_min + 1) * 10
                                                    Cells(x_plus_1_multifourn, col_min + 11).Value = Cells(x_plus_1_multifourn, col_min + 1) * 11
                                                    
                                                    Call modif_ca_couv(x_plus_1_multifourn)
            
                                                End If
                                            Next x_plus_1_multifourn
                                            num_ligne_multifourn_plus_1 = num_ligne_multifourn_plus_1 + 1
                                        Wend
            
                                        
                                        MsgBox "Modification effectué sur le Code " & TextBox1.Value & " à la ligne " & ligne_modif
                                        Unload UserForm1
                                        UserForm1.Show
                     
                                    
                                    ElseIf Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_lots).Range("A:A"), CLng(TextBox1.Value)) > 0 Then

                                           
                                        Dim var_typo As String
                                        var_typo = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_modif, col_typo)
            
                                        Dim num_ligne_comp As Long
                                        num_ligne_comp = Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=CLng(TextBox1.Value)).Row
            
                                        
            
            
                                        Dim x_1 As Long
            
                                        Dim num_ligne_comp_1 As Long
                                        num_ligne_comp_1 = num_ligne_comp
                                        While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_1, 5).Value = "Composant"
                                            
                                            For x_1 = 2 To ligne
                                                If Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_1, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_1, 1).Value And Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_1, col_typo) = var_typo Then
                                                    
                                                    Cells(x_1, col_min).Value = TextBox14.Value
                                                    Cells(x_1, col_min + 1).Value = Round(TextBox14.Value / 6, 0)
                                                    Cells(x_1, col_min + 2).Value = Cells(x_1, col_min + 1) * 2
                                                    Cells(x_1, col_min + 3).Value = Cells(x_1, col_min + 1) * 3
                                                    Cells(x_1, col_min + 4).Value = Cells(x_1, col_min + 1) * 4
                                                    Cells(x_1, col_min + 5).Value = Cells(x_1, col_min + 1) * 5
                                                    Cells(x_1, col_min + 6).Value = Cells(x_1, col_min + 1) * 6
                                                    Cells(x_1, col_min + 7).Value = Cells(x_1, col_min + 1) * 7
                                                    Cells(x_1, col_min + 8).Value = Cells(x_1, col_min + 1) * 8
                                                    Cells(x_1, col_min + 9).Value = Cells(x_1, col_min + 1) * 9
                                                    Cells(x_1, col_min + 10).Value = Cells(x_1, col_min + 1) * 10
                                                    Cells(x_1, col_min + 11).Value = Cells(x_1, col_min + 1) * 11
                                                    
                                                    Call modif_ca_couv(x_1)
                                                    
                                                End If
                                            Next x_1
                                            num_ligne_comp_1 = num_ligne_comp_1 - 1
                                        Wend
            
            
            
                                        Dim x_plus_1 As Long
            
                                        Dim num_ligne_comp_plus_1 As Long
                                        num_ligne_comp_plus_1 = num_ligne_comp
                                        While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_plus_1 + 1, 5).Value = "Composant"
                                            
                                            For x_plus_1 = 2 To ligne
                                                If Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_plus_1, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_plus_1 + 1, 1).Value And Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_plus_1, col_typo) = var_typo Then
                                                    
                                                    Cells(x_plus_1, col_min).Value = TextBox14.Value
                                                    Cells(x_plus_1, col_min + 1).Value = Round(TextBox14.Value / 6, 0)
                                                    Cells(x_plus_1, col_min + 2).Value = Cells(x_plus_1, col_min + 1) * 2
                                                    Cells(x_plus_1, col_min + 3).Value = Cells(x_plus_1, col_min + 1) * 3
                                                    Cells(x_plus_1, col_min + 4).Value = Cells(x_plus_1, col_min + 1) * 4
                                                    Cells(x_plus_1, col_min + 5).Value = Cells(x_plus_1, col_min + 1) * 5
                                                    Cells(x_plus_1, col_min + 6).Value = Cells(x_plus_1, col_min + 1) * 6
                                                    Cells(x_plus_1, col_min + 7).Value = Cells(x_plus_1, col_min + 1) * 7
                                                    Cells(x_plus_1, col_min + 8).Value = Cells(x_plus_1, col_min + 1) * 8
                                                    Cells(x_plus_1, col_min + 9).Value = Cells(x_plus_1, col_min + 1) * 9
                                                    Cells(x_plus_1, col_min + 10).Value = Cells(x_plus_1, col_min + 1) * 10
                                                    Cells(x_plus_1, col_min + 11).Value = Cells(x_plus_1, col_min + 1) * 11
                                                    
                                                    Call modif_ca_couv(x_plus_1)
            
                                                End If
                                            Next x_plus_1
                                            num_ligne_comp_plus_1 = num_ligne_comp_plus_1 + 1
                                        Wend
            
                                        
                                        MsgBox "Modification effectué sur le Code " & TextBox1.Value & " à la ligne " & ligne_modif
                                        Unload UserForm1
                                        UserForm1.Show
                                    Else
                                        Cells(ligne_modif, col_min).Value = TextBox14.Value
                                        Cells(ligne_modif, col_min + 1).Value = Round(TextBox14.Value / 6, 0)
                                        Cells(ligne_modif, col_min + 2).Value = Cells(ligne_modif, col_min + 1) * 2
                                        Cells(ligne_modif, col_min + 3).Value = Cells(ligne_modif, col_min + 1) * 3
                                        Cells(ligne_modif, col_min + 4).Value = Cells(ligne_modif, col_min + 1) * 4
                                        Cells(ligne_modif, col_min + 5).Value = Cells(ligne_modif, col_min + 1) * 5
                                        Cells(ligne_modif, col_min + 6).Value = Cells(ligne_modif, col_min + 1) * 6
                                        Cells(ligne_modif, col_min + 7).Value = Cells(ligne_modif, col_min + 1) * 7
                                        Cells(ligne_modif, col_min + 8).Value = Cells(ligne_modif, col_min + 1) * 8
                                        Cells(ligne_modif, col_min + 9).Value = Cells(ligne_modif, col_min + 1) * 9
                                        Cells(ligne_modif, col_min + 10).Value = Cells(ligne_modif, col_min + 1) * 10
                                        Cells(ligne_modif, col_min + 11).Value = Cells(ligne_modif, col_min + 1) * 11
                                        
                                        Call modif_ca_couv(ligne_modif)
                                        
                                        MsgBox "Modification effectué sur le Code " & TextBox1.Value & " à la ligne " & ligne_modif
                                        Unload UserForm1
                                        UserForm1.Show
                                    End If
                                Else
                                    Cells(ligne_modif, col_min).Value = TextBox14.Value
                                    Cells(ligne_modif, col_min + 1).Value = Round(TextBox14.Value / 6, 0)
                                    Cells(ligne_modif, col_min + 2).Value = Cells(ligne_modif, col_min + 1) * 2
                                    Cells(ligne_modif, col_min + 3).Value = Cells(ligne_modif, col_min + 1) * 3
                                    Cells(ligne_modif, col_min + 4).Value = Cells(ligne_modif, col_min + 1) * 4
                                    Cells(ligne_modif, col_min + 5).Value = Cells(ligne_modif, col_min + 1) * 5
                                    Cells(ligne_modif, col_min + 6).Value = Cells(ligne_modif, col_min + 1) * 6
                                    Cells(ligne_modif, col_min + 7).Value = Cells(ligne_modif, col_min + 1) * 7
                                    Cells(ligne_modif, col_min + 8).Value = Cells(ligne_modif, col_min + 1) * 8
                                    Cells(ligne_modif, col_min + 9).Value = Cells(ligne_modif, col_min + 1) * 9
                                    Cells(ligne_modif, col_min + 10).Value = Cells(ligne_modif, col_min + 1) * 10
                                    Cells(ligne_modif, col_min + 11).Value = Cells(ligne_modif, col_min + 1) * 11
                                    
                                    Call modif_ca_couv(ligne_modif)
                                    
                                    MsgBox "Modification effectué sur le Code " & TextBox1.Value & " à la ligne " & ligne_modif
                                    Unload UserForm1
                                    UserForm1.Show
                                End If
                            End If
                        Else
                            MsgBox "Veuillez rentrez un min"
                        End If
                    End If
                Else
                    cpt = cpt + 1
                End If
            Next j
            If cpt = ListBox4.ListCount Then
                MsgBox "veuillez selectionné un item"
            End If
        End If
        
End Sub

Private Sub delete_Click()

    Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Unprotect Password:=motperdu

        Dim j As Integer
        Dim cpt As Integer
        cpt = 0
        If WorksheetFunction.CountIf(Sheets(onglet_fichier_base).Range(lettre_col_codes & ":" & lettre_col_codes), Me.TextBox1.Value) = 0 Then
            MsgBox "Ce code ne peut pas être modifié car n'est pas encore renseigné"
        Else
            
        
            For j = 0 To ListBox4.ListCount - 1
                If ListBox4.Selected(j) = True Then
                    Dim ligne_modif As Integer
                    ligne_modif = ListBox4.List(j, 0)
                    If Cells(ligne_modif, col_flag_siege) = 1 Then
                        MsgBox ("Impossible de supprimer" & Chr(13) & Chr(10) & "Ce code correspond à un code rentré par le SIEGE")
                    Else
                        If MsgBox("Confirmez-vous la suppression du code " & TextBox1.Value & " à la ligne " & ligne_modif & "?", vbYesNo, "confirmation") = vbYes Then
                            
                            
                            If Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_modif, col_codes).Font.ColorIndex <> xlAutomatic Then
                                
                                If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_multifourn).Range("A:A"), CLng(TextBox1.Value)) > 0 Then
                                                                   
                                    Dim var_typo_multifourn As String
                                    var_typo_multifourn = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_modif, col_typo)
                                    
                                    Dim i_multifourn As Long
                                    For i_multifourn = 2 To nbre_ligne_multifourn
                                        If Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 1).Value = CLng(TextBox1.Value) And Replace(UCase(CStr(Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 6).Value)), " ", "") = Replace(region, " ", "") Then
                                            Dim ligne_multifourn
                                            ligne_multifourn = i_multifourn
                                        End If
                                    Next i_multifourn
                                    
                                    
                                    
                                    Dim x_1_multifourn As Long
        
                                    Dim num_ligne_multifourn_1 As Long
                                    num_ligne_multifourn_1 = ligne_multifourn
                                    While Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_1 - 1, 5).Value = "fournisseur"
                                        
                                        For x_1_multifourn = 2 To 200
                                            If Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_1_multifourn, col_codes) = Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_1 - 1, 1).Value And Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_1_multifourn, col_typo) = var_typo_multifourn Then
        
    '                                                Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Rows(x_1_multifourn).EntireRow.delete
                                                Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Range("A" & x_1_multifourn).EntireRow.delete
                                            End If
                                        Next x_1_multifourn
                                        num_ligne_multifourn_1 = num_ligne_multifourn_1 - 1
                                    Wend
        
        
        
                                    Dim x_plus_1_multifourn As Long

                                    Dim num_ligne_multifourn_plus_1 As Long
                                    num_ligne_multifourn_plus_1 = ligne_multifourn
                                    While Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_plus_1 + 1, 5).Value = "fournisseur"
                                        
                                        For x_plus_1_multifourn = 2 To 200
                                            If Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_plus_1_multifourn, col_codes) = Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_multifourn_plus_1 + 1, 1).Value And Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_plus_1_multifourn, col_typo) = var_typo_multifourn Then

    '                                                Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Rows(x_plus_1_multifourn).EntireRow.delete
                                                Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Range("A" & x_plus_1_multifourn).EntireRow.delete
                                            End If
                                        Next x_plus_1_multifourn
                                        num_ligne_multifourn_plus_1 = num_ligne_multifourn_plus_1 + 1
                                    Wend

                                    Dim last_ligne_cas_multifourn As Long
                                    last_ligne_cas_multifourn = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                                    Dim ligne_modif_cas_multifourn As Long
                                    Dim x_multifourn As Long
                                    For x_multifourn = 2 To last_ligne_cas_multifourn
                                        If Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_multifourn, col_codes) = CLng(TextBox1.Value) And Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_multifourn, col_typo) = var_typo_multifourn Then
                                            ligne_modif_cas_multifourn = x_multifourn
                                        End If
                                    Next x_multifourn


                                    Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Range("A" & ligne_modif_cas_multifourn).EntireRow.delete
                                    MsgBox "Suppression effectué sur le Code " & TextBox1.Value & " à la ligne " & ligne_modif_cas_multifourn
                                    Unload UserForm1
                                    UserForm1.Show
                                
                                
                                
                                
                                ElseIf Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_lots).Range("A:A"), CLng(TextBox1.Value)) > 0 Then

                                           
                                    Dim var_typo As String
                                    var_typo = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_modif, col_typo)
        
                                    Dim num_ligne_comp As Long
                                    num_ligne_comp = Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=CLng(TextBox1.Value)).Row
        
    '                                    Dim ligne_tot As Long
        
        
                                    Dim x_1 As Long
        
                                    Dim num_ligne_comp_1 As Long
                                    num_ligne_comp_1 = num_ligne_comp
                                    While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_1 - 1, 5).Value = "Composant"
    '                                        ligne_tot = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                                        For x_1 = 2 To 200
                                            If Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_1, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_1 - 1, 1).Value And Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_1, col_typo) = var_typo Then
        
    '                                                Rows(x_1).EntireRow.delete
                                                 Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Range("A" & x_1).EntireRow.delete
                                            End If
                                        Next x_1
                                        num_ligne_comp_1 = num_ligne_comp_1 - 1
                                    Wend
        
        
        
                                    Dim x_plus_1 As Long
        
                                    Dim num_ligne_comp_plus_1 As Long
                                    num_ligne_comp_plus_1 = num_ligne_comp
                                    While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_plus_1 + 1, 5).Value = "Composant"
    '                                        ligne_tot = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                                        For x_plus_1 = 2 To 200
                                            If Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_plus_1, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_comp_plus_1 + 1, 1).Value And Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x_plus_1, col_typo) = var_typo Then
        
    '                                                Rows(x_plus_1).EntireRow.delete
                                                Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Range("A" & x_plus_1).EntireRow.delete
                                            End If
                                        Next x_plus_1
                                        num_ligne_comp_plus_1 = num_ligne_comp_plus_1 + 1
                                    Wend
        
                                    Dim last_ligne_cas_comp As Long
                                    last_ligne_cas_comp = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(Application.Rows.Count, col_codes).End(xlUp).Row
                                    Dim ligne_modif_cas_comp As Long
                                    Dim x As Long
                                    For x = 2 To last_ligne_cas_comp
                                        If Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x, col_codes) = CLng(TextBox1.Value) And Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x, col_typo) = var_typo Then
                                            ligne_modif_cas_comp = x
                                        End If
                                    Next x
                                    
            
                                    Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Range("A" & ligne_modif_cas_comp).EntireRow.delete
                                    MsgBox "Suppression effectué sur le Code " & TextBox1.Value & " à la ligne " & ligne_modif_cas_comp
                                    Unload UserForm1
                                    UserForm1.Show
                                    
                                Else
                                    Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Range("A" & ligne_modif).EntireRow.delete
                                    MsgBox "Suppression effectué sur le Code " & TextBox1.Value & " à la ligne " & ligne_modif
                                    Unload UserForm1
                                    UserForm1.Show
                                End If
                            Else
    '                                Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Rows(ligne_modif).delete
                                Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Range("A" & ligne_modif).EntireRow.delete
                                MsgBox "Suppression effectué sur le Code " & TextBox1.Value & " à la ligne " & ligne_modif
                                Unload UserForm1
                                UserForm1.Show
                            End If
                                    

                        End If
                    End If
                Else
                    cpt = cpt + 1
                End If
            Next j
            If cpt = ListBox4.ListCount Then
                MsgBox "veuillez selectionné un item"
            End If
        End If

End Sub








Private Sub TextBox1_Change()

    'var_1 = TextBox1.Value
    'var_14 = TextBox14.Value

    If Len(TextBox1.Value) = 6 And IsNumeric(TextBox1.Value) = True Then
    
    'occurence
    TextBox7.Value = Application.WorksheetFunction.CountIf(Range(lettre_col_codes & ":" & lettre_col_codes), TextBox1.Value)
    'ligne occurrence pour modif ou delete
    If TextBox7.Value > 0 Then
        Dim i As Integer
        For i = 2 To ligne
            If CLng(TextBox1.Value) = Cells(i, col_codes).Value Then
                ListBox4.AddItem i
                ListBox4.BackColor = vbRed
            End If
        Next i
    End If
    
    'libelle
    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_transco).Range("B:B"), TextBox1.Value) > 0 Then
        TextBox18.Value = Application.WorksheetFunction.VLookup(CLng(Me.TextBox1), Workbooks(base_data).Sheets(onglet_transco).Range("B:N"), 13, 0)
    End If
    
    'pcb
    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_pcb).Range("B:B"), TextBox1.Value) > 0 Then
        TextBox17.Value = Application.WorksheetFunction.VLookup(CLng(Me.TextBox1), Workbooks(base_data).Sheets(onglet_pcb).Range("B:J"), 9, 0)
    End If
    
    'alerte onglet suivi
    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_suivi).Range("C:C"), TextBox1.Value) > 0 Then
        Dim num_ligne As Long
        num_ligne = Workbooks(base_data).Sheets(onglet_suivi).Range("C:C").Find(What:=CLng(TextBox1.Value)).Row
        If Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 3).Value = CLng(TextBox1.Value) And LCase(Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 7).Value) = "alerte" Then
            TextBox8.Value = Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 5).Value
        ElseIf Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 3).Value = CLng(TextBox1.Value) And LCase(Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 7).Value) = "oui" Then
            TextBox8.Font.Size = 10
            TextBox8.Value = "Le code " & Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 3).Value & " va être remplacé par le code " & Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 8).Value
        Else
            TextBox8.Value = ""
        End If
    End If


    'deref
    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_derefs).Range("G:G"), TextBox1.Value) > 0 Then
        Dim ligne_derefs As Long
        ligne_derefs = Workbooks(base_data).Sheets(onglet_derefs).Cells(Application.Rows.Count, 1).End(xlUp).Row
    
        Dim var_annee As Long
        var_annee = Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 3).Value
        Dim var_mois_num As Integer
        If LCase(Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 2).Value) = "janvier" Then
            var_mois_num = 1
        ElseIf Replace(LCase(Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 2).Value), "é", "e") = "fevrier" Then
            var_mois_num = 2
        ElseIf LCase(Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 2).Value) = "mars" Then
            var_mois_num = 3
        ElseIf LCase(Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 2).Value) = "avril" Then
            var_mois_num = 4
        ElseIf LCase(Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 2).Value) = "mai" Then
            var_mois_num = 5
        ElseIf LCase(Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 2).Value) = "juin" Then
            var_mois_num = 6
        ElseIf LCase(Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 2).Value) = "juillet" Then
            var_mois_num = 7
        ElseIf Replace(LCase(Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 2).Value), "û", "u") = "aout" Then
            var_mois_num = 8
        ElseIf LCase(Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 2).Value) = "septembre" Then
            var_mois_num = 9
        ElseIf LCase(Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 2).Value) = "octobre" Then
            var_mois_num = 10
        ElseIf LCase(Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 2).Value) = "novembre" Then
            var_mois_num = 11
        ElseIf Replace(LCase(Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 2).Value), "é", "e") = "decembre" Then
            var_mois_num = 12
        End If
        Dim n As Long
        For n = 2 To ligne_derefs
            If Workbooks(base_data).Sheets(onglet_derefs).Cells(n, 7).Value = CLng(TextBox1.Value) Then
                If var_annee < Year(Workbooks(base_data).Sheets(onglet_derefs).Cells(n, 5).Value) Then
                    TextBox8.Value = "Article DEREF"
                    TextBox8.BackColor = vbRed
                ElseIf var_annee = Year(Workbooks(base_data).Sheets(onglet_derefs).Cells(n, 5).Value) Then
                    If var_mois_num <= Month(Workbooks(base_data).Sheets(onglet_derefs).Cells(n, 5).Value) Then
                        TextBox8.Value = "Article DEREF"
                        TextBox8.BackColor = vbRed
                    End If
                End If
            End If
        Next n
    End If


    
    'ca + couverture
    '    If Application.WorksheetFunction.CountIf(Workbooks(workbook_fichier_base).Sheets("Synthese Article").Range("A:A"), TextBox1.Value) > 0 Then
    '
    '        Dim var_prmp As Long
    '        var_prmp = Application.WorksheetFunction.VLookup(CLng(Me.TextBox1), Sheets("Synthese Article").Range("A:AK"), 37, 0)
    '        Dim var_m_6 As Long
    '        var_m_6 = Application.WorksheetFunction.VLookup(CLng(Me.TextBox1), Sheets("Synthese Article").Range("A:S"), 19, 0)
    '        If IsNumeric(TextBox14.Value) = True And TextBox14.Value <> "" Then
    '            TextBox13.Value = Me.TextBox14 * var_prmp
    '            If var_m_6 <> 0 Then
    '                TextBox15.Value = Round(Me.TextBox14 / var_m_6, 2)
    '            Else
    '                TextBox15.Value = "Division par 0"
    '            End If
    '
    '
    '        Else
    '            TextBox13.Value = ""
    '            TextBox15.Value = ""
    '
    '        End If
    '    End If
        If Application.WorksheetFunction.IfError(Application.VLookup(CLng(TextBox1.Value), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0), 0) Then
            If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_BI).Range("C:C"), Application.WorksheetFunction.VLookup(CLng(TextBox1.Value), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)) > 0 Then
                Dim code_ean As Double
                code_ean = Application.WorksheetFunction.VLookup(CLng(TextBox1.Value), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)
                Dim prmp As Double
                prmp = Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("C:F"), 4, 0)
                Dim nbre_depot As Double
                nbre_depot = Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("C:G"), 5, 0)
                Dim min As Double
                min = 600 * nbre_depot
                Dim somme As Double
                somme = Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("C:T"), 18, 0)
                
                'TextBox15.Value = Round((CDbl(min) * CLng(prmp)) / (CDbl(somme) * CLng(prmp)) / (nbre_depot) * 12, 2)
                
                'TextBox15.Value = Round((TextBox14.Value * prmp) / (somme * prmp) / (nbre_depot) * 12, 2)
                
    '                If IsNumeric(TextBox14.Value) = True And TextBox14.Value <> "" Then
    '                    TextBox15.Value = Round((TextBox14.Value * prmp) / (somme * prmp) / (nbre_depot) * 12, 2)
    '                    TextBox13.Value = prmp * TextBox14.Value
    '                End If
                If IsNumeric(TextBox14.Value) = True And TextBox14.Value <> "" Then
                    If prmp <> 0 Then
    '                        TextBox13.Value = prmp * TextBox14.Value
                            If somme <> 0 Or nbre_depot <> 0 Then
    '                            TextBox15.Value = Round((TextBox14.Value * prmp) / (somme * prmp) / (nbre_depot) * 12, 2)
                                TextBox15.Value = Round((TextBox14.Value * prmp) / ((somme / nbre_depot) * prmp) * 12, 2)
                            Else
                                TextBox15.Value = "Division par 0"
                            End If
                        Else
    '                        TextBox13.Value = "PRMP égal à 0"
                        TextBox15.Value = "Division par 0"
                    End If
                End If
                


            End If
        End If
        
        'CA engagé + Marge
        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_marge).Range("C:C"), TextBox1.Value) > 0 Then
            Dim tx_marge As Double
            tx_marge = Application.WorksheetFunction.VLookup(CLng(TextBox1.Value), Workbooks(base_data).Sheets(onglet_marge).Range("C:E"), 3, 0)
            If tx_marge <> 0 Then
                TextBox20.Value = Round(tx_marge * 100, 2) & " %"
            End If
            If IsNumeric(TextBox14.Value) = True And TextBox14.Value <> "" Then
                Dim prix_vente As Double
                prix_vente = Application.WorksheetFunction.VLookup(CLng(TextBox1.Value), Workbooks(base_data).Sheets(onglet_marge).Range("C:D"), 2, 0)
                If prix_vente <> 0 Then
                    TextBox13.Value = Round(prix_vente * TextBox14.Value, 2) & " €"
                End If
            End If
        End If
            
    
    
    
        'taux de rattachement
        Dim ligne_article As Long
        Dim k As Integer
        Dim l As Double
        Dim m As Long
        Dim cpt_sap As Integer
        Dim cpt_rattachement As Integer
        Dim taux As Long

        For l = 8 To ligne_rattachement
            If Workbooks(base_data).Sheets(onglet_rattachement).Cells(l, col_codes_ratt).Value = CLng(TextBox1.Value) Then
                ligne_article = l
            End If
        Next l

        If ligne_article = 0 Then
            TextBox19.BackColor = vbWhite
            TextBox19.Value = ""
        Else
        
            For k = 1 To ligne_region
                If Workbooks(base_data).Sheets(onglet_code_region).Cells(k, 1).Value = region Then
                    For m = 14 To nbre_col_ratt
                        If Workbooks(base_data).Sheets(onglet_code_region).Cells(k, 2).Value = CLng(Workbooks(base_data).Sheets(onglet_rattachement).Cells(1, m).Value) Then
                            cpt_rattachement = cpt_rattachement + Workbooks(base_data).Sheets(onglet_rattachement).Cells(ligne_article, m).Value
                        End If
                    Next m
                    cpt_sap = cpt_sap + 1
                End If
            Next k

            taux = (cpt_rattachement / cpt_sap) * 100
            If taux < 100 Then
                TextBox19.Value = taux & " %"
    '            TextBox19.BackColor = vbRed
    '            MsgBox "Taux de rattachement inférieur à 100 % pour le code " & TextBox1.Value & " pour la région : " & region
            Else
                TextBox19.Value = taux & " %"
            End If
        
        End If

    Else
        TextBox7.Value = ""
    '    TextBox11.Value = ""
        TextBox13.Value = ""
        TextBox15.Value = ""
        TextBox20.Value = ""
    '    ListBox2.Clear
        ListBox4.Clear
        TextBox19.BackColor = vbWhite
        TextBox19.Value = ""
        TextBox8.BackColor = vbWhite
        TextBox8.Value = ""
        TextBox17.Value = ""
        TextBox18.Value = ""
        ListBox4.BackColor = vbWhite
        


    End If

    'facings
    If IsNumeric(TextBox14.Value) = True And TextBox14.Value <> "" Then

        Dim array_facing(0, 2)
        ListBox3.ColumnCount = 3
        array_facing(0, 0) = Round(TextBox14.Value / 6, 0)
        array_facing(0, 1) = Round(TextBox14.Value, 0)
        array_facing(0, 2) = Round(array_facing(0, 0) * 11, 0)
        ListBox3.List() = array_facing
    Else
        ListBox3.Clear
    End If





End Sub




Private Sub TextBox14_Change()



    If Len(TextBox1.Value) = 6 And IsNumeric(TextBox1.Value) = True Then
        
    'TextBox7.Value = Application.WorksheetFunction.CountIf(Range(lettre_col_codes & ":" & lettre_col_codes), TextBox1.Value)

        '    If Application.WorksheetFunction.CountIf(Workbooks("notes de verification fichier parametrage.xlsx").Sheets("notes fichier 012023 PARAMETRAG").Range("C:C"), TextBox1.Value) > 0 Then
        '        TextBox8.Value = "ARTICLE DEREF"
        '        Me.TextBox8.BackColor = vbRed
        '    End If

        '    If Application.WorksheetFunction.CountIf(Workbooks(workbook_fichier_base).Sheets("Synthese Article").Range("A:A"), TextBox1.Value) > 0 Then
        ''        Dim var_lib As String
        ''        var_lib = Application.WorksheetFunction.VLookup(CLng(Me.TextBox1), Sheets("Synthese Article").Range("A:B"), 2, 0)
        '        Dim var_prmp As Long
        '        var_prmp = Application.WorksheetFunction.VLookup(CLng(Me.TextBox1), Sheets("Synthese Article").Range("A:AK"), 37, 0)
        '        Dim var_m_6 As Long
        '        var_m_6 = Application.WorksheetFunction.VLookup(CLng(Me.TextBox1), Sheets("Synthese Article").Range("A:S"), 19, 0)
        '        'TextBox11.Value = var_lib
        '        'ListBox2.AddItem var_lib
        '        If IsNumeric(TextBox14.Value) = True And TextBox14.Value <> "" Then
        '            TextBox13.Value = Me.TextBox14.Value * var_prmp
        '            If var_m_6 <> 0 Then
        '                TextBox15.Value = Round(Me.TextBox14 / var_m_6, 2)
        '            Else
        '                TextBox15.Value = "Division par 0"
        '            End If
        ''            ListBox2.Clear
        ''            ListBox2.AddItem var_lib
        '
        '        Else
        '            TextBox13.Value = ""
        '            TextBox15.Value = ""
        '            'ListBox2.Clear
        '            'ListBox2.AddItem var_lib
        '
        '        End If
        '    End If
        '    If IsNumeric(TextBox14.Value) = True And TextBox14.Value <> "" Then
        '        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_BI).Range("B:B"), Application.WorksheetFunction.VLookup(CLng(TextBox1.Value), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)) > 0 Then
        '            Dim code_ean As Double
        '            code_ean = Application.WorksheetFunction.VLookup(CLng(TextBox1.Value), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)
        '        'MsgBox Round(Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("B:E"), 4, 0), 2) * TextBox14.Value
        '            TextBox13.Value = Round(Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("B:E"), 4, 0), 2) * TextBox14.Value
        '        End If
        '    Else
        '        TextBox13.Value = ""
        '    End If

            If Application.WorksheetFunction.IfError(Application.VLookup(CLng(TextBox1.Value), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0), 0) Then
                If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_BI).Range("C:C"), Application.WorksheetFunction.VLookup(CLng(TextBox1.Value), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)) > 0 Then
                    Dim code_ean As Double
                    code_ean = Application.WorksheetFunction.VLookup(CLng(TextBox1.Value), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)
                    Dim prmp As Double
                    prmp = Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("C:F"), 4, 0)
                    Dim nbre_depot As Double
                    nbre_depot = Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("C:G"), 5, 0)
                    Dim min As Double
                    min = 600 * nbre_depot
                    Dim somme As Double
                    somme = Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("C:T"), 18, 0)
                    
                    'TextBox15.Value = Round((CDbl(min) * CLng(prmp)) / (CDbl(somme) * CLng(prmp)) / (nbre_depot) * 12, 2)
                    
                    'TextBox15.Value = Round((TextBox14.Value * prmp) / (somme * prmp) / (nbre_depot) * 12, 2)
                    
        '                If IsNumeric(TextBox14.Value) = True And TextBox14.Value <> "" Then
        '                    If prmp <> 0 Or somme <> 0 Or nbre_depot <> 0 Then
        '                        TextBox15.Value = Round((TextBox14.Value * prmp) / (somme * prmp) / (nbre_depot) * 12, 2)
        '                        TextBox13.Value = prmp * TextBox14.Value
        '                    Else
        '
        '
        '                End If
                    
                    If IsNumeric(TextBox14.Value) = True And TextBox14.Value <> "" Then
                        If prmp <> 0 Then
        '                        TextBox13.Value = prmp * TextBox14.Value
                            If somme <> 0 Or nbre_depot <> 0 Then
        '                            TextBox15.Value = Round((TextBox14.Value * prmp) / (somme * prmp) / (nbre_depot) * 12, 2)
                                TextBox15.Value = Round((TextBox14.Value * prmp) / ((somme / nbre_depot) * prmp) * 12, 2)
                            Else
                                TextBox15.Value = "Division par 0"
                            End If
                        Else
        '                        TextBox13.Value = "PRMP égal à 0"
                            TextBox15.Value = "Division par 0"
                        End If
                    End If
                    


                End If
            End If
        
            'CA engagé
            If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_marge).Range("C:C"), TextBox1.Value) > 0 Then
                If IsNumeric(TextBox14.Value) = True And TextBox14.Value <> "" Then
                    Dim prix_vente As Double
                    prix_vente = Application.WorksheetFunction.VLookup(CLng(TextBox1.Value), Workbooks(base_data).Sheets(onglet_marge).Range("C:D"), 2, 0)
                    If prix_vente <> 0 Then
                        TextBox13.Value = Round(prix_vente * TextBox14.Value, 2) & " €"
                    End If
                End If
            End If
        
        

    Else
        TextBox7.Value = ""
        'TextBox11.Value = ""
        TextBox13.Value = ""
        TextBox15.Value = ""
        'ListBox2.Clear


    End If

    If IsNumeric(TextBox14.Value) = True And TextBox14.Value <> "" Then

        Dim array_facing(0, 2)
        ListBox3.ColumnCount = 3
        array_facing(0, 0) = Round(TextBox14.Value / 6, 0)
        array_facing(0, 1) = Round(TextBox14.Value, 0)
        array_facing(0, 2) = Round(array_facing(0, 0) * 11, 0)
        ListBox3.List() = array_facing
    Else
        ListBox3.Clear
    End If


End Sub

Public Function check_typo_code(code As Long) As String
    If Application.WorksheetFunction.CountIf(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Range(lettre_col_codes & ":" & lettre_col_codes), code) > 0 Then
        Dim new_typo As String
        If CheckBox2.Value = True And CheckBox3.Value = False Then
            new_typo = "T2"
        ElseIf CheckBox2.Value = False And CheckBox3.Value = True Then
            new_typo = "T3"
        ElseIf CheckBox2.Value = False And CheckBox3.Value = False Then
            new_typo = "T1"
        Else
            new_typo = "T3"
        End If
        Dim r_new As Integer
        r_new = CInt(Right(new_typo, 1))

        Dim x As Long
        Dim sup_old_typo As Integer
        sup_old_typo = 0
        For x = 2 To ligne
            If Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x, col_codes).Value = CLng(code) Then
                If CLng(Right(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x, col_typo).Value, 1)) > sup_old_typo Then
                    sup_old_typo = CInt(Right(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(x, col_typo).Value, 1))
                End If
            End If

                

        Next x
                
        If sup_old_typo = 3 Then
            MsgBox ("Le code " & code & " est déjà utilisé avec pour typologie T3. Vous ne pouvez pas ajouter ce code.")
    '        Unload UserForm1
    '        UserForm1.Show
    '        check_typo_code = "Le code " & code & " est déjà utilisé avec pour typologie T3. Vous ne pouvez pas ajouter ce code."
            check_typo_code = "3"
        Else
            If r_new = sup_old_typo Then
                MsgBox ("Le code " & code & " avec pour typologie " & "T" & sup_old_typo & " est déjà utilisé" & Chr(13) & Chr(10) & "Veuillez mettre la typolgie " & "T" & sup_old_typo + 1)
    '            flag_typo = 1
    '            Exit Function
                check_typo_code = "sup"
            ElseIf r_new < sup_old_typo Then
                MsgBox ("Le code " & code & " a déjà été ajouter avec une typologie de " & sup_old_typo & ". Veuillez selectionner la typologie suivante : " & "T" & sup_old_typo + 1)
                check_typo_code = "inf"
            End If
        End If
    End If

End Function


Public Sub ajout_data(ligne_ajout)

        'variable ca engage + couverture
        Dim code_ean As Double
        Dim prmp As Double
        Dim nbre_depot As Double
        Dim min As Double
        Dim somme As Double
        Dim tx_marge As Double
        Dim prix_vente As Double
        
        'variable taux de rattachement
        Dim ligne_article As Long
        Dim k As Integer
        Dim l As Long
        Dim m As Long
        Dim cpt_sap As Integer
        Dim cpt_rattachement As Integer
        Dim taux As Long


                        'mois
                        Cells(ligne_ajout, col_mois) = StrConv(Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 2), vbProperCase)
                        'region
                        Cells(ligne_ajout, col_region) = UCase(Workbooks(workbook_fichier_base).Sheets(onglet_info).Cells(3, 4))
                        'min
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min) = TextBox14.Value
                        'facings
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 1) = Round(TextBox14.Value / 6, 0)
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 2) = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 1) * 2
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 3) = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 1) * 3
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 4) = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 1) * 4
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 5) = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 1) * 5
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 6) = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 1) * 6
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 7) = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 1) * 7
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 8) = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 1) * 8
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 9) = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 1) * 9
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 10) = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 1) * 10
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 11) = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_min + 1) * 11
                        
                        'typo
                        If CheckBox2.Value = True And CheckBox3.Value = False Then
                           Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_typo) = "T2"
                        ElseIf CheckBox2.Value = False And CheckBox3.Value = True Then
                            Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_typo) = "T3"
                        ElseIf CheckBox2.Value = False And CheckBox3.Value = False Then
                           Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_typo) = "T1"
                        Else
                            Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_typo) = "T3"
                        End If
                        
                        'libelle/ean
                        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_transco).Range("B:B"), Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes)) > 0 Then
                            Cells(ligne_ajout, col_lib) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes)), Workbooks(base_data).Sheets(onglet_transco).Range("B:N"), 13, 0)
                            Cells(ligne_ajout, col_ean) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes)), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)
                        Else
                            Cells(ligne_ajout, col_lib) = "-"
                            Cells(ligne_ajout, col_ean) = "-"
                        End If
                        
                        'rayon
                        Cells(ligne_ajout, col_rayon) = "R" & Left(Cells(ligne_ajout, col_codes).Value, 1) & "0"
                        
                        'occurrence
                        Cells(ligne_ajout, col_occ).FormulaR1C1 = "=countif(C[" & dif & "],RC[" & dif & "])"
                        
                        'pcb
                        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_pcb).Range("B:B"), Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes)) > 0 Then
                            Cells(ligne_ajout, col_pcb) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes)), Workbooks(base_data).Sheets(onglet_pcb).Range("B:J"), 9, 0)
                        Else
                            Cells(ligne_ajout, col_pcb) = "-"
                        End If
                        
                        'couverture
                        If Application.WorksheetFunction.IfError(Application.VLookup(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0), 0) Then
                            If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_BI).Range("C:C"), Application.WorksheetFunction.VLookup(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)) > 0 Then
                                'Dim code_ean As Double
                                code_ean = Application.WorksheetFunction.VLookup(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)
                                'Dim prmp As Double
                                prmp = Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("C:F"), 4, 0)
                                'Dim nbre_depot As Double
                                nbre_depot = Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("C:G"), 5, 0)
                                'Dim min As Double
                                min = 600 * nbre_depot
                                'Dim somme As Double
                                somme = Application.WorksheetFunction.VLookup(code_ean, Workbooks(base_data).Sheets(onglet_BI).Range("C:T"), 18, 0)
                                If prmp <> 0 Then
                                    If somme <> 0 Or nbre_depot <> 0 Then
                                        Cells(ligne_ajout, col_couv) = Round((TextBox14.Value * prmp) / ((somme / nbre_depot) * prmp) * 12, 2)
                                    Else
                                        Cells(ligne_ajout, col_couv) = "-"
                                    End If
                                Else
                                    Cells(ligne_ajout, col_couv) = "-"
                                End If
                            Else
                                Cells(ligne_ajout, col_couv) = "-"
                            End If
                        Else
                            Cells(ligne_ajout, col_couv) = "-"
                        End If

                
                        
                        'CA engagé + Marge
                        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_marge).Range("C:C"), Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes)) > 0 Then
                
                            tx_marge = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes)), Workbooks(base_data).Sheets(onglet_marge).Range("C:E"), 3, 0)
                            If tx_marge <> 0 Then
                                Cells(ligne_ajout, col_marge) = Round(tx_marge, 2)
                            Else
                                Cells(ligne_ajout, col_marge) = "-"
                            End If
                
                            prix_vente = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes)), Workbooks(base_data).Sheets(onglet_marge).Range("C:D"), 2, 0)
                            If prix_vente <> 0 Then
                                Cells(ligne_ajout, col_ca) = prix_vente * TextBox14.Value
                            Else
                                Cells(ligne_ajout, col_ca) = "-"
                            End If
                
                        End If
                        
                        
                        'taux de rattachement
                        For l = 8 To ligne_rattachement
                            If Workbooks(base_data).Sheets(onglet_rattachement).Cells(l, col_codes_ratt).Value = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_ajout, col_codes) Then
                                ligne_article = l
                            End If
                        Next l

                        If ligne_article = 0 Then
                            Cells(ligne_ajout, col_rattachement) = "-"
                        Else
                            For k = 1 To ligne_region
                                If Workbooks(base_data).Sheets(onglet_code_region).Cells(k, 1).Value = region Then
                                    For m = 14 To nbre_col_ratt
                                        If Workbooks(base_data).Sheets(onglet_code_region).Cells(k, 2).Value = CLng(Workbooks(base_data).Sheets(onglet_rattachement).Cells(1, m).Value) Then
                                            cpt_rattachement = cpt_rattachement + Workbooks(base_data).Sheets(onglet_rattachement).Cells(ligne_article, m).Value
                                        End If
                                    Next m
                                    cpt_sap = cpt_sap + 1
                                End If
                            Next k

                            taux = (cpt_rattachement / cpt_sap) * 100
                            Cells(ligne_ajout, col_rattachement) = taux / 100
                        End If



End Sub



Private Sub SUBMIT_Click()

    Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Unprotect Password:=motperdu

    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_suivi).Range("C:C"), TextBox1.Value) > 0 Then
        Dim o As Long
        For o = 2 To ligne_suivi
            If Workbooks(base_data).Sheets(onglet_suivi).Cells(o, 3).Value = CLng(TextBox1.Value) And LCase(Workbooks(base_data).Sheets(onglet_suivi).Cells(o, 7).Value) = "oui" Then
                flag_remplacement = flag_remplacement + 1
                ligne_remplacement = o
                code_remplacement = Workbooks(base_data).Sheets(onglet_suivi).Cells(o, 8).Value
            End If
        Next o
    End If

    'On a la ligne ou y a le code dans l'onglet suivi donc correct
    'num_ligne = Workbooks(base_data).Sheets(onglet_suivi).Range("C:C").Find(What:=CLng(TextBox1.Value)).Row

    'Si ca apparait dans multifournisseur
    Dim flag_fourn As Integer
    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_multifourn).Range("A:A"), TextBox1.Value) > 0 Then
        Dim i_multifourn As Long
        For i_multifourn = 2 To nbre_ligne_multifourn
            If Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 1).Value = CLng(TextBox1.Value) And Replace(UCase(CStr(Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 6).Value)), " ", "") = Replace(region, " ", "") Then
                Dim ligne_multifourn
                ligne_multifourn = i_multifourn
                flag_fourn = 1
            End If
        Next i_multifourn
    End If


    If TextBox1.Value = "" Or TextBox14.Value = "" Then
        MsgBox "Veuillez remplir le code, le min, ou cocher la typologie 1 pour ajouter la donnée"
    ElseIf IsNumeric(TextBox1.Value) = False Or IsNumeric(TextBox14.Value) = False Then
        MsgBox "veuillez remplir un code numeric"
    ElseIf Len(TextBox1.Value) <> 6 Then
        MsgBox "veuillez remplir un code à 6 chiffres"

    Else

                
        Dim result As String

        If flag_fourn = 1 Then
                MsgBox "MULTIFOURNISSEURS"
                'fournisseur ligne actuel + avant
                    Dim num_ligne_1_multifourn As Long
                    num_ligne_1_multifourn = ligne_multifourn
                    While Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_1_multifourn, 5).Value = "fournisseur"
                        Dim ligne_suiv_1_multifourn As Long
                        ligne_suiv_1_multifourn = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Range(lettre_col_codes & "45678").End(xlUp).Row + 1
                        
                        'codes
                        
                        'check typologie différente pour un code déjà remplie
                        result = check_typo_code(Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_1_multifourn, 1).Value)
                        If result = "3" Then
                            Exit Sub
                        ElseIf result = "sup" Then
                            Exit Sub
                        ElseIf result = "inf" Then
                            Exit Sub
                        End If
                        
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_suiv_1_multifourn, col_codes) = Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_1_multifourn, 1).Value
                        Range(lettre_col_codes & ligne_suiv_1_multifourn).Font.Color = -65434 'violet
                        
                        Call ajout_data(ligne_suiv_1_multifourn)

                        
                        
                        num_ligne_1_multifourn = num_ligne_1_multifourn - 1
                    Wend
                    
                    'fournisseur ligne suivante
                    Dim num_ligne_plus_1_multifourn As Long
                    num_ligne_plus_1_multifourn = ligne_multifourn
                    While Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_plus_1_multifourn + 1, 5).Value = "fournisseur"
                        Dim ligne_suiv_plus_1_multifourn As Long
                        ligne_suiv_plus_1_multifourn = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Range(lettre_col_codes & "45678").End(xlUp).Row + 1
                        
                        'codes
                                                
                        'check typologie différente pour un code déjà remplie
                        result = check_typo_code(Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_plus_1_multifourn + 1, 1).Value)
                        If result = "3" Then
                            Exit Sub
                        ElseIf result = "sup" Then
                            Exit Sub
                        ElseIf result = "inf" Then
                            Exit Sub
                        End If
                        
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_suiv_plus_1_multifourn, col_codes) = Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_plus_1_multifourn + 1, 1).Value
                        Range(lettre_col_codes & ligne_suiv_plus_1_multifourn).Font.Color = -65434 'violet
                        
                        Call ajout_data(ligne_suiv_plus_1_multifourn)
                        
                        
                        num_ligne_plus_1_multifourn = num_ligne_plus_1_multifourn + 1
                        
                    Wend
                                            
                    
                    Unload UserForm1
                    UserForm1.Show
        


        ElseIf Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_lots).Range("A:A"), TextBox1.Value) > 0 Then
    
                
                Dim num_ligne As Long
                num_ligne = Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=TextBox1.Value).Row
                If Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne, 5).Value = "Lot" Then

                    MsgBox "LOT"

                    While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne + 1, 5).Value = "Composant"
                        Dim ligne_suiv As Long
                        ligne_suiv = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Range(lettre_col_codes & "45678").End(xlUp).Row + 1
                        
                        'codes
                                
                        'check typologie différente pour un code déjà remplie
                        result = check_typo_code(Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne + 1, 1).Value)
                        If result = "3" Then
                            Exit Sub
                        ElseIf result = "sup" Then
                            Exit Sub
                        ElseIf result = "inf" Then
                            Exit Sub
                        End If
                                               
                        
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_suiv, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne + 1, 1).Value
                        
                        Range(lettre_col_codes & ligne_suiv).Font.Color = -65281 'pink
                        
                        Call ajout_data(ligne_suiv)
                        
                        num_ligne = num_ligne + 1
                    Wend
                    
                    Unload UserForm1
                    UserForm1.Show
                ElseIf Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne, 5).Value = "Composant" Then
                    MsgBox "Composant"
                    
                    
                    'composant ligne actuel + avant
                    Dim num_ligne_1 As Long
                    num_ligne_1 = Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=TextBox1.Value).Row
                    While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_1, 5).Value = "Composant"
                        Dim ligne_suiv_1 As Long
                        ligne_suiv_1 = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Range(lettre_col_codes & "45678").End(xlUp).Row + 1
                        
                        'codes
                        
                        'check typologie différente pour un code déjà remplie
                        result = check_typo_code(Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_1, 1).Value)
                        If result = "3" Then
                            Exit Sub
                        ElseIf result = "sup" Then
                            Exit Sub
                        ElseIf result = "inf" Then
                            Exit Sub
                        End If
                        
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_suiv_1, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_1, 1).Value
                        Range(lettre_col_codes & ligne_suiv_1).Font.Color = -65281 'pink
                        
                        Call ajout_data(ligne_suiv_1)
                        
                        
                        num_ligne_1 = num_ligne_1 - 1
                    Wend
                    
                    'composant ligne suivante
                    Dim num_ligne_plus_1 As Long
                    num_ligne_plus_1 = Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=TextBox1.Value).Row
                    While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_plus_1 + 1, 5).Value = "Composant"
                        Dim ligne_suiv_plus_1 As Long
                        ligne_suiv_plus_1 = Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Range(lettre_col_codes & "45678").End(xlUp).Row + 1
                        
                        'codes
                                                
                        'check typologie différente pour un code déjà remplie
                        result = check_typo_code(Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_plus_1 + 1, 1).Value)
                        If result = "3" Then
                            Exit Sub
                        ElseIf result = "sup" Then
                            Exit Sub
                        ElseIf result = "inf" Then
                            Exit Sub
                        End If
                        
                        Workbooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_suiv_plus_1, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_plus_1 + 1, 1).Value
                        Range(lettre_col_codes & ligne_suiv_plus_1).Font.Color = -65281 'pink
                        
                        Call ajout_data(ligne_suiv_plus_1)
                        
                        
                        num_ligne_plus_1 = num_ligne_plus_1 + 1
                        
                    Wend
                                            
                    
                    Unload UserForm1
                    UserForm1.Show
                End If
        Else
    

    
            'cas article deref mais pas cas lot/comp/suivi delphine -> mettre code en rose

            If flag_remplacement > 0 Then
            
                'check typologie différente pour un code déjà remplie
                result = check_typo_code(Workbooks(base_data).Sheets(onglet_suivi).Cells(ligne_remplacement, 8).Value)
                If result = "3" Then
                    Exit Sub
                ElseIf result = "sup" Then
                    Exit Sub
                ElseIf result = "inf" Then
                    Exit Sub
                End If
                'CHANGEMENTTEST
                'Cells(ligne, col_codes).Value = Workbooks(base_data).Sheets(onglet_suivi).Cells(ligne_remplacement, 8).Value
                Cells(ligne, col_codes).Value = code_remplacement
                Range(lettre_col_codes & ligne).Font.Color = -65281 'pink
                
            ElseIf TextBox8.Value <> "" Then
            
                'check typologie différente pour un code déjà remplie
                result = check_typo_code(TextBox1.Value)
                If result = "3" Then
                    Exit Sub
                ElseIf result = "sup" Then
                    Exit Sub
                ElseIf result = "inf" Then
                    Exit Sub
                End If
            
                'codes
                Cells(ligne, col_codes) = TextBox1.Value
                Range(lettre_col_codes & ligne).Font.Color = -65281 'pink
                
            Else
                'check typologie différente pour un code déjà remplie
                result = check_typo_code(TextBox1.Value)
                If result = "3" Then
                    Exit Sub
                ElseIf result = "sup" Then
                    Exit Sub
                ElseIf result = "inf" Then
                    Exit Sub
                End If
                
                Cells(ligne, col_codes) = TextBox1.Value
                
            End If
                    
            Call ajout_data(ligne)
            
            'trouver la ligne qui vient d'etre ajoutée
            ligne_modif_code = WorkBooks(workbook_fichier_base).Sheets(onglet_fichier_base).Range("D:D").Find(What:=CLng(TextBox1.Value)).Row
            If ligne_modif_code <> 0 Then
                'Récupérer le code à changer
                num_ligne = Workbooks(base_data).Sheets(onglet_suivi).Range("C:C").Find(What:=CLng(TextBox1.Value)).Row
                code_remplacement = Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 8).Value

                'Modifier le code de la ligne
                anciencode = WorkBooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_modif_code, 4).Value
                WorkBooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_modif_code, 4).AddComment "Le code "+ anciencode + "a été remplacé"
                WorkBooks(workbook_fichier_base).Sheets(onglet_fichier_base).Cells(ligne_modif_code, 4).Value = code_remplacement

            End If
            Unload UserForm1
            
            UserForm1.Show
        End If

            

    End If

End Sub





Private Sub CommandButton2_Click()
    Dim flag_fourn As Integer

    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_multifourn).Range("A:A"), TextBox1.Value) > 0 Then
        Dim i_multifourn As Long
        For i_multifourn = 2 To nbre_ligne_multifourn
            If Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 1).Value = CLng(TextBox1.Value) And Replace(UCase(CStr(Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 6).Value)), " ", "") = Replace(region, " ", "") Then
                Dim ligne_multifourn
                ligne_multifourn = i_multifourn
                flag_fourn = 1
            End If
        Next i_multifourn
    End If
    MsgBox flag_fourn & ligne_multifourn
End Sub

Private Sub CommandButton3_Click()
    MsgBox col_flag_siege
End Sub










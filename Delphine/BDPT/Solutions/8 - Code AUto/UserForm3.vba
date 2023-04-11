
Option Explicit


Public workbook_compil As String
Public onglet_compil As String

'Public emplacement_fichier_suivi As String
'Public fichier_suivi As String

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
Public onglet_entrepots As String

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

Public ligne_region As Long
Public ligne_rattachement As Long
Public nbre_col_ratt As Long

Public ligne_suivi As Long
Public flag_remplacement As Integer
Public ligne_remplacement As Long
Public col_codes_ratt As Long

Public lettre_col_codes As String

Public last_row_ajout As Long




Private Sub UserForm_Initialize()




    workbook_compil = ActiveWorkbook.Name
    onglet_compil = "COMPIL"

    'emplacement_fichier_suivi = "C:\Users\dsi\Documents\"
    ''filename_fichier_suivi = "C:\Users\dsi\Documents\Fichier_suivi.xlsx"
    'fichier_suivi = "Fichier_suivi.xlsx"

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
    onglet_entrepots = "entrepôts"

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

    CheckBox1.Value = True
    CheckBox1.Enabled = False


End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '    Windows(fichier_suivi).Activate
    '    ActiveWorkbook.Save
    '    ActiveWindow.Close

    Unload Me


End Sub

Private Sub fermer_Click()


    Unload Me
End Sub

Private Sub all_region_Click()

    If all_region.Value = True Then
        check_centre_aquitaine = True
        check_centre_atlantique = True
        check_centre_est = True
        check_est = True
        check_idf = True
        check_nord_est = True
        check_nord = True
        check_normandie = True
        check_ouest = True
        check_rhone_alpes = True
        check_sud_est = True
        check_sud_ouest = True
    Else
        check_centre_aquitaine = False
        check_centre_atlantique = False
        check_centre_est = False
        check_est = False
        check_idf = False
        check_nord_est = False
        check_nord = False
        check_normandie = False
        check_ouest = False
        check_rhone_alpes = False
        check_sud_est = False
        check_sud_ouest = False
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
        
            'libelle
        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_transco).Range("B:B"), TextBox4.Value) > 0 Then
            TextBox18.Value = Application.WorksheetFunction.VLookup(CLng(Me.TextBox4), Workbooks(base_data).Sheets(onglet_transco).Range("B:N"), 13, 0)
        End If
        
        'pcb
        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_pcb).Range("B:B"), TextBox4.Value) > 0 Then
            TextBox17.Value = Application.WorksheetFunction.VLookup(CLng(Me.TextBox4), Workbooks(base_data).Sheets(onglet_pcb).Range("B:J"), 9, 0)
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
        
        'couverture
        If Application.WorksheetFunction.IfError(Application.VLookup(CLng(TextBox4.Value), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0), 0) Then
                If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_BI).Range("C:C"), Application.WorksheetFunction.VLookup(CLng(TextBox4.Value), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)) > 0 Then
                    Dim code_ean As Double
                    code_ean = Application.WorksheetFunction.VLookup(CLng(TextBox4.Value), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)
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
                                TextBox15.Value = Round((TextBox14.Value * prmp) / (somme * prmp) / (nbre_depot) * 12, 2)
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
            If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_marge).Range("C:C"), TextBox4.Value) > 0 Then
    '            Dim tx_marge As Double
    '            tx_marge = Application.WorksheetFunction.VLookup(CLng(TextBox4.Value), Workbooks(base_data).Sheets(onglet_marge).Range("C:E"), 3, 0)
    '            If tx_marge <> 0 Then
    '                TextBox20.Value = Round(tx_marge * 100, 2) & " %"
    '            End If
                If IsNumeric(TextBox14.Value) = True And TextBox14.Value <> "" Then
                    Dim prix_vente As Double
                    prix_vente = Application.WorksheetFunction.VLookup(CLng(TextBox4.Value), Workbooks(base_data).Sheets(onglet_marge).Range("C:D"), 2, 0)
                    If prix_vente <> 0 Then
                        TextBox13.Value = Round(prix_vente * TextBox14.Value, 2) & " €"
                    End If
                End If
            End If
        
        
        
    Else
        TextBox5.BackColor = vbWhite
        TextBox5.Value = ""
        TextBox5.Font.Size = 10
        ListBox2.Clear
        
        TextBox13.Value = ""
        TextBox15.Value = ""
    '    TextBox20.Value = ""

    End If
End Sub


Private Sub TextBox14_Change()



    If Len(TextBox4.Value) = 6 And IsNumeric(TextBox4.Value) = True Then


            If Application.WorksheetFunction.IfError(Application.VLookup(CLng(TextBox4.Value), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0), 0) Then
                If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_BI).Range("C:C"), Application.WorksheetFunction.VLookup(CLng(TextBox4.Value), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)) > 0 Then
                    Dim code_ean As Double
                    code_ean = Application.WorksheetFunction.VLookup(CLng(TextBox4.Value), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)
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
                                TextBox15.Value = Round((TextBox14.Value * prmp) / (somme * prmp) / (nbre_depot) * 12, 2)
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
            If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_marge).Range("C:C"), TextBox4.Value) > 0 Then
                If IsNumeric(TextBox14.Value) = True And TextBox14.Value <> "" Then
                    Dim prix_vente As Double
                    prix_vente = Application.WorksheetFunction.VLookup(CLng(TextBox4.Value), Workbooks(base_data).Sheets(onglet_marge).Range("C:D"), 2, 0)
                    If prix_vente <> 0 Then
                        TextBox13.Value = Round(prix_vente * TextBox14.Value, 2) & " €"
                    End If
                End If
            End If
        
        

    Else
    '    TextBox7.Value = ""
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



Public Function check_typo_code(code As Long, region As String) As String
    If Application.WorksheetFunction.CountIf(Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & ":" & lettre_col_codes), code) > 0 Then
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
            If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x, col_codes).Value = CLng(code) And Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x, col_region).Value = CStr(region) Then
                If CLng(Right(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x, col_typo).Value, 1)) > sup_old_typo Then
                    sup_old_typo = CInt(Right(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(x, col_typo).Value, 1))
                End If
            End If

                

        Next x
                
        If sup_old_typo = 3 Then
            MsgBox ("Le code " & code & " pour région " & region & " est déjà utilisé avec pour typologie T3. Vous ne pouvez pas ajouter ce code.")

            check_typo_code = "3"
        Else
            If r_new = sup_old_typo Then
                MsgBox ("Le code " & code & " pour région " & region & " avec pour typologie " & "T" & sup_old_typo & " est déjà utilisé" & Chr(13) & Chr(10) & "Veuillez mettre la typolgie " & "T" & sup_old_typo + 1)

                check_typo_code = "sup"
            ElseIf r_new < sup_old_typo Then
                MsgBox ("Le code " & code & " pour région " & region & " a déjà été ajouter avec une typologie de " & sup_old_typo & ". Veuillez selectionner la typologie suivante : " & "T" & sup_old_typo + 1)
                check_typo_code = "inf"
            End If
        End If
    End If

End Function



Public Sub ajout_data(ligne_ajout, region)

        'variable ca engage + couverture
        Dim code_ean As Double
        Dim prmp As Double
        Dim nbre_depot As Double
        Dim min As Double
        Dim somme As Double
        Dim prix_vente As Double
        

        Dim nbre_depot_procedure_region As Long
        nbre_depot_procedure_region = Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_code_region).Range("A:A"), region)


                        'mois
                        Cells(ligne_ajout, col_mois) = mois_lettre
                        
                        'region
                        Cells(ligne_ajout, col_region) = region
                        
                        'typo
                        If CheckBox2.Value = True And CheckBox3.Value = False Then
                           Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_typo) = "T1 + T2"
                        ElseIf CheckBox2.Value = False And CheckBox3.Value = True Then
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_typo) = "T1 + T2 + T3"
                        ElseIf CheckBox2.Value = False And CheckBox3.Value = False Then
                           Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_typo) = "T1"
                        Else
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_typo) = "T1 + T2 + T3"
                        End If
                        
                        'easier
                        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_transco).Range("B:B"), Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes)) > 0 Then
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_easier) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes)), Workbooks(base_data).Sheets(onglet_transco).Range("B:C"), 2, 0)
                        Else
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_easier) = "-"
                        End If
                        'easier conc
                        If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_easier) = "-" Then
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_easier_conc) = "-"
                        Else
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_easier_conc) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_easier) & "EA"
                        End If
                        
                        'libelle/ean
                        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_transco).Range("B:B"), Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes)) > 0 Then
                            Cells(ligne_ajout, col_ean) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes)), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)
                            Cells(ligne_ajout, col_lib) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes)), Workbooks(base_data).Sheets(onglet_transco).Range("B:N"), 13, 0)
                        Else
                            Cells(ligne_ajout, col_ean) = "-"
                            Cells(ligne_ajout, col_lib) = "-"
                        End If
                        
                        'marche/cat/sous cat
                        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_souscat).Range("A:A"), Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes)) > 0 Then
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_marche) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes)), Workbooks(base_data).Sheets(onglet_souscat).Range("A:D"), 4, 0)
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_cat) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes)), Workbooks(base_data).Sheets(onglet_souscat).Range("A:E"), 5, 0)
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_souscat) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes)), Workbooks(base_data).Sheets(onglet_souscat).Range("A:F"), 6, 0)
                        Else
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_marche) = "-"
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_cat) = "-"
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_souscat) = "-"
                        End If
                        
                        'pcb
                        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_pcb).Range("B:B"), Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes)) > 0 Then
                            Cells(ligne_ajout, col_pcb) = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes)), Workbooks(base_data).Sheets(onglet_pcb).Range("B:J"), 9, 0)
                        Else
                            Cells(ligne_ajout, col_pcb) = "-"
                        End If
                                                                                            
                        'min
                        Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_minsouh) = TextBox14.Value
                        
                        'Min tot region
                        Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_mintot) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_minsouh) * nbre_depot_procedure_region
                        'Nombre de facing merch
                        Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_nbrefac) = 6
                        
                        'min facing
                        Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_minfac) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_minsouh) / 6, 2)
                        
                        
                        'CA engagé + Marge
                        If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_marge).Range("C:C"), Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes)) > 0 Then
                
                            prix_vente = Application.WorksheetFunction.VLookup(CLng(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes)), Workbooks(base_data).Sheets(onglet_marge).Range("C:D"), 2, 0)
                            If prix_vente <> 0 Then
                                Cells(ligne_ajout, col_ca) = prix_vente * TextBox14.Value
                            Else
                                Cells(ligne_ajout, col_ca) = "-"
                            End If
                
                        End If
                        

                        
                        'couverture
                        If Application.WorksheetFunction.IfError(Application.VLookup(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0), 0) Then
                            If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_BI).Range("C:C"), Application.WorksheetFunction.VLookup(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)) > 0 Then
                                'Dim code_ean As Double
                                code_ean = Application.WorksheetFunction.VLookup(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_codes), Workbooks(base_data).Sheets(onglet_transco).Range("B:D"), 3, 0)
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
                                        Cells(ligne_ajout, col_couv) = Round((TextBox14.Value * prmp) / (somme * prmp) / (nbre_depot) * 12, 2)
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
                        
                       
                        'entrepot
                        Dim ligne_entrepot_procedure_region As Long
                        ligne_entrepot_procedure_region = Workbooks(base_data).Sheets(onglet_code_region).Range("A:A").Find(What:=region).Row
                        Dim SAP_procedure_region As Double
                        SAP_procedure_region = Workbooks(base_data).Sheets(onglet_code_region).Cells(ligne_entrepot_procedure_region, 2)
                    '    Dim entrepot_procedure_region As String
                        Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_ajout, col_entrepot) = Application.WorksheetFunction.VLookup(SAP_procedure_region, Workbooks(base_data).Sheets(onglet_entrepots).Range("A:F"), 6, 0)
                                   

                        
                        




End Sub



Sub procedure_ajout_button(code As Long, region As String)

    If Application.WorksheetFunction.CountIf(Workbooks(workbook_compil).Sheets(onglet_compil).Range("B:B"), region) > 0 Then
        last_row_ajout = Workbooks(workbook_compil).Sheets(onglet_compil).Range("B:B").Find(What:=region).Row
        While Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_region) = region
            last_row_ajout = last_row_ajout + 1
        Wend

    Else
        MsgBox "La région " & region & " n'est pas remplie dans l'onglet Compil"
        Exit Sub
    End If

        


    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_suivi).Range("C:C"), TextBox4.Value) > 0 Then
        Dim o As Long
        For o = 2 To ligne_suivi
            If Workbooks(base_data).Sheets(onglet_suivi).Cells(o, 3).Value = CLng(TextBox4.Value) And LCase(Workbooks(base_data).Sheets(onglet_suivi).Cells(o, 7).Value) = "oui" Then
                flag_remplacement = flag_remplacement + 1
                ligne_remplacement = o
            End If
        Next o
    End If


    Dim flag_fourn As Integer
    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_multifourn).Range("A:A"), TextBox4.Value) > 0 Then
        Dim i_multifourn As Long
        For i_multifourn = 2 To nbre_ligne_multifourn
            If Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 1).Value = CLng(TextBox4.Value) And Replace(UCase(CStr(Workbooks(base_data).Sheets(onglet_multifourn).Cells(i_multifourn, 6).Value)), " ", "") = Replace(region, " ", "") Then
                Dim ligne_multifourn
                ligne_multifourn = i_multifourn
                flag_fourn = 1
            End If
        Next i_multifourn
    End If

            Dim result As String


            If flag_fourn = 1 Then
                    MsgBox "MULTIFOURNISSEURS"
                    'fournisseur ligne actuel + avant
                        Dim num_ligne_1_multifourn As Long
                        num_ligne_1_multifourn = ligne_multifourn
                        While Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_1_multifourn, 5).Value = "fournisseur"
                            
                            
                            'check typologie différente pour un code déjà remplie
                            result = check_typo_code(Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_1_multifourn, 1).Value, region)
                            If result = "3" Then
                                Exit Sub
                            ElseIf result = "sup" Then
                                Exit Sub
                            ElseIf result = "inf" Then
                                Exit Sub
                            End If
                                                
                                                
                            'Modification codes onglet compil
                            Rows(last_row_ajout & ":" & last_row_ajout).Insert Shift:=xlDown
                            'codes
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes) = Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_1_multifourn, 1).Value
                            Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & last_row_ajout).Font.Color = -65434 'violet
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).AddComment
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).Comment.Text Text:="Fournisseur"
                            
                            Call ajout_data(last_row_ajout, region)

                            
                            
                            num_ligne_1_multifourn = num_ligne_1_multifourn - 1
                        Wend

                        'fournisseur ligne suivante
                        Dim num_ligne_plus_1_multifourn As Long
                        num_ligne_plus_1_multifourn = ligne_multifourn
                        While Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_plus_1_multifourn + 1, 5).Value = "fournisseur"
                        
                            

                            'check typologie différente pour un code déjà remplie
                            result = check_typo_code(Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_plus_1_multifourn + 1, 1).Value, region)
                            If result = "3" Then
                                Exit Sub
                            ElseIf result = "sup" Then
                                Exit Sub
                            ElseIf result = "inf" Then
                                Exit Sub
                            End If

                            'Modification codes onglet compil
                            Rows(last_row_ajout & ":" & last_row_ajout).Insert Shift:=xlDown
                            'codes
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes) = Workbooks(base_data).Sheets(onglet_multifourn).Cells(num_ligne_plus_1_multifourn, 1).Value
                            Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & last_row_ajout).Font.Color = -65434 'violet
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).AddComment
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).Comment.Text Text:="Fournisseur"
                            
                            Call ajout_data(last_row_ajout, region)


                            num_ligne_plus_1_multifourn = num_ligne_plus_1_multifourn + 1

                        Wend


                        Unload UserForm3
                        UserForm3.Show



            ElseIf Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_lots).Range("A:A"), code) > 0 Then


                    Dim num_ligne As Long
                    num_ligne = Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=code).Row
                    If Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne, 5).Value = "Lot" Then

                        MsgBox "LOT"

                        While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne + 1, 5).Value = "Composant"


                            'check typologie différente pour un code déjà remplie
                            result = check_typo_code(Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne + 1, 1).Value, region)
                            If result = "3" Then
                                Exit Sub
                            ElseIf result = "sup" Then
                                Exit Sub
                            ElseIf result = "inf" Then
                                Exit Sub
                            End If

                            'Modification codes onglet compil
                            Rows(last_row_ajout & ":" & last_row_ajout).Insert Shift:=xlDown
                            'codes
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne + 1, 1).Value
                            Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & last_row_ajout).Font.Color = -65281 'pink
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).AddComment
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).Comment.Text Text:="Code Composant"

                            Call ajout_data(last_row_ajout, region)

                            num_ligne = num_ligne + 1
                        Wend


                    ElseIf Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne, 5).Value = "Composant" Then
                        MsgBox "Composant"


                        'composant ligne actuel + avant
                        Dim num_ligne_1 As Long
                        num_ligne_1 = Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=code).Row
                        While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_1, 5).Value = "Composant"


                            'check typologie différente pour un code déjà remplie
                            result = check_typo_code(Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_1, 1).Value, region)
                            If result = "3" Then
                                Exit Sub
                            ElseIf result = "sup" Then
                                Exit Sub
                            ElseIf result = "inf" Then
                                Exit Sub
                            End If

                            'Modification codes onglet compil
                            Rows(last_row_ajout & ":" & last_row_ajout).Insert Shift:=xlDown
                            'codes
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_1, 1).Value
                            Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & last_row_ajout).Font.Color = -65281 'pink
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).AddComment
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).Comment.Text Text:="Code Composant"

                            Call ajout_data(last_row_ajout, region)


                            num_ligne_1 = num_ligne_1 - 1
                        Wend

                        'composant ligne suivante
                        Dim num_ligne_plus_1 As Long
                        num_ligne_plus_1 = Workbooks(base_data).Sheets(onglet_lots).Range("A:A").Find(What:=code).Row
                        While Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_plus_1 + 1, 5).Value = "Composant"


                            'check typologie différente pour un code déjà remplie
                            result = check_typo_code(Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_plus_1 + 1, 1).Value, region)
                            If result = "3" Then
                                Exit Sub
                            ElseIf result = "sup" Then
                                Exit Sub
                            ElseIf result = "inf" Then
                                Exit Sub
                            End If

                            'Modification codes onglet compil
                            Rows(last_row_ajout & ":" & last_row_ajout).Insert Shift:=xlDown
                            'codes
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes) = Workbooks(base_data).Sheets(onglet_lots).Cells(num_ligne_plus_1 + 1, 1).Value
                            Workbooks(workbook_compil).Sheets(onglet_compil).Range(lettre_col_codes & last_row_ajout).Font.Color = -65281 'pink
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).AddComment
                            Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).Comment.Text Text:="Code Composant"

                            Call ajout_data(last_row_ajout, region)


                            num_ligne_plus_1 = num_ligne_plus_1 + 1

                        Wend



                    End If
            Else



                'cas article deref mais pas cas lot/comp/suivi delphine -> mettre code en rose

                If flag_remplacement > 0 Then

                    'check typologie différente pour un code déjà remplie
                    result = check_typo_code(Workbooks(base_data).Sheets(onglet_suivi).Cells(ligne_remplacement, 8).Value, region)
                    If result = "3" Then
                        Exit Sub
                    ElseIf result = "sup" Then
                        Exit Sub
                    ElseIf result = "inf" Then
                        Exit Sub
                    End If

                    'Modification codes onglet compil
                    Rows(last_row_ajout & ":" & last_row_ajout).Insert Shift:=xlDown
                    'Code
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).Value = Workbooks(base_data).Sheets(onglet_suivi).Cells(ligne_remplacement, 8).Value
                    Range(lettre_col_codes & last_row_ajout).Font.Color = -65281 'pink
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).AddComment
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).Comment.Text Text:="Le code " & TextBox4.Value & " a été remplacé par le code " & Workbooks(base_data).Sheets(onglet_suivi).Cells(ligne_remplacement, 8).Value
                            
                ElseIf TextBox5.Value <> "" Then

                    'check typologie différente pour un code déjà remplie
                    result = check_typo_code(code, region)
                    If result = "3" Then
                        Exit Sub
                    ElseIf result = "sup" Then
                        Exit Sub
                    ElseIf result = "inf" Then
                        Exit Sub
                    End If

                    'Modification codes onglet compil
                    Rows(last_row_ajout & ":" & last_row_ajout).Insert Shift:=xlDown
                    'Code
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).Value = code
                    Range(lettre_col_codes & last_row_ajout).Font.Color = -65281 'pink
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).AddComment
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).Comment.Text Text:="Article DEREF"


                Else
                    'check typologie différente pour un code déjà remplie
                    result = check_typo_code(code, region)
                    If result = "3" Then
                        Exit Sub
                    ElseIf result = "sup" Then
                        Exit Sub
                    ElseIf result = "inf" Then
                        Exit Sub
                    End If

                    'Code
                    Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).Value = code

                End If

                Call ajout_data(last_row_ajout, region)


            End If
            'Remplacement du code

            'ligne_modif et nouveau code à remplacer

            num_ligne = Workbooks(base_data).Sheets(onglet_suivi).Range("H:H").Find(What:=CLng(TextBox4.Value)).Row
            ancien_code = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).Value
            code_remplacement = Workbooks(base_data).Sheets(onglet_suivi).Cells(num_ligne, 8).Value

            'Modif du code
            WorkBooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).AddComment "Le code "+ ancien_code + "a été remplacé"
            WorkBooks(workbook_compil).Sheets(onglet_compil).Cells(last_row_ajout, col_codes).Value = code_remplacement

End Sub

Private Sub ajouter_button_Click()

    If check_centre_aquitaine = False And check_centre_atlantique = False And check_centre_est = False And check_est = False And check_idf = False And check_nord_est = False And check_nord = False And check_normandie = False And check_ouest = False And check_rhone_alpes = False And check_sud_est = False And check_sud_ouest = False Then
        MsgBox "Veuillez selectionner au moins une région"
    End If






    If TextBox4.Value = "" Or TextBox14.Value = "" Then
        MsgBox "Veuillez remplir le code, le min, ou cocher la typologie 1 pour ajouter la donnée"
    ElseIf IsNumeric(TextBox4.Value) = False Or IsNumeric(TextBox14.Value) = False Then
        MsgBox "veuillez remplir un code numeric"
    ElseIf Len(TextBox4.Value) <> 6 Then
        MsgBox "veuillez remplir un code à 6 chiffres"

    Else


        Dim region_click As String
        If check_centre_aquitaine = True Then
            region_click = "CENTRE AQUITAINE"
            Call procedure_ajout_button(TextBox4.Value, region_click)
        End If
        If check_centre_atlantique = True Then
            region_click = "CENTRE ATLANTIQUE"
            Call procedure_ajout_button(TextBox4.Value, region_click)
        End If
        If check_centre_est = True Then
            region_click = "CENTRE EST"
            Call procedure_ajout_button(TextBox4.Value, region_click)
        End If
        If check_est = True Then
            region_click = "EST"
            Call procedure_ajout_button(TextBox4.Value, region_click)
        End If
        If check_idf = True Then
            region_click = "IDF"
            Call procedure_ajout_button(TextBox4.Value, region_click)
        End If
        If check_nord_est = True Then
            region_click = "NORD EST"
            Call procedure_ajout_button(TextBox4.Value, region_click)
        End If
        If check_nord = True Then
            region_click = "NORD"
            Call procedure_ajout_button(TextBox4.Value, region_click)
        End If
        If check_normandie = True Then
            region_click = "NORMANDIE"
            Call procedure_ajout_button(TextBox4.Value, region_click)
        End If
        If check_ouest = True Then
            region_click = "OUEST"
            Call procedure_ajout_button(TextBox4.Value, region_click)
        End If
        If check_rhone_alpes = True Then
            region_click = "RHONE ALPES"
            Call procedure_ajout_button(TextBox4.Value, region_click)
        End If
        If check_sud_est = True Then
            region_click = "SUD EST"
            Call procedure_ajout_button(TextBox4.Value, region_click)
        End If
        If check_sud_ouest = True Then
            region_click = "SUD OUEST"
            Call procedure_ajout_button(TextBox4.Value, region_click)
        End If


                
        Unload UserForm3
        UserForm3.Show
        
        
    End If


End Sub



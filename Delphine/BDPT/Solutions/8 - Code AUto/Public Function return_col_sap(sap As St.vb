Public Function return_col_sap(sap As String) As Long
Dim i_col_sap
For i_col_sap = 1 To nbre_col_ratt
    If Workbooks(base_data).Sheets(onglet_rattachement).Cells(1, i_col_sap).Value = sap Then
        return_col_sap = i_col_sap
    End If
Next i_col_sap

End Function


Public Function return_ligne_ratt(code As Long) As Long

    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_rattachement).Range(lettre_col_codes_ratt & ":" & lettre_col_codes_ratt), code) > 0 Then
        Dim num_ligne_ratt As Long
        num_ligne_ratt = Workbooks(base_data).Sheets(onglet_rattachement).Range(lettre_col_codes_ratt & ":" & lettre_col_codes_ratt).Find(What:=code).Row
        return_ligne_ratt = num_ligne_ratt
    End If

End Function

Public Function return_ligne_art_depot(code As Long) As Long

    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_art_depot).Range(lettre_col_codes_art_depot & ":" & lettre_col_codes_art_depot), code) > 0 Then
        Dim num_ligne_art_depot As Long
        num_ligne_art_depot = Workbooks(base_data).Sheets(onglet_art_depot).Range(lettre_col_codes_art_depot & ":" & lettre_col_codes_art_depot).Find(What:=code).Row
        return_ligne_art_depot = num_ligne_art_depot
    End If

End Function

Sub procedure_compil_final(region As String, fichier_compil_merch As String, fichier_compil_depot As String, fichier_suppr As String)

    If Application.WorksheetFunction.CountIf(Workbooks(workbook_compil).Sheets(onglet_compil).Range("B:B"), region) > 0 Then
        Dim min_ligne_region As Long
        min_ligne_region = Workbooks(workbook_compil).Sheets(onglet_compil).Range("B:B").Find(What:=region).Row
        Dim last_ligne_region As Long
        last_ligne_region = min_ligne_region
        While Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_ligne_region, col_region) = region
            last_ligne_region = last_ligne_region + 1
        Wend
        last_ligne_region = last_ligne_region - 1
    Else
        MsgBox "Aucune données renseignées pour la région " & region
        Exit Sub
    End If
    
    Dim ligne_suiv As Long
    Dim ligne_suiv_2 As Long
    Dim ligne_suiv_suppr As Long
    

    Dim i_code_region As Long
    For i_code_region = 2 To ligne_region
        If CStr(Workbooks(base_data).Sheets(onglet_code_region).Cells(i_code_region, 1).Value) = region Then
            Dim var_sap As String
            var_sap = Workbooks(base_data).Sheets(onglet_code_region).Cells(i_code_region, 2).Value
            
            Dim var_sap_lib As String
            var_sap_lib = Workbooks(base_data).Sheets(onglet_code_region).Cells(i_code_region, 4).Value

            Dim var_easier As String
            var_easier = WorkBooks(base_data).Sheets(onglet_code_region).Cells(i_code_region, 3).Value 'Code EASIER (rajouté)
                     
            Dim i_region_compil As Long
            For i_region_compil = min_ligne_region To last_ligne_region
            
            
                If check_typologie(i_region_compil, i_code_region) = True Then
                    'If return_ligne_art_depot(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_codes)) <> 0 Then
                    If return_ligne_ratt(WorkBooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_codes)) <> 0 Then
                        'If Workbooks(base_data).Sheets(onglet_art_depot).Cells(return_ligne_art_depot(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_codes)), return_col_sap(var_sap)).Value <> "KO" Then
                        If WorkBooks(base_data).Sheets(onglet_rattachement).Cells(return_ligne_ratt(WorkBooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_codes)), return_col_sap(var_sap)).Value = "1" Then    
                            ligne_suiv = Workbooks(fichier_compil_merch).Sheets(1).Cells(Application.Rows.Count, 1).End(xlUp).Row + 1
                            Workbooks(fichier_compil_merch).Sheets(1).Cells(ligne_suiv, 1) = var_easier
                            Workbooks(fichier_compil_merch).Sheets(1).Cells(ligne_suiv, 2) = Workbooks(base_data).Sheets(onglet_code_region).Cells(i_code_region, 4).Value
                            Workbooks(fichier_compil_merch).Sheets(1).Cells(ligne_suiv, 3) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_easier_conc).Value
                            Workbooks(fichier_compil_merch).Sheets(1).Cells(ligne_suiv, 4) = 6
                            Workbooks(fichier_compil_merch).Sheets(1).Cells(ligne_suiv, 5) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value
                            Workbooks(fichier_compil_merch).Sheets(1).Cells(ligne_suiv, 6) = 1
                            
                            ligne_suiv_2 = Workbooks(fichier_compil_depot).Sheets(1).Cells(Application.Rows.Count, 1).End(xlUp).Row + 1
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 1) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_mois).Value
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 2) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_region).Value
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 3) = var_sap_lib
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 4) = "R" & Left(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_codes).Value, 1) & "0"
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 5) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_codes).Value
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 6) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_easier).Value
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 7) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_ean).Value
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 8) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_lib).Value
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 9) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 10) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value)
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 11) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 2
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 12) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 3
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 13) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 4
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 14) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 5
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 15) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 6
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 16) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 7
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 17) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 8
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 18) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 9
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 19) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 10
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 20) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 11
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 21) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, newrange).Value
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 22) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_marche).Value
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 23) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_cat).Value
                            Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 24) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_souscat).Value
                            
                        Else
                            ligne_suiv_suppr = Workbooks(fichier_suppr).Sheets(1).Cells(Application.Rows.Count, 1).End(xlUp).Row + 1
                            Workbooks(fichier_suppr).Sheets(1).Cells(ligne_suiv_suppr, 1) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_mois).Value
                            Workbooks(fichier_suppr).Sheets(1).Cells(ligne_suiv_suppr, 2) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_region).Value
                            Workbooks(fichier_suppr).Sheets(1).Cells(ligne_suiv_suppr, 3) = var_sap
                            Workbooks(fichier_suppr).Sheets(1).Cells(ligne_suiv_suppr, 4) = var_sap_lib
                            Workbooks(fichier_suppr).Sheets(1).Cells(ligne_suiv_suppr, 5) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_codes).Value
                            Workbooks(fichier_suppr).Sheets(1).Cells(ligne_suiv_suppr, 6) = "supprimé"
                        End If
                    Else
                        ligne_suiv = Workbooks(fichier_compil_merch).Sheets(1).Cells(Application.Rows.Count, 1).End(xlUp).Row + 1
                        Workbooks(fichier_compil_merch).Sheets(1).Cells(ligne_suiv, 1) = var_easier
                        Workbooks(fichier_compil_merch).Sheets(1).Cells(ligne_suiv, 2) = Workbooks(base_data).Sheets(onglet_code_region).Cells(i_code_region, 4).Value
                        Workbooks(fichier_compil_merch).Sheets(1).Cells(ligne_suiv, 3) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_easier_conc).Value
                        Workbooks(fichier_compil_merch).Sheets(1).Cells(ligne_suiv, 4) = 6
                        Workbooks(fichier_compil_merch).Sheets(1).Cells(ligne_suiv, 5) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value
                        Workbooks(fichier_compil_merch).Sheets(1).Cells(ligne_suiv, 6) = 1
                        
                        ligne_suiv_2 = Workbooks(fichier_compil_depot).Sheets(1).Cells(Application.Rows.Count, 1).End(xlUp).Row + 1
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 1) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_mois).Value
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 2) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_region).Value
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 3) = var_sap_lib
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 4) = "R" & Left(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_codes).Value, 1) & "0"
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 5) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_codes).Value
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 6) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_easier).Value
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 7) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_ean).Value
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 8) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_lib).Value
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 9) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 10) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value)
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 11) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 2
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 12) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 3
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 13) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 4
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 14) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 5
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 15) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 6
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 16) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 7
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 17) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 8
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 18) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 9
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 19) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 10
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 20) = Round(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value) * 11
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 21) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, newrange).Value
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 22) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_marche).Value
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 23) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_cat).Value
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv_2, 24) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_souscat).Value
                    End If
                        
                Else
                
                    ligne_suiv_suppr = Workbooks(fichier_suppr).Sheets(1).Cells(Application.Rows.Count, 1).End(xlUp).Row + 1
                    Workbooks(fichier_suppr).Sheets(1).Cells(ligne_suiv_suppr, 1) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_mois).Value
                    Workbooks(fichier_suppr).Sheets(1).Cells(ligne_suiv_suppr, 2) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_region).Value
                    Workbooks(fichier_suppr).Sheets(1).Cells(ligne_suiv_suppr, 3) = var_sap
                    Workbooks(fichier_suppr).Sheets(1).Cells(ligne_suiv_suppr, 4) = var_sap_lib
                    Workbooks(fichier_suppr).Sheets(1).Cells(ligne_suiv_suppr, 5) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_codes).Value
                    Workbooks(fichier_suppr).Sheets(1).Cells(ligne_suiv_suppr, 6) = "supprimé"
                    
                End If
            


            Next i_region_compil
            
            

        End If
       
    Next i_code_region

End Sub

'ancienne procedure check typo + onglet rattachement par rapport à article
Sub old_procedure_compil_final(region As String, fichier_compil_depot As String)

    If Application.WorksheetFunction.CountIf(Workbooks(workbook_compil).Sheets(onglet_compil).Range("B:B"), region) > 0 Then
        Dim min_ligne_region As Long
        min_ligne_region = Workbooks(workbook_compil).Sheets(onglet_compil).Range("B:B").Find(What:=region).Row
        Dim last_ligne_region As Long
        last_ligne_region = min_ligne_region
        While Workbooks(workbook_compil).Sheets(onglet_compil).Cells(last_ligne_region, col_region) = region
            last_ligne_region = last_ligne_region + 1
        Wend
        last_ligne_region = last_ligne_region - 1
    Else
        MsgBox "Aucune données renseignées pour la région " & region
        Exit Sub
    End If


    Dim i_code_region As Long
    For i_code_region = 2 To ligne_region
        If CStr(Workbooks(base_data).Sheets(onglet_code_region).Cells(i_code_region, 1).Value) = region Then
            Dim var_sap As String
            var_sap = Workbooks(base_data).Sheets(onglet_code_region).Cells(i_code_region, 2).Value



            Dim i_region_compil As Long
            For i_region_compil = min_ligne_region To last_ligne_region
                If return_ligne_ratt(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_codes)) <> 0 Then


                    If Workbooks(base_data).Sheets(onglet_rattachement).Cells(return_ligne_ratt(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_codes)), return_col_sap(var_sap)).Value = 1 And check_typologie(i_region_compil, i_code_region) = True Then


                        Dim ligne_suiv As Long
                        ligne_suiv = Workbooks(fichier_compil_depot).Sheets(1).Cells(Application.Rows.Count, 1).End(xlUp).Row + 1
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv, 1) = var_sap
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv, 2) = Workbooks(base_data).Sheets(onglet_code_region).Cells(i_code_region, 4).Value
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv, 3) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_easier_conc).Value
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv, 4) = 6
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv, 5) = Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_minfac).Value
                        Workbooks(fichier_compil_depot).Sheets(1).Cells(ligne_suiv, 6) = 1

                    End If
                End If
            Next i_region_compil



       End If

    Next i_code_region

End Sub
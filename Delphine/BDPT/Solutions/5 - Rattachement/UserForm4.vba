Option Explicit


Public workbook_compil As String
Public onglet_compil As String
Public prefix As String
Public annee As String

'Public emplacement_fichier_suivi As String
'Public fichier_suivi As String

Public base_data As String
Public onglet_art_depot As String
Public onglet_rattachement As String
Public onglet_code_region As String


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
Public ligne_art_depot As Long
Public nbre_col_ratt As Long
Public nbre_col_art_depot As Long

Public ligne_suivi As Long
Public flag_remplacement As Integer
Public ligne_remplacement As Long
Public col_codes_ratt As Long
Public col_codes_art_depot As Long

Public lettre_col_codes As String
Public lettre_col_codes_ratt As String
Public lettre_col_codes_art_depot As String

Public last_row_ajout As Long

Public chemin_depot As String






Private Sub UserForm_Initialize()




workbook_compil = ActiveWorkbook.Name
onglet_compil = "COMPIL"
prefix = Left(Replace(Replace(workbook_compil, "COMPILATION_", ""), ".xlsm", ""), 2) & "_"
annee = CStr(Right(Replace(workbook_compil, ".xlsm", ""), 4))

chemin_depot = "https://kfplc.sharepoint.com/teams/OGRP-Marketingdesventes/Shared Documents/Animation commerciale de gamme/" & annee & "/" & prefix & "animco_F&R/3. MDV/"


base_data = "Base data.xlsx"

onglet_art_depot = "couple art dépôt"
onglet_rattachement = "rattachement"
onglet_code_region = "code region"


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
lettre_col_codes_ratt = Split(Workbooks(base_data).Sheets(onglet_rattachement).Cells(1, col_codes_ratt).Address, "$")(1)


'couple_art_depot
ligne_art_depot = Workbooks(base_data).Sheets(onglet_art_depot).Cells(Application.Rows.Count, 1).End(xlUp).Row
nbre_col_art_depot = Workbooks(base_data).Sheets(onglet_art_depot).Range("A1").End(xlToRight).Column
Dim k_art As Long
For k_art = 1 To nbre_col_art_depot
    If LCase(Workbooks(base_data).Sheets(onglet_art_depot).Cells(1, k_art).Value) = "article" Then
        col_codes_art_depot = k_art
    End If
Next k_art
lettre_col_codes_art_depot = Split(Workbooks(base_data).Sheets(onglet_art_depot).Cells(1, col_codes_art_depot).Address, "$")(1)



End Sub

'ancienne fonction check ligne et colonne rattachement
Public Function old_return_col_sap(sap As String) As Long
Dim i_col_sap
For i_col_sap = 1 To nbre_col_ratt
    If Workbooks(base_data).Sheets(onglet_rattachement).Cells(1, i_col_sap).Value = sap Then
        return_col_sap = i_col_sap
    End If
Next i_col_sap

End Function


Public Function old_return_ligne_ratt(code As Long) As Long

    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_rattachement).Range(lettre_col_codes_ratt & ":" & lettre_col_codes_ratt), code) > 0 Then
        Dim num_ligne_ratt As Long
        num_ligne_ratt = Workbooks(base_data).Sheets(onglet_rattachement).Range(lettre_col_codes_ratt & ":" & lettre_col_codes_ratt).Find(What:=code).Row
        return_ligne_ratt = num_ligne_ratt
    End If

End Function

Public Function return_col_sap(sap As String) As Long
Dim i_col_sap
For i_col_sap = 1 To nbre_col_art_depot
    If Workbooks(base_data).Sheets(onglet_art_depot).Cells(1, i_col_sap).Value = sap Then
        return_col_sap = i_col_sap
    End If
Next i_col_sap

End Function


Public Function return_ligne_art_depot(code As Long) As Long

    If Application.WorksheetFunction.CountIf(Workbooks(base_data).Sheets(onglet_art_depot).Range(lettre_col_codes_art_depot & ":" & lettre_col_codes_art_depot), code) > 0 Then
        Dim num_ligne_art_depot As Long
        num_ligne_art_depot = Workbooks(base_data).Sheets(onglet_art_depot).Range(lettre_col_codes_art_depot & ":" & lettre_col_codes_art_depot).Find(What:=code).Row
        return_ligne_art_depot = num_ligne_art_depot
    End If

End Function

Public Function check_typologie(ligne_compil As Long, ligne_sap As Long) As Boolean

Dim var_typo_compil As String
var_typo_compil = Right(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_compil, col_typo).Value, 1)

Dim var_typo_sap As String
var_typo_sap = Right(Workbooks(base_data).Sheets(onglet_code_region).Cells(ligne_sap, 6).Value, 1)


If var_typo_compil = "1" Then
    check_typologie = True
ElseIf var_typo_compil = "2" And (var_typo_sap = "2" Or var_typo_sap = "3") Then
    check_typologie = True
ElseIf var_typo_compil = "3" And var_typo_sap = "3" Then
    check_typologie = True
Else
    check_typologie = False
End If


End Function

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
                    If return_ligne_art_depot(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_codes)) <> 0 Then
                        If Workbooks(base_data).Sheets(onglet_art_depot).Cells(return_ligne_art_depot(Workbooks(workbook_compil).Sheets(onglet_compil).Cells(i_region_compil, col_codes)), return_col_sap(var_sap)).Value <> "KO" Then
                            
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


Private Sub compil_final_button_Click()
            Workbooks.Add
            ActiveCell.FormulaR1C1 = "Numero_depot"
            Range("B1").Select
            ActiveCell.FormulaR1C1 = "Nom_depot"
            Range("C1").Select
            ActiveCell.FormulaR1C1 = "Ref_easier_Lettre"
            Range("D1").Select
            ActiveCell.FormulaR1C1 = "nombre facing merch"
            Range("E1").Select
            ActiveCell.FormulaR1C1 = "min facing"
            Range("F1").Select
            ActiveCell.FormulaR1C1 = "NSPL"
            Range("A1").Select
            Dim nom_fichier_merch As String
            nom_fichier_merch = "Compil_merch" & ".xlsx"
            ActiveWorkbook.SaveAs Filename:=chemin_depot & nom_fichier_merch _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                
                
            Workbooks.Add
            ActiveCell.FormulaR1C1 = "Mois"
            Range("B1").Select
            ActiveCell.FormulaR1C1 = "Region"
            Range("C1").Select
            ActiveCell.FormulaR1C1 = "Nom du dépôt"
            Range("D1").Select
            ActiveCell.FormulaR1C1 = "Rayon"
            Range("E1").Select
            ActiveCell.FormulaR1C1 = "Code"
            Range("F1").Select
            ActiveCell.FormulaR1C1 = "Eaiser"
            Range("G1").Select
            ActiveCell.FormulaR1C1 = "EAN"
            Columns("G:G").Select
            Selection.NumberFormat = "0.00"
            Selection.NumberFormat = "0.0"
            Selection.NumberFormat = "0"
            Range("H1").Select
            ActiveCell.FormulaR1C1 = "Libelle"
            Range("I1").Select
            ActiveCell.FormulaR1C1 = "min facing"
            Range("J1").Select
            ActiveCell.FormulaR1C1 = "facing 1"
            Range("K1").Select
            ActiveCell.FormulaR1C1 = "facing 2"
            Range("L1").Select
            ActiveCell.FormulaR1C1 = "facing 3"
            Range("M1").Select
            ActiveCell.FormulaR1C1 = "facing 4"
            Range("N1").Select
            ActiveCell.FormulaR1C1 = "facing 5"
            Range("O1").Select
            ActiveCell.FormulaR1C1 = "facing 6"
            Range("P1").Select
            ActiveCell.FormulaR1C1 = "facing 7"
            Range("Q1").Select
            ActiveCell.FormulaR1C1 = "facing 8"
            Range("R1").Select
            ActiveCell.FormulaR1C1 = "facing 9"
            Range("S1").Select
            ActiveCell.FormulaR1C1 = "facing 10"
            Range("T1").Select
            ActiveCell.FormulaR1C1 = "facing 11"
            Range("U1").Select
            ActiveCell.FormulaR1C1 = "flag code siege"
            Range("V1").Select
            ActiveCell.FormulaR1C1 = "Marché"
            Range("W1").Select
            ActiveCell.FormulaR1C1 = "Catégorie"
            Range("X1").Select
            ActiveCell.FormulaR1C1 = "Sous-Catégorie"
            Range("A1").Select
            Dim nom_fichier_depot As String
            nom_fichier_depot = "Compil_depot" & ".xlsx"
            ActiveWorkbook.SaveAs Filename:=chemin_depot & nom_fichier_depot _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            
        
                
            Workbooks.Add
            ActiveCell.FormulaR1C1 = "Mois"
            Range("B1").Select
            ActiveCell.FormulaR1C1 = "Région"
            Range("C1").Select
            ActiveCell.FormulaR1C1 = "Numéro de dépôt"
            Range("D1").Select
            ActiveCell.FormulaR1C1 = "Nom du dépôt"
            Range("E1").Select
            ActiveCell.FormulaR1C1 = "Code"
            Range("F1").Select
            ActiveCell.FormulaR1C1 = "Suppression"
            Range("A1").Select
            Dim nom_fichier_suppr As String
            nom_fichier_suppr = "Fichier_suivi_art_depot" & ".xlsx"
            ActiveWorkbook.SaveAs Filename:=chemin_depot & nom_fichier_suppr _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False





Dim region_encours As String

region_encours = "CENTRE AQUITAINE"
Call procedure_compil_final(region_encours, nom_fichier_merch, nom_fichier_depot, nom_fichier_suppr)

region_encours = "CENTRE ATLANTIQUE"
Call procedure_compil_final(region_encours, nom_fichier_merch, nom_fichier_depot, nom_fichier_suppr)

region_encours = "CENTRE EST"
Call procedure_compil_final(region_encours, nom_fichier_merch, nom_fichier_depot, nom_fichier_suppr)

region_encours = "REGION EST"
Call procedure_compil_final(region_encours, nom_fichier_merch, nom_fichier_depot, nom_fichier_suppr)

region_encours = "IDF"
Call procedure_compil_final(region_encours, nom_fichier_merch, nom_fichier_depot, nom_fichier_suppr)

region_encours = "NORD EST"
Call procedure_compil_final(region_encours, nom_fichier_merch, nom_fichier_depot, nom_fichier_suppr)

region_encours = "NORD"
Call procedure_compil_final(region_encours, nom_fichier_merch, nom_fichier_depot, nom_fichier_suppr)

region_encours = "NORMANDIE"
Call procedure_compil_final(region_encours, nom_fichier_merch, nom_fichier_depot, nom_fichier_suppr)

region_encours = "OUEST"
Call procedure_compil_final(region_encours, nom_fichier_merch, nom_fichier_depot, nom_fichier_suppr)

region_encours = "RHONE ALPES"
Call procedure_compil_final(region_encours, nom_fichier_merch, nom_fichier_depot, nom_fichier_suppr)

region_encours = "SUD EST"
Call procedure_compil_final(region_encours, nom_fichier_merch, nom_fichier_depot, nom_fichier_suppr)

region_encours = "SUD OUEST"
Call procedure_compil_final(region_encours, nom_fichier_merch, nom_fichier_depot, nom_fichier_suppr)


Windows(nom_fichier_merch).Activate
ActiveWorkbook.Save

Windows(nom_fichier_depot).Activate
ActiveWorkbook.Save

Windows(nom_fichier_suppr).Activate
ActiveWorkbook.Save

MsgBox "Compilation final réalisée avec succès ! " & Chr(13) & Chr(10) & "BRAVO !"

Unload UserForm4



End Sub

Private Sub fermer_Click()
Unload UserForm4
End Sub

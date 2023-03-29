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
                If ListBox1.Selected(j) = True Then 'si c'est cet element qu'on a selectionné donc
                
                        Dim ligne_modif As Long
                        ligne_modif = ListBox1.List(j, 0) 'modifier la recup ici du coup
                        
                        'cas code flag siege
                        If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, newrange) = 1 Then
                            Dim choix

                            choix = MsgBox("Le code " & TextBox1.Value & " a été ajouté par le siège à la ligne " & ligne_modif & " et pour région " & Cells(ligne_modif, col_region) & " . Souhaitez vous le remplacer?", 36, "Confirmation")
                            If choix = vbNo Then
                                Exit Sub
                            End If
                        End If
                        
                        'si la police est pas comme l'originale (code digf)
                        If Workbooks(workbook_compil).Sheets(onglet_compil).Cells(ligne_modif, col_codes).Font.ColorIndex <> xlAutomatic Then
                            
                            'COMPOSANT, si c'est un composant donc
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
                                
                                      'MsgBox ligne_modif_cas_comp
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

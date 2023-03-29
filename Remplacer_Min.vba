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
                If ListBox1.Selected(j) = True Then 'si c'est cet element qu'on a selectionné donc
                
                        Dim ligne_modif As Long
                        ligne_modif = ListBox1.List(j, 0)
                        pos = InStr(ligne_modif, " ")
                        ligne_modif = Left(ligne_modif, pos - 1)
                        ligne_modif = CInt(ligne_modif)
                        'Numéro de ligne récupéré


                        'A partir d'ici faire le remplacement des mins.
                        
                        
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

' VBA pour envoyer des emails automatiquement depuis Excel avec Outlook
Sub SendEmailsFromExcel()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim lastRow As Long, i As Integer
    
    ' Définir la feuille contenant les emails
    Set ws = ThisWorkbook.Sheets("Emails")
    
    ' Trouver la dernière ligne remplie
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Initialiser Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' Boucle sur chaque ligne pour envoyer un email
    For i = 2 To lastRow
        ' Créer un nouvel email
        Set OutlookMail = OutlookApp.CreateItem(0)
        
        With OutlookMail
            .To = ws.Cells(i, 1).Value  ' Adresse email
            .Subject = "[Alerte Stock] Mise à jour de l’inventaire"
            .Body = ws.Cells(i, 2).Value  ' Corps du message
            
            ' Ajouter une pièce jointe si précisée
            If ws.Cells(i, 3).Value <> "" Then
                .Attachments.Add ws.Cells(i, 3).Value
            End If
            
            .Send ' Envoyer l'email
        End With
        
        ' Nettoyage
        Set OutlookMail = Nothing
    Next i
    
    ' Nettoyer l'objet Outlook
    Set OutlookApp = Nothing
    
    MsgBox "Tous les emails ont été envoyés avec succès !", vbInformation, "Envoi Terminé"
End Sub

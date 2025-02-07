 PrivatPrivate sub Form_Load()
'Carica l'elenco dei caffè
Dim strCaffè as string
Dim strPercorsoFile 
On Error GoTo GestioneErrori
strPercorsoFile = App.Path & "\Caffe.dat"
Open strPercorsoFile For Input As #1
Do Until EOF()
   Input #1, strCaffè
   cmbCaffe.addItem strCaffè
loop
Close #1

Form_Load_Exit:
    Exit sub

GestioneErrori:
    dim intRisposta as Integer

    Select case Err.Number
           Case 53, 76
           intRisposta = MsgBox("Creare un nuovo file?", _ vbYesNo + vbQuestion
           , "File non trovato")

           If intRisposta = vbYes Then 
              Resume Form_Load_Exit
           else
              mnuFile_Esci_Click
           End If
           Case 71
              intRisposta = MsgBox("Disco non pronto. Riprovare?", vbRetryCancel + vbQuestion, "Errore di disco")
              if intRisposta = vbRetry Then
                 Resume
              else
                 mnuFile_Esci_Click
              end if
           Case else
              Err.Raise Err
    End Select
End sub
e sub Form_Load()
'Carica l'elenco dei caffè
Dim strCaffè as string
Dim strPercorsoFile 
On Error GoTo GestioneErrori
strPercorsoFile = App.Path & "\Caffe.dat"
Open strPercorsoFile For Input As #1
Do Until EOF()
   Input #1, strCaffè
   cmbCaffe.addItem strCaffè
loop
Close #1

Form_Load_Exit:
    Exit sub

GestioneErrori:
    dim intRisposta as Integer

    Select case Err.Number
           Case 53, 76
           intRisposta = MsgBox("Creare un nuovo file?", _ vbYesNo + vbQuestion
           , "File non trovato")

           If intRisposta = vbYes Then 
              Resume Form_Load_Exit
           else
              mnuFile_Esci_Click
           End If
           Case 71
              intRisposta = MsgBox("Disco non pronto. Riprovare?", vbRetryCancel + vbQuestion, "Errore di disco")
              if intRisposta = vbRetry Then
                 Resume
              else
                 mnuFile_Esci_Click
              end if
           Case else
              Err.Raise Err
    End Select
End sub
Private Sub Form_QueryUnload(Cancel as Integer, UnloadMode as Integer)  
'Chiede all'utente di salvare il file
Dim intRisposta as Integer
   If mblnNuoveModifiche = True Then
      intRisposta = MsgBox("Lista del caffè modificata. Va salvata?", vbYesNo + vbQuestion, "Lista del caffè modificata")
         if intRisposta = vbYes Then
         mnuFileSalva_Click
      End If
   End If
End Sub

Private Sub mnuModificaAzzera_Click()
    'Azzera l'elenco dei caffè
    Dim intRisposta as Integer
        intRisposta = MsgBox("Azzerare l'elenco dei tipi di caffè?", vbYesNo + vbQuestion, "Azzera elenco caffè")
        if intRisposta = vbYes Then
           cmbCaffe.Clear
           mblnNuoveModifiche = True
        End if
End Sub
Private Sub mnuModificaRimuovi_Click()
    'Rimuove il caffè selezionato dall'elenco
    Dim strMsg as String 
        If cmbCaffe.ListIndex <> -1 Then
           cmbCaffe.RemoveItem cmbCaffe.ListIndex 
           mblnNuoveModifiche = True
        else
           strMsg = "Selezionare prima il caffè da rimuovere."
           MsgBox strMsg, vbInformation, "Nessuna selezione eseguita"
        End if
End Sub
Private Sub mnuFileSalva_Click()
'Salva il contenuto della casella di riepilogo su un file sequenziale
Dim intIndice as Integer
Dim intMassimo as Integer
Dim strPercorsoFile as String

strPercorsoFile = App.Path & "\Caffe.dat"
Open strPercorsoFile for Output As #1
intMassimo = cboCaffe.ListCount - 1
for intIndice = 0 to intMassimo
    Write #1, cboCaffe.List(intIndice)
Next intIndice
Close #1
mblnNuoveModifiche = False
End Sub

Private Sub mnuFileEsci_Click()
'Termina il progetto
   Unload frmAbout
   Unload Me 
   End
End Sub

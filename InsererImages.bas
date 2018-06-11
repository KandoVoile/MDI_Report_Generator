Attribute VB_Name = "InsererImages"
Sub Initialisation_des_données()
'Menu directed Inspection GE Measurement & control
'Olivier Talbot 02-2015
'importation d'une inspection réaliser sur XLG3 ou Mentor dans la trame du client Air France Industries
'déclaration des variables

Dim Repertoire As String
Dim Extension As String
Dim Fichier As String
Dim selDossier(1) As String 'déclaration de la variable de selection de dossier MDI
Dim Variable(1) As String
Dim DossierMdi As String
Dim Message(1) ' déclarationde la variable pour demandé si l'inspection à été faite avec deux appareils
Dim nom As String
Dim emplac(27) As Variant
Dim serie_moteur(1) As String
Dim image(25) As String
Dim FoundSatisfactory(25) As String
Dim recherche As Variant
Dim NomSansExtenxion As String
Dim Cheminxml As String
Dim rep As String
Dim nouveau As String
Dim strChemin1 As String
Dim Tag(3) As String
Dim ESN As String
Dim nouveau_nom As String
Dim test As String
Dim arborescence As String
Dim nomdocument As String
Dim fd As FileDialog




'Dim ObjXml, Buffer

arborescence = "*.inspection"

'emplacement d'insertion des images dans le rapport

'emplac(0) = "IDPlate"
emplac(0) = "lpcstg2"
emplac(1) = "lpcstg3"
emplac(2) = "lpcstg4"
emplac(3) = "lpcstg5"
emplac(4) = "hpcstg1"
emplac(5) = "hpcstg2"
emplac(6) = "hpcstg3"
emplac(7) = "hpcstg4"
emplac(8) = "hpcstg5"
emplac(9) = "hpcstg6"
emplac(10) = "hpcstg7"
emplac(11) = "hpcstg8"
emplac(12) = "hpcstg9"
emplac(13) = "cc"
emplac(14) = "hptngv1"
emplac(15) = "hptstg1"
emplac(16) = "lptngv1"
emplac(17) = "lptstg1"
emplac(18) = "lptngv2"
emplac(19) = "lptstg2"
emplac(20) = "lptngv3"
emplac(21) = "lptstg3"
emplac(22) = "lptngv4"
emplac(23) = "lptstg4"
emplac(24) = "lptngv5"
emplac(25) = "lptstg5"
emplac(26) = "customer"
emplac(27) = "location"
serie_moteur(1) = "serie_moteur"

'emplacement commentaire si pas d'image

FoundSatisfactory(0) = "lpcstg2_C"
FoundSatisfactory(1) = "lpcstg3_C"
FoundSatisfactory(2) = "lpcstg4_C"
FoundSatisfactory(3) = "lpcstg5_C"
FoundSatisfactory(4) = "hpcstg1_C"
FoundSatisfactory(5) = "hpcstg2_C"
FoundSatisfactory(6) = "hpcstg3_C"
FoundSatisfactory(7) = "hpcstg4_C"
FoundSatisfactory(8) = "hpcstg5_C"
FoundSatisfactory(9) = "hpcstg6_C"
FoundSatisfactory(10) = "hpcstg7_C"
FoundSatisfactory(11) = "hpcstg8_C"
FoundSatisfactory(12) = "hpcstg9_C"
FoundSatisfactory(13) = "CC_C"
FoundSatisfactory(14) = "hptngv1_C"
FoundSatisfactory(15) = "hptstg1_C"
FoundSatisfactory(16) = "lptngv1_C"
FoundSatisfactory(17) = "lptstg1_C"
FoundSatisfactory(18) = "lptngv2_C"
FoundSatisfactory(19) = "lptstg2_C"
FoundSatisfactory(20) = "lptngv3_C"
FoundSatisfactory(21) = "lptstg3_C"
FoundSatisfactory(22) = "lptngv4_C"
FoundSatisfactory(23) = "lptstg4_C"
FoundSatisfactory(24) = "lptngv5_C"
FoundSatisfactory(25) = "lptstg5_C"



'nom des photos générées par MDI
image(0) = "*_LPC_stg2_*.jpg"
image(1) = "*_LPC_stg3_*.jpg"
image(2) = "*_LPC_stg4_*.jpg"
image(3) = "*_LPC_stg5_*.jpg"
image(4) = "*_HPC_stg1_*.jpg"
image(5) = "*_HPC_stg2_*.jpg"
image(6) = "*_HPC_stg3_*.jpg"
image(7) = "*_HPC_stg4_*.jpg"
image(8) = "*_HPC_stg5_*.jpg"
image(9) = "*_HPC_stg6_*.jpg"
image(10) = "*_HPC_stg7_*.jpg"
image(11) = "*_HPC_stg8_*.jpg"
image(12) = "*_HPC_stg9_*.jpg"
image(13) = "*_CC_*.jpg"
image(14) = "*_HPT_NGV_1_*.jpg"
image(15) = "*_HPT_stg1_*.jpg"
image(16) = "*_LPT_NGV_1_*.jpg"
image(17) = "*_LPT_stg1_*.jpg"
image(18) = "*_LPT_NGV_2_*.jpg"
image(19) = "*_LPT_stg2_*.jpg"
image(20) = "*_LPT_stg3_*.jpg"
image(21) = "*_LPT_NGV_3_*.jpg"
image(22) = "*_LPT_stg4_*.jpg"
image(23) = "*_LPT_NGV_4_*.jpg"
image(24) = "*_LPT_stg5_*.jpg"
image(25) = "*_LPT_NGV_5_*.jpg"


Message(0) = "Selectionnez le dossier de la première partie de l'inspection"
Message(1) = "Selectionnez le dossier de la seconde partie de l'inspection"

Tag(0) = "TAG='0015:2065'"
Tag(1) = "TAG='0013:4019'"
Tag(2) = "TAG='0040:A30A'"

Dim NomWtihSn As String




'----------------------------------------------------------------------------------------------------------------------------------------------------------
'nouveau = InputBox("sous quel nom voulez-vous enregistrer votre rapport?")
'If nouveau = "" Then wdapp.Quit
'activefile.saveas filename:=
'demande à l'utilisateur de selectionner le dossier dans lequel ce trouve l'inspection MDI
If MsgBox("avez vous réalisez l'inspection sur deux appareils différent?", vbYesNo, "demande de confirmation") = vbYes Then
    For i = 0 To 1

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    MsgBox (Message(i))
        
        Do Until Variable(i) <> ""
            With fd
            .InitialFileName = Defaut
                If .Show = -1 Then
            Variable(i) = fd.SelectedItems(1)
                End If
          
            End With
        Exit Do
        Loop


    Next
    DossierMdi = Merge2MDI(Variable(0), Variable(1))
    
End If
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        Do Until DossierMdi <> ""
           With fd
            .InitialFileName = Defaut
            If .Show = -1 Then
                DossierMdi = fd.SelectedItems(1)
            End If
        If DossierMdi = "" Then
        MsgBox ("Selectionner un dossier")
        i = i + 1
            If i = 2 Then
            MsgBox ("Pas de dossier selectionné le programme est arrêté")
            Exit Do
            
            End If
        
        End If
    
    End With
Loop

Set fd = Nothing
' va chercher le numero de serie du moteur

    If Len(Dir(DossierMdi & "\" & arborescence)) Then
    rep = Dir(DossierMdi & "\" & arborescence)
    ESN = LectureXML(DossierMdi, rep, 0)
      
        test = LectureXML(DossierMdi, rep, 1)
        If test = "" Then
        Else
        nomdocument = (ESN & " - " & test)
        
        ActiveDocument.Save
        ChangeFileOpenDirectory (DossierMdi)
        ActiveDocument.SaveAs Filename:=(nomdocument & ".doc")
        ActiveDocument.Bookmarks("serie_moteur").Select
        Selection.TypeText ESN
        ActiveDocument.ActiveWindow.View = wdPrintView
                
        ActiveDocument.Bookmarks("Esn").Select
        Selection.TypeText nomdocument
        ActiveDocument.ActiveWindow.View = wdPrintView
        
        test = 0
        End If
        
    End If


'Extension = InputBox("Type de fichier (sans le point, ex : jpg, png, bmp, jpeg)", "Type de fichier", "jpg")
' recherche toute les photos correspondant au titre MDI
For i = 0 To 25
    'regarde si le fichier selectionner correspond à celui recherché
    If Len(Dir(DossierMdi & "\" & image(i))) <> 0 Then
            rep = Dir(DossierMdi & "\" & image(i))
            'enregistre le nom du fichier dans une variable pour recherche le xml
            strChemin1 = rep
            'appel la fonction NomSansExtension pour extraire le nom sans .jpeg
            NomSansExtenxion = FilenameWithoutExt(strChemin1)
                                                          
            'appel la fonction lecture xml et extrai les commentaires
                                    
            While Not rep = ""
            'nouveau_nom = Esn & "_" & rep '(change le nom du fichier pour inserer le sn moteur)
            
            Selection.GoTo what:=wdGoToBookmark, Name:=emplac(i)
            Set objshape = Selection.InlineShapes.AddPicture(Filename:=(DossierMdi & "\" & rep))
                        With objshape
                        .LockAspectRatio = msoTrue
                        .Width = 220
                        End With
            NomSansExtenxion = FilenameWithoutExt(rep)
            'Name rep As nouveau_nom '(a activer pour renomer les fichiers)
            
              For u = 2 To 4
                test = LectureXML(DossierMdi, NomSansExtenxion, (u))
                If test = "" Then
                Else
                Selection.GoTo what:=wdGoToBookmark, Name:=FoundSatisfactory(i)
                Selection.TypeText (test)
                End If
            Next

            rep = Dir
            Wend
                      
                        
    Else
    Selection.GoTo what:=wdGoToBookmark, Name:=FoundSatisfactory(i)
    Selection.TypeText ("Found Satisfactory")
                      
  
   End If
Next
'efface le bouton de lancement macro avant enregistrement

    'For Each shp In ActiveDocument.InlineShapes
    '    If shp.OLEFormat.ClassType = "Forms.MDI" Then shp.Delete
   ' Next
    
ActiveDocument.Save
ChangeFileOpenDirectory (DossierMdi)
ActiveDocument.SaveAs Filename:=(nomdocument & ".doc")
MsgBox " insertion des images terminé & fichier sauvegardé"



End Sub



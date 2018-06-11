Attribute VB_Name = "Merge2MdiFiles"
Function Merge2MDI(selDossier0 As Variant, selDossier1 As Variant)

'déclaration des variables
Dim NomFichier As String
Dim Fichier
Dim Fichierxml
Dim hml
Dim Chemin
Dim Chemin2
Dim Cheminxml
Dim Chemin2xml
Dim fd As FileDialog
Dim image



image = "*.jpeg"
' ask the technician to select the 2 MDI files.


Chemin = selDossier1 & "\"
Chemin2 = selDossier0 & "\"
Fichier = Dir(Chemin & "*.jpg")
    
    If selDossier1 <> False Then
                     
            Do While Len(Fichier) > 0
            
            FileCopy Chemin & Fichier, Chemin2 & Fichier
                
            Fichier = Dir
            Loop
      
    End If
Cheminxml = selDossier1 & "\" & "xml" & "\"
Chemin2xml = selDossier0 & "\" & "xml" & "\"
Fichierxml = Dir(Cheminxml & "*.xml")
    If selDossier1 <> False Then
                     
            Do While Len(Fichierxml) > 0
            
            FileCopy Cheminxml & Fichierxml, Chemin2xml & Fichierxml
                
            Fichierxml = Dir
            Loop
      
    End If
    
Merge2MDI = Chemin2

End Function


Attribute VB_Name = "extractionchemin"
Option Explicit

' EXTRACTION D'UN NOM DE FICHIER AVEC SON EXTENSION
' ---
' Entrée : strPath : Chemin d'un fichier
'                    Ex. : test.jpg
'                          C:\un\chemin\quelconque\test.jpg
' Sortie : Nom du fichier avec son extension (ex. : test.jpg).
'
Function Filename(ByVal strPath As String) As String
  ' Trouver le dernier backslash, s'il y en a un...
  Dim intI As Integer
  intI = InStrRev(strPath, "\", -1, vbTextCompare)
  
  ' Renvoyer la partie après le backslash
  Filename = IIf(intI = 0, strPath, Mid(strPath, intI + 1))
End Function

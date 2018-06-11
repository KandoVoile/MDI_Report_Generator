Attribute VB_Name = "Fichiersansextenxion"
Option Explicit

Function FilenameWithoutExt( _
  ByVal strPath As String)
  
  Dim intI As Integer
  
  ' Extraire uniquement le nom de fichier
  ' (au cas où on aurait transmis un chemin complet)
  strPath = Filename(strPath)
  
  intI = InStrRev(strPath, ".", -1, vbTextCompare)
  If intI = 0 Then
    FilenameWithoutExt = strPath
  Else
    FilenameWithoutExt = Left(strPath, intI - 1)
  End If
End Function



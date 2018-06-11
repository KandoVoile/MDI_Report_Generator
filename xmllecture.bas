Attribute VB_Name = "xmllecture"
Option Explicit

Function LectureXML(DossierMdi As String, Chemin As String, recherche As Byte)
Dim typemesure As Integer
Dim mesure As Integer
Dim ObjXml, Buffer
Dim test As String
'test = "E:\CFM56-5B-5555\xml\HPC_High_Pressure_Compressor_HPC_STG_1001.xml"


Set ObjXml = CreateObject("Microsoft.XMLDOM")
ObjXml.Async = False
    If recherche = 0 Then
    ' recherche du champ ESN
    ObjXml.Load (DossierMdi & "\" & Chemin)
    Set Buffer = ObjXml.DocumentElement.SelectSingleNode("//MDI_FIELD[@VALUE='Serial Number']")
    ElseIf recherche = 1 Then
    ' recherche du champ date
    ObjXml.Load (DossierMdi & "\" & Chemin)
    Set Buffer = ObjXml.DocumentElement.SelectSingleNode("//MDI_FIELD[@VALUE='Date']")
    ElseIf recherche = 2 Then
    'recherche du commentaire
    ObjXml.Load (DossierMdi & "\" & "xml" & "\" & Chemin & ".xml")
    Set Buffer = ObjXml.DocumentElement.SelectSingleNode("//DICOM_ATTRIBUTE[@TAG='0015:2065']")
    ElseIf recherche = 3 Then
    'recherche du type de mesure
    ObjXml.Load (DossierMdi & "\" & "xml" & "\" & Chemin & ".xml")
    Set Buffer = ObjXml.DocumentElement.SelectSingleNode("//DICOM_ATTRIBUTE[@TAG='0013:4019']")
    ElseIf recherche = 4 Then
    'recherche de la valeur de la mesure
    ObjXml.Load (DossierMdi & "\" & "xml" & "\" & Chemin & ".xml")
    Set Buffer = ObjXml.DocumentElement.SelectSingleNode("//DICOM_ATTRIBUTE[@TAG='0040:A30A']")
    End If
    
    If Buffer Is Nothing Then
        
        Else
        
        LectureXML = Buffer.Text
    End If
    
        


End Function

  



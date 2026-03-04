Attribute VB_Name = "KMZ_Fission"
Sub Split_KMZ()
    Dim XDoc As MSXML2.DOMDocument60
    Dim Namespace As String
    Dim polygonNodes As MSXML2.IXMLDOMNodeList
    Dim polygonNode As MSXML2.IXMLDOMNode
    Dim polName As String
    Dim coords As Variant
    Dim polData() As Variant
    Dim XPath As String
    Dim PolygonsPath As String
    Dim ToSplitPath As String
    
    Namespace = "xmlns:kml='http://www.opengis.net/kml/2.2'"
    PolygonsPath = ThisWorkbook.Sheets("KMZ Separator").Range("Polygons_KML").Value
    ToSplitPath = ThisWorkbook.Sheets("KMZ Separator").Range("KMZ_ToSplit").Value
    
    'Load Polygons XML
    Set PolygonsXDoc = Load_From_KML_Or_KMZ(PolygonsPath)
    
    'Get Nodes that have a child "Polygon" within the polygon XML
    Set polygonNodes = PolygonsXDoc.SelectNodes("//kml:Polygon/parent::*")
    
    'Load ToSplit XML
    Set ToSplitXDoc = Load_From_KML_Or_KMZ(ToSplitPath)
    
    'Get all relevant nodes within the ToSplit XML
    Set toSplitNodes = ToSplitXDoc.SelectNodes("//kml:Placemark")
    Set remainderNodes = toSplitNodes 'Remainder nodes will track any leftovers that weren't within any polygon
    
    ' If ToSplitPath Like "*.kml" Then
        FileType = ".kml"
    ' ElseIf ToSplitPath Like "*.kmz" Then
    '     FileType = ".kmz"
    ' End If
    
    ' Prompt the user for a filepath for the KMZ output
    FileName = "Separated KML"
    suggestedName = FileName & FileType 'TODO: Do some logic to come up with a filename recommendation based on original KMZ/KML
    
    ' Ask the user for the save location for the KMZ
    ' If FileType = ".kml" Then
        outputPath = Application.GetSaveAsFilename( _
            fileFilter:="KML Files (*.kml), *.kml", _
            title:="Choose KML output location", _
            InitialFileName:=suggestedName)
    ' ElseIf FileType = ".kmz" Then
    '    outputPath = Application.GetSaveAsFilename( _
            fileFilter:="KMZ Files (*.kmz), *.kmz", _
            title:="Choose KMZ output location", _
            InitialFileName:=suggestedName)
    ' End If
    
    ' If the user clicked Cancel, ask them whether they want to continue the macro
    If outputPath = False Then
        MsgBox ("You didn't select a save location for the separated file." & vbNewLine & "Please try again.")
        End
    End If
    
    fileTitleWithExt = Mid(outputPath, InStrRev(outputPath, "\") + 1)
    FileTitle = Left(fileTitleWithExt, Len(fileTitleWithExt) - 4)
    outputPathNoExt = Left(outputPath, Len(outputPath) - 4)
    
    ' Initialize a new XML to save the separated info
    Dim newFile As MSXML2.DOMDocument60
    Set newFile = InitKML(FileTitle)
        
    ' Get the top level folder of the new XML
    Dim topLevel As IXMLDOMNode
    Set topLevel = newFile.SelectSingleNode("//kml:Folder")
    
    ' Copy all the style info from the original XML
    Set styleNodes = ToSplitXDoc.SelectNodes("//kml:Style")
    For i = 0 To styleNodes.Length - 1
        topLevel.appendChild styleNodes(i)
    Next i
    Set styleMaps = ToSplitXDoc.SelectNodes("//kml:StyleMap")
    For i = 0 To styleMaps.Length - 1
        topLevel.appendChild styleMaps(i)
    Next i
    
    'Iterate through the Nodes/Polygons
    For i = 0 To polygonNodes.Length - 1
        Set polygonNode = polygonNodes(i)
        polName = polygonNode.SelectSingleNode(".//kml:name").Text
        'Extract the coordinates
        coords = Split(polygonNode.SelectSingleNode(".//kml:coordinates").Text)
        ReDim polData(0 To UBound(coords), 1 To 2)
        For j = 0 To UBound(coords)
            polData(j, 1) = CDbl(Split(coords(j), ",")(0))
            polData(j, 2) = CDbl(Split(coords(j), ",")(1))
        Next j
        
        ' Make the XML folder for this polygon
        Dim polFolder As IXMLDOMNode
        Set polFolder = MakeFolder(newFile, polName)
        topLevel.appendChild polFolder
        
        'Loop through each object and separate them into folders(more than one if there's overlap)
        For Each toSplitNode In toSplitNodes
            'Name = toSplitNode.SelectSingleNode(".//kml:name").Text
            NodeCoords = Split(toSplitNode.SelectSingleNode(".//kml:coordinates").Text)
            result = False
            For k = 0 To UBound(NodeCoords)
                NodeLong = CDbl(Split(NodeCoords(k), ",")(0))
                NodeLat = CDbl(Split(NodeCoords(k), ",")(1))
                If PtInPoly(NodeLong, NodeLat, polData) = True Then
                    result = True
                End If
            Next k
            
            If result = True Then
                polFolder.appendChild toSplitNode
                'remainderNodes.RemoveChild toSplitNode
            End If
        Next toSplitNode
    Next i
    
    'If the original ToSplit file was a KMZ, copy the original "files" folder from it, and zip it up with the new file into a KMZ
    Call CombineXMLandKMZ(newFile, ToSplitPath, outputPath, FileTitle)
End Sub


Function PtInPoly(ByVal Xcoord As Double, ByVal Ycoord As Double, ByVal polygon As Variant) As Variant
' Returns a true/false result if the given coordinates are inside the given polygon

  Dim x As Long, NumSidesCrossed As Long, m As Double, b As Double, Poly As Variant
  
  Poly = polygon
  
  ' Raycast from the given point to a known point in the Polygon
  For x = LBound(Poly) To UBound(Poly) - 1
    If Poly(x, 1) > Xcoord Xor Poly(x + 1, 1) > Xcoord Then
      m = (Poly(x + 1, 2) - Poly(x, 2)) / (Poly(x + 1, 1) - Poly(x, 1))
      b = (Poly(x, 2) * Poly(x + 1, 1) - Poly(x, 1) * Poly(x + 1, 2)) / (Poly(x + 1, 1) - Poly(x, 1))
      If m * Xcoord + b > Ycoord Then NumSidesCrossed = NumSidesCrossed + 1
    End If
  Next
  
  ' If the raycast passes the shape border an even number of times, the given point is also inside the Polygon
  PtInPoly = CBool(NumSidesCrossed Mod 2)
  
End Function

Function InitKML(ByVal title As String) As MSXML2.DOMDocument60
    ' Make a new XML document
    Dim kml As MSXML2.DOMDocument60
    Set kml = New MSXML2.DOMDocument60
    
    ' Load in some initial, boilerplate XML, as well as a starting folder
    kml.LoadXML ("<?xml version=""1.0"" encoding=""UTF-8""?>" & _
        "<kml xmlns='http://www.opengis.net/kml/2.2' xmlns:gx='http://www.google.com/kml/ext/2.2' xmlns:kml='http://www.opengis.net/kml/2.2' xmlns:atom='http://www.w3.org/2005/Atom'>" & _
        "<Folder><name>" & title & "</name><open>1</open>" & _
        "<Style id=""dot""><scale>1</scale><Icon><href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href></Icon><LabelStyle><scale>0.85</scale></LabelStyle></Style>" & _
        "</Folder>" & _
        "</kml>")
        
    ' Set document properties and return the document
    kml.SetProperty "SelectionNamespaces", "xmlns='http://www.opengis.net/kml/2.2' xmlns:gx='http://www.google.com/kml/ext/2.2' xmlns:kml='http://www.opengis.net/kml/2.2' xmlns:atom='http://www.w3.org/2005/Atom'"
    kml.SetProperty "SelectionLanguage", "XPath"
        
    Set InitKML = kml
End Function

Function Load_From_KML_Or_KMZ(path As String) As MSXML2.DOMDocument60
    ' Takes a filepath argument, returns an XML object

    ' If a KML path was given, we can proceed without hassle!
    If path Like "*.kml" Then
        kmlPath = path
    ' If a KMZ path was given, the KML will need to be extracted
    ElseIf path Like "*.kmz" Then
        ' Establish the path to the temp folder
        tempPath = Left(path, InStrRev(path, "\"))
        tempFolderName = "KMZ_LOAD_TEMP"
        tempFullPath = tempPath & tempFolderName
    
        ' Create the temp folder
        CreateFolderPath (tempFullPath)
        
        ' Extract the KML into the temp folder
        kmlPath = tempFullPath & "\ripeKML.kml"
        ExtractKML path, kmlPath, tempFullPath
        
        ' Alert the user if the KML file is missing (this happens if the KMZ fails to unzip properly)
        If dir(kmlPath) = "" Then
            MsgBox "KMZ couldn't be extracted! To fix, install 7-Zip, or open the KMZ file in Google Earth and re-save it before using with this tool."
            End
        End If
    Else
        MsgBox "Expected KMZ or KML file!" & vbNewLine & "Path: " & path
        End
    End If
    
    
    ' Declare variables
    Dim xmlObj As MSXML2.DOMDocument60
    Dim Namespace As String
    
    ' Set KML Standard Namespace
    ' Namespace = "xmlns='http://www.opengis.net/kml/2.2' xmlns:gx='http://www.google.com/kml/ext/2.2' xmlns:kml='http://www.opengis.net/kml/2.2' xmlns:atom='http://www.w3.org/2005/Atom'"
    Namespace = "xmlns:kml='http://www.opengis.net/kml/2.2'"
    
    ' Load XML
    Set xmlObj = New MSXML2.DOMDocument60
    Call xmlObj.SetProperty("SelectionNamespaces", Namespace)
    Call xmlObj.SetProperty("SelectionLanguage", "XPath")
    xmlObj.async = False: xmlObj.validateOnParse = False
    xmlObj.Load (kmlPath)
    
    ' Clean up the temp directory if one exists
    If Not IsEmpty(tempFullPath) Then
        Dim fileSystemObject As Object
        Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
        fileSystemObject.DeleteFolder tempFullPath
    End If

    ' Return the XML object
    Set Load_From_KML_Or_KMZ = xmlObj
End Function

Sub ExtractKML(ByVal pathToKMZ As String, ByVal returnPath As String, Optional ByVal temp As String = "")
    Dim applicationObject As Object
    Set applicationObject = CreateObject("Shell.Application")
    Dim fileSystemObject As Object
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    
    ' If a temporary working folder is not provided, create one in the same folder as the KMZ
    tempDirectory = ""
    If (temp = "") Then
        tempDirectory = fileSystemObject.CreateFolder(fileSystemObject.GetParentFolderName(pathToKMZ) & "\KML_CONVERSION_TEMP")
    Else
        tempDirectory = temp
    End If
    
    ' Copy the KMZ to the temp folder as a Zip file
    pathAsZip = tempDirectory & "\" & Replace(fileSystemObject.GetFileName(pathToKMZ), ".kmz", ".7z")
    fileSystemObject.CopyFile pathToKMZ, pathAsZip, True
    
    ' Loop through the items in the zip file and operate on the one that ends in ".kml"
    For Each f In applicationObject.Namespace(pathAsZip).Items
        If Right(f.path, 4) = ".kml" Then
            ' Copy the kml file to the temp folder with whatever name it had within the kmz
            applicationObject.Namespace(tempDirectory).CopyHere f.path, 20
            
            ' Copy the kml to the destination path, which includes the intended file name
            fileSystemObject.CopyFile tempDirectory & "\" & f.Name, returnPath, True
            
            ' Clean up the unnamed kml
            fileSystemObject.DeleteFile tempDirectory & "\" & f.Name
        End If
    Next
    
    ' Clean up the zip file
    fileSystemObject.DeleteFile pathAsZip
    
    ' If we made a temp folder, clean that up as well
    If (temp = "") Then
        fileSystemObject.DeleteFolder tempDirectory
    End If
End Sub

Sub CombineXMLandKMZ(ByVal xml As MSXML2.DOMDocument60, pathToSource As String, ByVal pathToTarget As String, ByVal FileNameNoExt As String)
    Dim applicationObject As Object
    Set applicationObject = CreateObject("Shell.Application")
    Dim fileSystemObject As Object
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")

    ' If the desired file is a KML, that means we don't need to do extra work, so we just save it
    xml.Save (pathToTarget)
    If (pathToSource Like "*.kml") Then
        Exit Sub
    End If
    
    ' Create a temporary working folder in the same folder as the source KMZ
    tempDirectory = ""
    parentFolderPath = fileSystemObject.GetParentFolderName(pathToSource)
    tempDirectory = fileSystemObject.CreateFolder(parentFolderPath & "\KML_CONVERSION_TEMP")

    ' Copy the source KMZ to the temp folder as a Zip file
    pathAsZip = tempDirectory & "\" & Replace(fileSystemObject.GetFileName(pathToSource), ".kmz", ".7z")
    fileSystemObject.CopyFile pathToSource, pathAsZip, True
    
    ' Loop through the items in the zip file and operate on the folder named "files"
    For Each f In applicationObject.Namespace(pathAsZip).Items
        If Right(f.path, 5) = "files" Then
            ' Copy the files folder to the temp directory
            applicationObject.Namespace(parentFolderPath).CopyHere f.path, 20
            
            ' Save the XML data as a KML file, also in the temp directory
            ' targetNoExt = Left(pathToTarget, Len(pathToTarget) - 4)
            ' xml.Save (tempDirectory & "\" & FileNameNoExt & ".kml")
            
            ' Zip up the KML and the files folder together
            ' fileSystemObject.DeleteFile pathAsZip
            ' zipFileName = targetNoExt + ".zip"
            ' NewZip (zipFileName)
            ' applicationObject.Namespace(zipFileName).CopyHere applicationObject.Namespace(tempDirectory & "\").Items

            
            ' Keep script waiting until Compressing is done
            ' On Error Resume Next
            ' counter = 0
            ' Do Until oApp.Namespace(FileNameZip).Items.Count = _
            '     applicationObject.Namespace(FolderName).Items.Count
            '     Application.Wait (Now + TimeValue("0:00:01"))
            '     counter = counter + 1
            '    If counter >= 10 Then
            '        MsgBox "Macro taking too long - could not compress files into zip. Delete temp folders and files before trying again"
            '         Exit Sub
            '     End If
            ' Loop
            ' On Error GoTo 0
            
            ' Move the KML to the desired location
            ' fileSystemObject.CopyFile (tempDirectory & "\" & FileNameNoExt & ".kml"), pathToTarget, True
        End If
    Next
    ' Clean up the temp folder
    fileSystemObject.DeleteFolder tempDirectory
    
    MsgBox "Your KMZ has been separated into a KML and a 'files' folder containing the icons." & vbNewLine & "If you want to combine these, open the KML in Google Earth and re-save it as a KMZ." & vbNewLine & "Then you can delete the KML and files folder."
End Sub

Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub


Function bIsBookOpen(ByRef szBookName As String) As Boolean
' Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function


Function Split97(sStr As Variant, sdelim As String) As Variant
'Tom Ogilvy
    Split97 = Evaluate("{""" & _
                       Application.Substitute(sStr, sdelim, """,""") & """}")
End Function


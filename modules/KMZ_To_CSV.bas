Attribute VB_Name = "KMZ_To_CSV"
Sub ConvertKMZToCSV()
    Dim Dash As Worksheet
    Set Dash = ThisWorkbook.Sheets("KMZ To CSV")
    Dim pathKMZ As String
    Dim uuidList As Dictionary
    Set uuidList = New Dictionary
    
    Dim PolygonsPath As String
    Dim XDoc As MSXML2.DOMDocument60
    Dim Namespace As String
    Dim polygonNodes As MSXML2.IXMLDOMNodeList
    Dim polygonNode As MSXML2.IXMLDOMNode
    Dim PolCoords As Variant
    Dim polData() As Variant
    
    ' Load Polygons XML
    Namespace = "xmlns:kml='http://www.opengis.net/kml/2.2'"
    PolygonsPath = CStr(ThisWorkbook.Sheets("KMZ Separator").Range("Polygons_KML").Value)
    Set PolygonsXDoc = Load_From_KML_Or_KMZ(PolygonsPath)
    
    ' Store the Polygon data to polData
    Set polygonNodes = PolygonsXDoc.SelectNodes("//kml:Polygon/parent::*")
    Set polygonNode = polygonNodes(0) ' Here we assume there's only one Polygon in the KMZ
    PolCoords = Split(polygonNode.SelectSingleNode(".//kml:coordinates").Text)
    ReDim polData(0 To UBound(PolCoords), 1 To 2)
    For j = 0 To UBound(PolCoords)
        polData(j, 1) = CDbl(Split(PolCoords(j), ",")(0))
        polData(j, 2) = CDbl(Split(PolCoords(j), ",")(1))
    Next j
    
    ' Get the path of the KMZ file
    pathKMZ = CStr(Dash.[KMZ_ToConvert].Value)
    
    ' If there isn't a KMZ path available, end the macro and inform the user
    If pathKMZ = "" Then
        ThisWorkbook.Sheets("Dashboard").Activate
        [KMZ_ToConvert].Select
        MsgBox ("KMZ_ToConvert is not set. Please select a file, then try again.")
        End
    End If
    
    ' If the file doesn't exist at the selected path, end the macro and inform the user
    KMZName = dir(pathKMZ)
    If KMZName = "" Then
        ThisWorkbook.Sheets("File Imports").Activate
        [Path_KMZ_Report].Select
        MsgBox ("File doesn't exist at path_KMZ_Report. Please select a different file, then try again.")
        End
    End If
    
    ' Load the file from the given path, whether it's a KML or KMZ file
    Dim inputKML As MSXML2.DOMDocument60
    Set inputKML = Load_From_KML_Or_KMZ(pathKMZ)

    ' Define which node types we want to extract
    ' Note: "point" means just the coordinates/name will be extracted, "connection" means the length will be extracted too
    Dim dataByType As Dictionary
    Set dataByType = New Dictionary
    With dataByType
        .Add "CUT MARK", "point"
        .Add "Comms Pole", "point"
        .Add "Guy Pole", "point"
        .Add "Joint Use  Transformer", "point" ' Sic: two spaces
        .Add "Joint Use", "point"
        .Add "Power", "point"
        .Add "Steel Transmission {Secondary}", "point"
        .Add "Transmission (Secondary)", "point"
        .Add "Transformer Pole", "point"
        .Add "VAULT", "point"
        .Add "Aerial Strand", "connection"
        .Add "overhead guy", "connection"
        .Add "overlash", "connection"
        .Add "slack span", "connection"
        .Add "underground cable", "connection"
    End With
    
    ' Make a dictionary that will be populated with the lists of XML nodes, separated by feature class
    Dim results As Dictionary
    Set results = New Dictionary
        
    ' Read in the nodes separated by feature class
    Dim nodeFolders As Variant
    Dim nodeFolder As Variant
    For Each fc In dataByType.Keys
        results.Add fc, inputKML.SelectNodes("//kml:Folder[kml:name='" & fc & "']//kml:Placemark[not(kml:styleUrl[contains(text(), '#Hidden')])]")
    Next
    
    ' Reset the output sheet
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets("KMZ Dump")
    Call Clear_KMZ_Output(ws)
    
    ' Output the new data
    Dim currentRow As Integer
    Dim includeLength As Boolean
    currentRow = 2
    
    For Each fc In results.Keys
        If dataByType(fc) = "connection" Then
            includeLength = True
        Else
            includeLength = False
        End If
    
        For Each Node In results(fc)
            ' Get the list of coordinate pairs
            ' NOTE: Each member of a pair is separated by a comma, but each pair is separated by a space
            Dim rawCoords As String
            Dim coords() As String
            Dim result As Boolean
            rawCoords = Node.SelectSingleNode(".//kml:coordinates").Text
            coords = Split(rawCoords) ' Split() defaults to split by space
            
            ' Evaluate whether any of this node's coordinates are within the polygon
            result = False
            For k = 0 To UBound(coords)
                ' NOTE: KMZs store coordinate pairs backwards, so we reverse the order here
                NodeLat = CDbl(Split(coords(k), ",")(1))
                NodeLong = CDbl(Split(coords(k), ",")(0))
                If PtInPoly(NodeLong, NodeLat, polData) = True Then
                    result = True
                End If
            Next k
            
            ' If any coordinate pair is inside the polygon, we output the object to the sheet. Otherwise, skip it
            If result = True Then
                ' NOTE: We're just using the most recently saved lat/long, which should be fine
                ws.Cells(currentRow, 1).Value = NodeLat ' Lat
                ws.Cells(currentRow, 2).Value = NodeLong ' Long
                ws.Cells(currentRow, 3).Value = fc ' Node type
                
                ' Node Name
                Dim nameNode As Variant
                Set nameNode = Node.SelectSingleNode("kml:name")
                If Not nameNode Is Nothing Then
                    ws.Cells(currentRow, 4).Value = nameNode.Text
                End If
                
                ' Connection Length (if applicable)
                If includeLength Then
                    Dim desc As String
                    desc = Node.SelectSingleNode(".//kml:description").Text
                    
                    ws.Cells(currentRow, 5).Value = ExtractBetween(desc, "Length</td><td style=""border: 1px solid black;"">", "</td>")
                    ws.Cells(currentRow, 6).Value = rawCoords ' Output the entire coordinate string, so each pair can be accessed
                End If
                
                currentRow = currentRow + 1
            End If
        Next Node
    Next
    
    ws.Columns.AutoFit
    ws.Activate
End Sub


Sub Clear_KMZ_Output(ws As Worksheet)
    Dim lastRow As Integer
    lastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    
    If Not lastRow = 1 Then
        ws.Range("A2:A" & lastRow).EntireRow.ClearContents
    End If
End Sub

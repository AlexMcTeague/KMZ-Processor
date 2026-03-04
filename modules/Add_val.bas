Attribute VB_Name = "Add_val"
Sub Split_AddVal()
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
    
    Set Dash = ThisWorkbook.Sheets("Address Separator")
    Namespace = "xmlns:kml='http://www.opengis.net/kml/2.2'"
    PolygonsPath = Dash.Range("Polygon_AddVal_Separator").Value
    AddValPath = Dash.Range("Add_Val_1").Value
    NewValPath = Dash.Range("New_AddVal").Value
    
    'Load Polygons XML
    Set PolygonsXDoc = Load_From_KML_Or_KMZ(PolygonsPath)
    
    'Get Nodes that have a child "Polygon" within the polygon XML
    Set polygonNodes = PolygonsXDoc.SelectNodes("//kml:Polygon/parent::*")
    
    Set polygonNode = polygonNodes(0)
    polName = polygonNode.SelectSingleNode(".//kml:name").Text
    'Extract the coordinates
    PolCoords = Split(polygonNode.SelectSingleNode(".//kml:coordinates").Text)
    ReDim polData(0 To UBound(PolCoords), 1 To 2)
    For j = 0 To UBound(PolCoords)
        polData(j, 1) = CDbl(Split(PolCoords(j), ",")(0))
        polData(j, 2) = CDbl(Split(PolCoords(j), ",")(1))
    Next j
    
    Set AddValWb = OpenPath(Dash.Range("Add_Val_1"))
    Set NewValWb = OpenPath(Dash.Range("New_AddVal"))
    
    LR = AddValWb.Sheets("Sheet1").Cells(Rows.Count, "B").End(xlUp).Row
    CR = 2
    
    For i = 2 To LR
        coords = AddValWb.Sheets("Sheet1").Range("B" & i).Value
        If coords Like "*,*" Then
            cLat = CDbl(Split(coords, ",")(0))
            cLong = CDbl(Split(coords, ",")(1))
        ElseIf coords Like "*, *" Then
            cLat = CDbl(Split(coords, ", ")(0))
            cLong = CDbl(Split(coords, ", ")(1))
        ElseIf coords Like "* *" Then
            cLat = CDbl(Split(coords, " ")(0))
            cLong = CDbl(Split(coords, " ")(1))
        Else
            MsgBox ("Unhandled coordinates on row " & i)
        End If
        
        If PtInPoly(cLong, cLat, polData) = True Then
            AddValWb.Sheets("Sheet1").Rows(i).Copy Destination:=NewValWb.Sheets("Sheet1").Rows(CR)
            CR = CR + 1
        End If
    Next i
End Sub

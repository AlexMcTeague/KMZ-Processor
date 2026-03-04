Attribute VB_Name = "Address_Count"
Sub Count_Addresses()
    Dim CalcXDoc As MSXML2.DOMDocument60
    Dim Namespace As String
    Dim CalcNodes As MSXML2.IXMLDOMNodeList
    Dim FullCalcNodes As MSXML2.IXMLDOMNodeList
    Dim CalcNode As MSXML2.IXMLDOMNode
    Dim polName As String
    Dim coords As Variant
    Dim polData() As Variant
    Dim XPath As String
    Dim CalcPath As String
    
    Namespace = "xmlns:kml='http://www.opengis.net/kml/2.2'"
    Set Dash = ThisWorkbook.Sheets("KMZ Address Counter")
    CalcPath = Dash.Range("KMZ_ToCalc").Value
    
    Dash.Range("G9:G14").ClearContents
    
    'Load Polygons XML
    Set CalcXDoc = Load_From_KML_Or_KMZ(CalcPath)

    'Get Placemark Nodes that have a styleUrl representing a red house
    Set CalcNodes = CalcXDoc.SelectNodes("//kml:Placemark[contains(kml:styleUrl,'#i118')]")
    Dash.Range("G9").Value = CalcNodes.Length
    
    'Get Placemark Nodes that have a styleUrl representing a blue house
    Set CalcNodes = CalcXDoc.SelectNodes("//kml:Placemark[contains(kml:styleUrl,'#i120')]")
    Dash.Range("G10").Value = CalcNodes.Length
    
    'Get Placemark Nodes that have a styleUrl representing an orange house
    Set CalcNodes = CalcXDoc.SelectNodes("//kml:Placemark[contains(kml:styleUrl,'#i119')]")
    Dash.Range("G11").Value = CalcNodes.Length
End Sub

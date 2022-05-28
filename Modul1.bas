Attribute VB_Name = "Modul1"
'Calculate Google Maps distance between two addresses
Public Function GetDistance(start As String, dest As String)
    Dim firstVal As String, secondVal As String, lastVal As String
    firstVal = "https://maps.googleapis.com/maps/api/distancematrix/json?origins="
    secondVal = "&destinations="
    
    ' In der folgenden Zeile muss der Google Cloud Key eingesetzt werden
    lastVal = "&mode=car&language=pl&sensor=false&key=<DeinGoogleCloudKey>"
    
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    Url = firstVal & Replace(start, " ", "+") & secondVal & Replace(dest, " ", "+") & lastVal
    objHTTP.Open "GET", Url, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ("")
    If InStr(objHTTP.responseText, """distance"" : {") = 0 Then GoTo ErrorHandl
    ' Set regex = CreateObject("VBScript.RegExp"): regex.Pattern = """value"".*?([0-9]+)": regex.Global = False
    Set regex = CreateObject("VBScript.RegExp"): regex.Pattern = """value"".*?([0-9]+)": regex.Global = False
    Set matches = regex.Execute(objHTTP.responseText)
    tmpVal = Replace(matches(0).SubMatches(0), ".", Application.International(xlListSeparator))
    umrechnung = Round((tmpVal / 1000), 1)
    GetDistance = CDbl(umrechnung)
    Exit Function
ErrorHandl:
    GetDistance = -1
    
End Function


'Ermittle Google Maps Zusammenfassung der Route zwischen 2 Adressen
Public Function GetRouteSummary(start As String, dest As String)
    Dim firstVal As String, secondVal As String, lastVal As String
    firstVal = "https://maps.googleapis.com/maps/api/directions/json?origin="
    secondVal = "&destination="
    
    ' In der folgenden Zeile muss der Google Cloud Key eingesetzt werden
    lastVal = "&key=<DeinGoogleCloudKey>"
    
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    Url = firstVal & Replace(start, " ", "+") & secondVal & Replace(dest, " ", "+") & lastVal
    Debug.Print Url
    objHTTP.Open "GET", Url, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ("")

    If InStr(objHTTP.responseText, """summary"" :") = 0 Then GoTo ErrorHandl
    
    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = "summary"" :(.*?),"
    regex.Global = False

    Set matches = regex.Execute(objHTTP.responseText)
    
    tmpVal = Replace(matches(0).SubMatches(0), ".", Application.International(xlListSeparator))
    Debug.Print tmpVal
    GetRouteSummary = tmpVal
    Exit Function
    
ErrorHandl:
    Text2 = "Error im Antorttext"
    GetRouteSummary = Text2
    
End Function

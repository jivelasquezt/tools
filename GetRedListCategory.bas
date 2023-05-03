Attribute VB_Name = "Module1"
Function GetRedListCategory(scientificName As String) As String
    Dim apiKey As String
    Dim url As String
    Dim command As String
    Dim result As String
    scientificName = Replace(scientificName, " ", "%20")
    apiKey = "your key"
    url = "http://apiv3.iucnredlist.org/api/v3/species/" & scientificName & "?token=" & apiKey

    ' Download the JSON response from the API endpoint using curl
    command = "curl " & url
    Dim script As String
    script = "do shell script """ & command & """"
    Dim output As String
    output = MacScript(script)
    Debug.Print output
    
    ' Find the "category" value in the JSON response
    Dim startIndex As Long
    Dim endIndex As Long
    startIndex = InStr(output, """category"":""") + 12 ' Add 12 to skip over the length of the "category" key
    endIndex = InStr(startIndex, output, """")
    GetRedListCategory = Mid(output, startIndex, endIndex - startIndex)
End Function



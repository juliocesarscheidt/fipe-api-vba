# VBA Code

```vba

Public Sub getBrands()
    ' clear results
    Sheets("data").Range("data!A2:B10000").ClearContents
    Sheets("data").Range("data!D2:E10000").ClearContents
    Sheets("data").Range("data!G2:H10000").ClearContents
    ' clear indexes
    Sheets("main").Range("main!P4").Value = 1
    Sheets("main").Range("main!K12").Value = 1
    Sheets("main").Range("main!P12").Value = 1
    ' clear previous fipe values
    Call clearPreviousFipeValues

    Dim apiUrl$, apiResult$
    apiUrl = getApiUrlBase() _
        & getSelectedVehicleType() _
        & "/" & getApiToken()
    apiResult = REQUEST(apiUrl, "GET")

    Dim apiResultJson As Object
    Set apiResultJson = JsonConverter.ParseJson(apiResult)

    Dim index As Integer
    index = 2
    
    Dim jsonElement As Dictionary
    For Each jsonElement In apiResultJson
        Range("data!A" & index).Value = Strings.LCase(jsonElement("nomeMarca"))
        Range("data!B" & index).Value = jsonElement("codMarca")
        index = index + 1
    Next
End Sub


Public Sub getModels()
    ' clear results
    'Sheets("data").Range("data!A2:B10000").ClearContents
    Sheets("data").Range("data!D2:E10000").ClearContents
    Sheets("data").Range("data!G2:H10000").ClearContents
    ' clear indexes
    'Sheets("main").Range("main!P4").Value = 1
    Sheets("main").Range("main!K12").Value = 1
    Sheets("main").Range("main!P12").Value = 1
    ' clear previous fipe values
    Call clearPreviousFipeValues

    If getSelectedVehicleBrand() = "-" Then
        Exit Sub
    End If

    Dim apiUrl$, apiResult$
    apiUrl = getApiUrlBase() _
        & getSelectedVehicleType() _
        & "/" & getSelectedVehicleBrand() _
        & "/" & getApiToken()
    apiResult = REQUEST(apiUrl, "GET")

    Dim apiResultJson As Object
    Set apiResultJson = JsonConverter.ParseJson(apiResult)

    Dim index As Integer
    index = 2
    
    Dim jsonElement As Dictionary
    For Each jsonElement In apiResultJson
        Range("data!D" & index).Value = Strings.LCase(jsonElement("nomeModelo"))
        Range("data!E" & index).Value = jsonElement("codModelo")
        index = index + 1
    Next
End Sub


Public Sub getModelsYears()
    ' clear results
    'Sheets("data").Range("data!A2:B10000").ClearContents
    'Sheets("data").Range("data!D2:E10000").ClearContents
    Sheets("data").Range("data!G2:H10000").ClearContents
    ' clear indexes
    'Sheets("main").Range("main!P4").Value = 1
    'Sheets("main").Range("main!K12").Value = 1
    Sheets("main").Range("main!P12").Value = 1
    ' clear previous fipe values
    Call clearPreviousFipeValues

    If getSelectedVehicleBrand() = "-" Or getSelectedVehicleModel() = "-" Then
        Exit Sub
    End If
    
    Dim apiUrl$, apiResult$
    apiUrl = getApiUrlBase() _
        & getSelectedVehicleType() _
        & "/" & getSelectedVehicleBrand() _
        & "/" & getSelectedVehicleModel() _
        & "/" & getApiToken()
    apiResult = REQUEST(apiUrl, "GET")

    Dim apiResultJson As Object
    Set apiResultJson = JsonConverter.ParseJson(apiResult)

    Dim index As Integer
    index = 2
    
    Dim jsonElement As Dictionary
    For Each jsonElement In apiResultJson
        Range("data!G" & index).Value = Strings.LCase(jsonElement("nomeAno"))
        Range("data!H" & index).Value = jsonElement("codAno")
        index = index + 1
    Next
End Sub


Public Sub getFipe()
    ' clear previous fipe values
    Call clearPreviousFipeValues

    If getSelectedVehicleBrand() = "-" Or getSelectedVehicleModel() = "-" Or getSelectedVehicleModelYear() = "-" Then
        Exit Sub
    End If

    Dim apiUrl$, apiResult$
    apiUrl = getApiUrlBase() _
        & getSelectedVehicleType() _
        & "/" & getSelectedVehicleBrand() _
        & "/" & getSelectedVehicleModel() _
        & "/" & getSelectedVehicleModelYear() _
        & "/" & getApiToken()
    apiResult = REQUEST(apiUrl, "GET")
    
    Dim apiResultJson As Object
    Set apiResultJson = JsonConverter.ParseJson(apiResult)

    Range("main!M22").Value = CDbl(Replace(Replace(apiResultJson("valorVeiculo"), "R$ ", ""), ".", ""))
    Range("main!N22").Value = apiResultJson("codFipe")
End Sub


Public Function getApiUrlBase() As String
    getApiUrlBase = Range("config!B1").Value
End Function

Public Function getApiToken() As String
    getApiToken = Range("config!B2").Value
End Function


Public Function getSelectedVehicleType() As String
    getSelectedVehicleType = Range("main!L4").Value
End Function

Public Function getSelectedVehicleBrand() As String
    getSelectedVehicleBrand = Range("main!Q4").Value
End Function

Public Function getSelectedVehicleModel() As String
    getSelectedVehicleModel = Range("main!L12").Value
End Function

Public Function getSelectedVehicleModelYear() As String
    getSelectedVehicleModelYear = Range("main!Q12").Value
End Function


Public Function REQUEST(ByVal apiUrl$, ByVal method$, Optional ByVal jsonDataString$, Optional ByVal bearerToken$, Optional ByVal basicToken$) As String
    Dim objHTTP As Object
    Dim responseCode$, responseText$
    
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
    objHTTP.Open method, apiUrl, False
    objHTTP.setRequestHeader "Content-type", "application/json"
    
    'setting oauth token when provided
    If bearerToken <> "" Then
        objHTTP.setRequestHeader "Authorization", "Bearer " & bearerToken
    End If
    
    'setting oauth token when provided
    If basicToken <> "" Then
        objHTTP.setRequestHeader "Authorization", "Basic " & basicToken
    End If

    'setting payload when provided
    If Not jsonDataString = "" Then
        objHTTP.Send (jsonDataString)
    Else
        objHTTP.Send
    End If
    
    responseCode = objHTTP.Status
    responseText = objHTTP.responseText
    
    Set objHTTP = Nothing
    
    'returns responseText
    REQUEST = responseText
End Function

```

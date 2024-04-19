Attribute VB_Name = "Functions"
'VBA에서는 static method를 제공하지 않으므로, class module에 메소드 작성 시에는 클래스의 인스턴스 생성 후에 사용을 해야합니다.
'고로 이러한 방식을 피하기 위해서 자주 사용하는 함수는 Functions모듈에 모으게 되었습니다.
'SOC(Separation of Concern), 모듈 구조

'@method GetHttpResponseText
'@param {String} url
'@return {String}
'@usage - HTTP GET request를 하고, response text를 돌려 받는 함수
Public Function GetHttpResponseText(url As String) As String
    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    With httpRequest
        .Open "GET", url, False
        .Send
        GetHttpResponseText = .responseText
    End With
End Function

'@method ParseJson
'@param {String} jsonString
'@return {Object}
'@usage - JSON string을 VBA 객체로 파싱하는 함수
Public Function ConverterParseJson(jsonString As String) As Object
    ' JsonConverter 모듈이 존재하고 있어야 한다.
    Set ConverterParseJson = JsonConverter.ParseJson(jsonString)
End Function

'@method GetJsonResponse
'@param {String} url
'@return {Object}
'@usage - GetHttpResponseText함수와 ParseJson함수 이용. url을 input으로 넣으면, 해당 json string을 파싱해주고, 결과 vba 객체를 return한다.
Public Function GetJsonResponse(url As String) As Object
    Dim jsonString As String
    jsonString = GetHttpResponseText(url)
    Set GetJsonResponse = ConverterParseJson(jsonString)
End Function

'@method URLEncode
'@param {String} StringVal
'@param {Boolean} [SpaceAsPlus=False]
'@return {String}
'@usage - StringVal을 인코딩해주는 함수. EncodeChar함수 이용
Public Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String
    Dim StringLen As Long: StringLen = Len(StringVal)

    If StringLen > 0 Then
        ReDim result(StringLen) As String 'StringLen길이로 string의 dynamic array를 result로 초기화한다.
        Dim i As Long

        For i = 1 To StringLen
            result(i) = EncodeChar(Mid$(StringVal, i, 1), SpaceAsPlus) 'dynamic array인 result에 encoding된 character가 순서대로 저장된다.
        Next i

        URLEncode = Join(result, "") 'Join함수: result array의 모든 요소를 empty string("")을 separator로 다 붙여 버린다. 결국에는 모든 character를 이어서 encoding된 string을 만든다.
    End If
End Function

'@method EncodeChar
'@param {String} Char
'@param {Boolean} SpaceAsPlus
'@return {String}
'@usage - 하나의 character인 string을 input으로 받는다. 하나의 character인 string을 인코딩하는 메소드이다.
'SpaceAsPlus는 space가 encoding되는 방식을 결정한다. true일 시, +로, false일 시, %20으로 인코딩 된다.
Private Function EncodeChar(Char As String, SpaceAsPlus As Boolean) As String
    Dim CharCode As Integer: CharCode = Asc(Char)
    Dim Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
            EncodeChar = Char
        Case 32
            EncodeChar = Space
        Case 0 To 15
            EncodeChar = "%0" & Hex(CharCode)
        Case Else
            EncodeChar = "%" & Hex(CharCode)
    End Select
End Function
'@method IsDuplicate
'@param {String} dataId1
'@param {String} dataId2
'@param {Dictionary} uniquePairs
'@return {Boolean} - data ID pair가 중복이면, Returns True. 아니면 False
'@usage - 함수는 2개의 dataId와 한개의 dictionary를 받는다. 함수는 dataId1과 dataId2의 조합 혹은 dataId2와 dataId1 조합이 dictionary에 존재하는지 확인한다.
'duplicate을 나타내는, dictionary에서 combination중 하나가 발견되면 True를 return한다.
Function IsDuplicate(ByVal dataId1 As String, ByVal dataId2 As String, ByRef uniquePairs As Scripting.Dictionary) As Boolean
    ' Check if the combinedKey or its reverse is in the dictionary
    IsDuplicate = uniquePairs.Exists(dataId1 & ":" & dataId2) Or uniquePairs.Exists(dataId2 & ":" & dataId1)
End Function

Function MapDataIdToCode(dataId As String) As String
    ' mapping함수
    Select Case dataId
        Case "KOSPI200_LOC"
            MapDataIdToCode = "KOSPI_LV"
        Case "HSI_LOC"
            MapDataIdToCode = "HSI_LV"
        Case "N225_LOC"
            MapDataIdToCode = "NKY_LV"
        Case "HSCEI_LOC"
            MapDataIdToCode = "HSCEI_LV"
        
        Case Else
            MapDataIdToCode = dataId ' mapping에 해당되는 경우가 없을 때는, 그대로 return한다.
    End Select
End Function



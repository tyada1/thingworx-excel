Attribute VB_Name = "ThingworxExcelConnector"
'Thingworx - Excel Connector VBA Code
'(c) Toshihiko Fujisawa - https://github.com/tofujisawa/thingworx-excel
'
' Depending on VBA-JSON library (https://github.com/VBA-tools/VBA-JSON)
'
' The MIT License (MIT)
'
' Copyright (c) 2016-2019 Tim Hall
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Sub GetTagButton_Click()
    getTagList
End Sub

Sub GetThingsButton_Click()
    Dim tag As String
    tag = Selection.Offset(0, -1).Value & ": " & Selection.Value
    Module1.getThingsFromTag (tag)
    
End Sub
Sub GetThingProperties_Click()
    Dim thingName As String
    thingName = Selection.Value
    Module1.getThingProperties (thingName)
    
End Sub

Function getTagList()
    Dim host, port, appKey, serviceLoc, url As String
    host = Range("B1").Value
    port = Range("B2").Value
    appKey = Range("B3").Value
    service = "/Thingworx/Resources/SearchFunctions/Services/SearchVocabularyTerms"
    url = host & ":" & port & service
    
    Dim params As New Dictionary
    params("maxItems") = 100
    params("MaxSerchItems") = 1000
    
    Dim res As Dictionary
    Set res = KickWebApiOfJson("POST", url, appKey, params)
    
    Range("A8:F50").Clear
    
    Dim i As Integer
    For i = 1 To res("rows").Count
        Cells(8 + i - 1, 1) = res("rows")(i)("vocabulary")
        Cells(8 + i - 1, 2) = res("rows")(i)("vocabularyTerm")
    Next i
    
End Function

Function getThingsFromTag(ByVal tag As String)
    Dim host, port, appKey, serviceLoc, url As String
    host = Range("B1").Value
    port = Range("B2").Value
    appKey = Range("B3").Value
    service = "/Thingworx/Resources/SearchFunctions/Services/SearchThings"
    url = host & ":" & port & service
    
    Dim params As New Dictionary
    params("modelTags") = tag
    
    Dim res As Dictionary
    Set res = KickWebApiOfJson("POST", url, appKey, params)
        
    Dim thingsCount As Integer
    thingsCount = res("rows")(1)("commonResults")("rows").Count
    
    Range("C8:F50").Clear
    
    Dim i As Integer
    For i = 1 To thingsCount
        Cells(8 + i - 1, 3) = res("rows")(1)("commonResults")("rows")(i)("name")
        Cells(8 + i - 1, 4) = res("rows")(1)("commonResults")("rows")(i)("description")
    Next i
    
End Function

Function getThingProperties(ByVal thingName)
    Dim host, port, appKey, serviceLoc, url As String
    host = Range("B1").Value
    port = Range("B2").Value
    appKey = Range("B3").Value
    service = "/Thingworx/Things/" & thingName & "/Properties"
    url = host & ":" & port & service

    Dim res As Dictionary
    Set res = KickWebApiOfJson("GET", url, appKey)
    
    Range("E8:F50").Clear
    
    Dim i As Integer
    i = 0
    For Each Var In res("rows")(1)
        If ((Var <> "tags") And (Var <> "name") And (Var <> "description") And (Var <> "thingTemplate")) Then
        Cells(8 + i, 5) = Var
        Cells(8 + i, 6) = res("rows")(1)(Var)
        i = i + 1
        End If
    Next Var
    
    
End Function

Function KickWebApiOfJson(ByVal request As String, ByVal url As String, ByVal appKey As String, Optional ByVal params As Object) As Object
    Dim json
    json = ConvertToJson(params)

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open request, url, False
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "appKey", appKey
        .SetRequestHeader "Content-Type", "application/json"
        .send json

        If .ResponseText <> "" Then
            Set KickWebApiOfJson = ParseJson(.ResponseText)
        End If
    End With
End Function

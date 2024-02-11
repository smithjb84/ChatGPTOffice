

'=====================================================
' GET YOUR API KEY: https://openai.com/api/
Const API_KEY As String = ""
'=====================================================

' Constants for API endpoint and request properties
Const API_ENDPOINT As String = "https://api.openai.com/v1/chat/completions"
Const MODEL As String = "gpt-3.5-turbo"
Const MAX_TOKENS As String = "1024"
Const TEMPERATURE As String = "0.5"

'Output worksheet name
Const OUTPUT_WORKSHEET As String = "Result"

Sub ReWriteSection()
       

          ' Check if API key is available
        If API_KEY = "<API_KEY>" Then
            MsgBox "Please input a valid API key. You can get one from https://openai.com/api/", vbCritical, "No API Key Found"
           Application.ScreenUpdating = True
            Exit Sub
        End If

          ' Get the prompt
           
Dim OutApp As Object
Dim OutMail As Object
Dim olInsp As Object
Dim wdDoc As Object
Dim strText As String

    On Error Resume Next
    'Get Outlook if it's running
    Set OutApp = GetObject(, "Outlook.Application")


    Set OutMail = OutApp.ActiveExplorer.Selection.Item(1)
    With OutMail
        Set olInsp = .GetInspector
        Set wdDoc = olInsp.WordEditor
        strText = "rewrite the following text " & wdDoc.Application.Selection.Range.text
    End With

          
          If Trim(strText) <> "" Then
              ' Clean prompt to avoid parsing error in JSON payload
           prompt = CleanJSONString(strText)
       Else
           MsgBox "Please enter some text in the selected cell before executing the macro", vbCritical, "Empty Input"
           Application.ScreenUpdating = True
           Exit Sub
       End If

          ' Create XMLHTTP object
          Dim httpRequest As Object
      Set httpRequest = CreateObject("MSXML2.XMLHTTP")

          ' Define request body
          Dim requestBody As String
      requestBody = "{" & _
              """model"": """ & MODEL & """," & _
              """messages"": [{""role"":""user"", ""content"": """ & prompt & """}]," & _
              """max_tokens"": " & MAX_TOKENS & "," & _
              """temperature"": " & TEMPERATURE & _
              "}"
              
          ' Open and send the HTTP request
      With httpRequest
           .Open "POST", API_ENDPOINT, False
           .SetRequestHeader "Content-Type", "application/json"
           .SetRequestHeader "Authorization", "Bearer " & API_KEY
           .Send (requestBody)
       End With

          'Check if the request is successful
       If httpRequest.Status = 200 Then
              'Parse the JSON response
              Dim response As String
           response = httpRequest.responseText
           

              'Get the completion and clean it up
              Dim completion As String
           completion = ParseResponse(response)
              
              'Split the completion into lines
              Dim lines As Variant
           lines = Split(completion, "\n")

              'Write the lines to the worksheet
              Dim i As Long
              Dim MyText As String
              Dim inspector As Outlook.inspector
    Set inspector = Application.ActiveInspector
    Dim wordDoc As Word.Document
            Set wordDoc = inspector.WordEditor
           For i = LBound(lines) To UBound(lines)
           If lines(i) <> "" Then
                wordDoc.Application.Selection.text = lines(i) & vbCrLf
                End If
           Next i
              
       Else
           MsgBox "Request failed with status " & httpRequest.Status & vbCrLf & vbCrLf & "ERROR MESSAGE:" & vbCrLf & httpRequest.responseText, vbCritical, "OpenAI Request Failed"
       End If
          
       Exit Sub
          
End Sub


Sub WriteSection()

          ' Check if API key is available
        If API_KEY = "<API_KEY>" Then
            MsgBox "Please input a valid API key. You can get one from https://openai.com/api/", vbCritical, "No API Key Found"
           Application.ScreenUpdating = True
            Exit Sub
        End If

          ' Get the prompt
           
Dim OutApp As Object
Dim OutMail As Object
Dim olInsp As Object
Dim wdDoc As Object
Dim strText As String

    On Error Resume Next
    'Get Outlook if it's running
    Set OutApp = GetObject(, "Outlook.Application")


    Set OutMail = OutApp.ActiveExplorer.Selection.Item(1)
    With OutMail
        Set olInsp = .GetInspector
        Set wdDoc = olInsp.WordEditor
        strText = wdDoc.Application.Selection.Range.text
    End With

          
          If Trim(strText) <> "" Then
              ' Clean prompt to avoid parsing error in JSON payload
           prompt = CleanJSONString(strText)
       Else
           MsgBox "Please enter some text in the selected cell before executing the macro", vbCritical, "Empty Input"
           Application.ScreenUpdating = True
           Exit Sub
       End If
     

          ' Create XMLHTTP object
          Dim httpRequest As Object
      Set httpRequest = CreateObject("MSXML2.XMLHTTP")

          ' Define request body
          Dim requestBody As String
      requestBody = "{" & _
              """model"": """ & MODEL & """," & _
              """messages"": [{""role"":""user"", ""content"": """ & prompt & """}]," & _
              """max_tokens"": " & MAX_TOKENS & "," & _
              """temperature"": " & TEMPERATURE & _
              "}"
              
          ' Open and send the HTTP request
      With httpRequest
           .Open "POST", API_ENDPOINT, False
           .SetRequestHeader "Content-Type", "application/json"
           .SetRequestHeader "Authorization", "Bearer " & API_KEY
           .Send (requestBody)
       End With

          'Check if the request is successful
       If httpRequest.Status = 200 Then
              'Parse the JSON response
              Dim response As String
           response = httpRequest.responseText
           

              'Get the completion and clean it up
              Dim completion As String
           completion = ParseResponse(response)
              
              'Split the completion into lines
              Dim lines As Variant
           lines = Split(completion, "\n")

              'Write the lines to the worksheet
              Dim i As Long
              Dim MyText As String
              Dim inspector As Outlook.inspector
    Set inspector = Application.ActiveInspector
    Dim wordDoc As Word.Document
            Set wordDoc = inspector.WordEditor
           For i = LBound(lines) To UBound(lines)
           If lines(i) <> "" Then
                wordDoc.Application.Selection.text = lines(i) & vbCrLf
                End If
           Next i
              
       Else
           MsgBox "Request failed with status " & httpRequest.Status & vbCrLf & vbCrLf & "ERROR MESSAGE:" & vbCrLf & httpRequest.responseText, vbCritical, "OpenAI Request Failed"
       End If
          
       Exit Sub
          
End Sub
' Helper function to check if worksheet exists
Function WorksheetExists(worksheetName As String) As Boolean
       On Error Resume Next
       WorksheetExists = (Not (Sheets(worksheetName) Is Nothing))
       On Error GoTo 0
End Function
' Helper function to clean text
Function CleanJSONString(inputStr As String) As String
       On Error Resume Next
          ' Remove line breaks
       CleanJSONString = Replace(inputStr, vbCrLf, "")
       CleanJSONString = Replace(CleanJSONString, vbCr, "")
       CleanJSONString = Replace(CleanJSONString, vbLf, "")

          ' Replace all double quotes with single quotes
       CleanJSONString = Replace(CleanJSONString, """", "'")
       On Error GoTo 0
End Function
' Replaces the backslash character only if it is immediately followed by a double quote.
Function ReplaceBackslash(text As Variant) As String
       On Error Resume Next
          Dim i As Integer
          Dim newText As String
       newText = ""
       For i = 1 To Len(text)
           If Mid(text, i, 2) = "\" & Chr(34) Then
               newText = newText & Chr(34)
               i = i + 1
           Else
               newText = newText & Mid(text, i, 1)
           End If
       Next i
       ReplaceBackslash = newText
       On Error GoTo 0
End Function

Function ParseResponse(ByVal response As String) As String
       On Error Resume Next
          Dim startIndex As Long
       startIndex = InStr(response, """content"":") + 12
          Dim endIndex As Long
       endIndex = InStr(response, """logprobs"":") - 17
       ParseResponse = Mid(response, startIndex, endIndex - startIndex)
       On Error GoTo 0
End Function

Sub test()
Dim OutApp As Object
Dim OutMail As Object
Dim olInsp As Object
Dim wdDoc As Object
Dim strText As String

    On Error Resume Next
    'Get Outlook if it's running
    Set OutApp = GetObject(, "Outlook.Application")

    'Outlook wasn't running, so cancel
    If Err <> 0 Then
        MsgBox "Outlook is not running so nothing can be selected!"
        GoTo lbl_Exit
    End If
    On Error GoTo 0

    Set OutMail = OutApp.ActiveExplorer.Selection.Item(1)
    With OutMail
        Set olInsp = .GetInspector
        Set wdDoc = olInsp.WordEditor
        strText = wdDoc.Application.Selection.Range.text
    End With
    MsgBox strText
lbl_Exit:
    Set OutMail = Nothing
    Set OutApp = Nothing
    Set olInsp = Nothing
    Set wdDoc = Nothing
    Exit Sub
End Sub



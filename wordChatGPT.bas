

'=====================================================
' GET YOUR API KEY: https://openai.com/api/
Const API_KEY As String = ""
'=====================================================

' Constants for API endpoint and request properties
Const API_ENDPOINT As String = "https://api.openai.com/v1/chat/completions"
Const MODEL As String = "gpt-3.5-turbo"
Const MAX_TOKENS As String = "1024"
Const TEMPERATURE As String = "0.5"


Sub ReWriteSection()

        On Error GoTo ErrorHandler
        Application.ScreenUpdating = False

          ' Check if API key is available
        If API_KEY = "<API_KEY>" Then
            MsgBox "Please input a valid API key. You can get one from https://openai.com/api/", vbCritical, "No API Key Found"
           Application.ScreenUpdating = True
            Exit Sub
        End If

          ' Get the prompt
          Dim Sel As Selection
          Set Sel = Application.Selection
          Dim prompt As String
          prompt = "rewrite the following text " & Sel.text
          
          If Trim(prompt) <> "" Then
              ' Clean prompt to avoid parsing error in JSON payload
           prompt = CleanJSONString(prompt)
       Else
           MsgBox "Please enter some text in the selected cell before executing the macro", vbCritical, "Empty Input"
           Application.ScreenUpdating = True
           Exit Sub
       End If
     
          ' Show status in status bar
        Application.StatusBar = "Processing OpenAI request..."

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
           .send (requestBody)
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
           For i = LBound(lines) To UBound(lines)
                If lines(i) <> "" Then
                Selection.TypeText (ReplaceBackslash(lines(i)) & vbCrLf)
                End If
           Next i
              
       Else
           MsgBox "Request failed with status " & httpRequest.Status & vbCrLf & vbCrLf & "ERROR MESSAGE:" & vbCrLf & httpRequest.responseText, vbCritical, "OpenAI Request Failed"
       End If
          
       Application.StatusBar = False
       Application.ScreenUpdating = True
          
       Exit Sub
          
ErrorHandler:
       MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "Line: " & Erl, vbCritical, "Error"
       Application.StatusBar = False
       Application.ScreenUpdating = True
          
End Sub


Sub WriteSection()

        On Error GoTo ErrorHandler
        Application.ScreenUpdating = False

          ' Check if API key is available
        If API_KEY = "<API_KEY>" Then
            MsgBox "Please input a valid API key. You can get one from https://openai.com/api/", vbCritical, "No API Key Found"
           Application.ScreenUpdating = True
            Exit Sub
        End If

          ' Get the prompt
          Dim Sel As Selection
          Set Sel = Application.Selection
          Dim prompt As String
          prompt = Sel.text
          
          If Trim(prompt) <> "" Then
              ' Clean prompt to avoid parsing error in JSON payload
           prompt = CleanJSONString(prompt)
       Else
           MsgBox "Please enter some text in the selected cell before executing the macro", vbCritical, "Empty Input"
           Application.ScreenUpdating = True
           Exit Sub
       End If
     
          ' Show status in status bar
        Application.StatusBar = "Processing OpenAI request..."

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
           .send (requestBody)
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
           For i = LBound(lines) To UBound(lines)
                If lines(i) <> "" Then
                Selection.TypeText (ReplaceBackslash(lines(i)) & vbCrLf)
                End If
           Next i
              
       Else
           MsgBox "Request failed with status " & httpRequest.Status & vbCrLf & vbCrLf & "ERROR MESSAGE:" & vbCrLf & httpRequest.responseText, vbCritical, "OpenAI Request Failed"
       End If
          
       Application.StatusBar = False
       Application.ScreenUpdating = True
          
       Exit Sub
          
ErrorHandler:
       MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "Line: " & Erl, vbCritical, "Error"
       Application.StatusBar = False
       Application.ScreenUpdating = True
          
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




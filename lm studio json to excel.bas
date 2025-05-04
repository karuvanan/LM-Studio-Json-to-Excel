Attribute VB_Name = "Module1"
Sub ImportConversationToExcel()
    ' Purpose: Import JSON conversation data into Excel with emojis and thinking process
    ' Requirements: VBA-JSON library and Microsoft Scripting Runtime reference
    
    ' Let user select JSON file
    Dim filePath As Variant
    filePath = Application.GetOpenFilename("JSON Files (*.json), *.json", , "Select JSON File", , False)
    If filePath = False Then Exit Sub ' User canceled
    
    ' Read JSON file using ADODB.Stream
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile filePath
    Dim jsonString As String
    jsonString = objStream.ReadText
    objStream.Close
    
    ' Parse JSON string into a VBA object
    Dim jsonData As Object
    Set jsonData = JsonConverter.ParseJson(jsonString)
    
    ' Access the messages array
    Dim messages As Collection
    Set messages = jsonData("messages")
    
    ' Set the worksheet to write to
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    ws.Cells.Clear ' Clear existing data, optional
    
    ' Add headers
    ws.Cells(1, 1).Value = "Emoji"
    ws.Cells(1, 2).Value = "Content"
    
    ' Initialize row counter for Excel output
    Dim row As Long
    row = 2 ' Start from row 2
    
    ' Loop through each message in the conversation
    Dim i As Long
    For i = 1 To messages.Count
        Dim msg As Object
        Set msg = messages(i)
        
        ' Get the first version of the message
        Dim versions As Collection
        Set versions = msg("versions")
        Dim version As Object
        Set version = versions(1)
        
        ' Check the role (user or assistant)
        Dim role As String
        role = version("role")
        
        If role = "user" Then
            ' Extract user question
            Dim content As Collection
            Set content = version("content")
            Dim textObj As Object
            Set textObj = content(1)
            Dim question As String
            question = textObj("text")
            
            ' Write question to Column 2 with emoji in Column 1
            ws.Cells(row, 1).Value = ChrW(&HD83E) & ChrW(&HDD14) ' ??
            ws.Cells(row, 2).Value = question
            row = row + 1
        ElseIf role = "assistant" Then
            ' Extract assistant's thinking and answer
            Dim steps As Collection
            Set steps = version("steps")
            Dim j As Long
            Dim thinking As String
            Dim answer As String
            thinking = "No thinking found"
            answer = "No answer found"
            
            For j = 1 To steps.Count
                Dim stepObj As Object
                Set stepObj = steps(j)
                If stepObj("type") = "contentBlock" Then
                    ' Debug: Log step details
                    Debug.Print "Message " & i & ", Step " & j & ": Type=" & stepObj("type") & ", HasPrefix=" & stepObj.Exists("prefix") & ", Prefix=" & IIf(stepObj.Exists("prefix"), stepObj("prefix"), "None")
                    
                    If stepObj.Exists("prefix") And stepObj("prefix") = "<think>" Then
                        Dim thinkContent As Collection
                        Set thinkContent = stepObj("content")
                        If TypeName(thinkContent) = "Collection" And thinkContent.Count >= 1 Then
                            Dim thinkTextObj As Object
                            Set thinkTextObj = thinkContent(1)
                            If thinkTextObj.Exists("text") Then
                                thinking = thinkTextObj("text")
                            End If
                        End If
                    ElseIf stepObj.Exists("prefix") Then
                        Dim ansContent As Collection
                        Set ansContent = stepObj("content")
                        If TypeName(ansContent) = "Collection" And ansContent.Count >= 1 Then
                            Dim ansTextObj As Object
                            Set ansTextObj = ansContent(1)
                            If ansTextObj.Exists("text") Then
                                answer = ansTextObj("text")
                            End If
                        End If
                    End If
                End If
            Next j
            
            ' Write thinking to Column 2 with emoji in Column 1
            ws.Cells(row, 1).Value = ChrW(&HD83E) & ChrW(&HDD10) ' ??
            ws.Cells(row, 2).Value = thinking
            row = row + 1
            
            ' Write answer to Column 2 with emoji in Column 1
            ws.Cells(row, 1).Value = ChrW(&H2705) ' ?
            ws.Cells(row, 2).Value = answer
            row = row + 1
            
            ' Write empty row
            ws.Cells(row, 1).Value = ""
            ws.Cells(row, 2).Value = ""
            row = row + 1
        End If
    Next i
        ' Set column widths and alignments
    ws.Columns(1).ColumnWidth = 5
    ws.Columns(2).ColumnWidth = 75
    ws.Columns(1).VerticalAlignment = xlTop
    ws.Columns(2).VerticalAlignment = xlTop
    ws.Columns(2).WrapText = True
End Sub

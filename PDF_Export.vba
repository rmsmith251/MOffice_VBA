' Another function developed by my colleague Sam Adams.
' This pulls PDFs from provided hyperlinks.

Sub Full_TP_Export()
Dim rng As Range
For Each rng In Selection
    If rng.EntireRow.Hidden = False Then
            Dim WinHttpReq As Object
            Dim TP_number As String
            Dim Row_Num As Long
            Dim UserPath As String
            Dim TP_folder As String
            Dim modelURL As String
            Dim modelSTR As String
            Dim blindURL As String
            Dim blindSTR As String
            Dim instURL As String
            Dim instSTR As String
            Dim llURL As String
            Dim llSTR As String
            Dim isoURL As String
            Dim isoSTR As String
            Dim pidURL As String
            Dim pidSTR As String
            Dim psURL As String
            Dim psSTR As String
            Dim P01URL As String
            Dim P01STR As String
            Dim P02URL As String
            Dim P02STR As String
            Dim P04URL As String
            Dim P04STR As String
            Dim P05URL As String
            Dim P05STR As String
            Dim Q01URL As String
            Dim Q01STR As String
            Dim Q03URL As String
            Dim Q03STR As String
            Dim Q04URL As String
            Dim Q04STR As String
            Dim Q05URL As String
            Dim Q05STR As String
            Dim X01URL As String
            Dim X01STR As String
            Dim pecURL As String
            Dim pecSTR As String
            Dim ptrURL As String
            Dim ptrSTR As String
            Dim pcoverURL As String
            Dim pcoverSTR As String
            TP_folder = Environ("USERPROFILE") & "\Desktop\TEST_PACKS\"
            UserPath = Environ("USERPROFILE") & "\Desktop\"
            TP_number = rng.Value
            Row_Num = rng.Row
                If Dir(TP_folder, vbDirectory) = "" Then
                    MkDir TP_folder
                End If
                If Dir(TP_folder & TP_number & "\", vbDirectory) <> "" Then
                    Kill TP_folder & TP_number & "\*.*"
                    RmDir TP_folder & TP_number & "\"
                Else
                End If
                If Dir(TP_folder & TP_number & "\", vbDirectory) = "" Then
                    MkDir TP_folder & TP_number & "\"
                Else
                End If
'start pid section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 5").Value = 1 Then
            pidSTR = TP_number + " - PIDs"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & pidSTR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                pidURL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & pidSTR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", pidURL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\03_PID.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end pid section
'start iso section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 3").Value = 1 Then
            isoSTR = TP_number + " - ISOs"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & isoSTR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                isoURL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & isoSTR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", isoURL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\02_ISO_" & TP_number & ".pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end iso section
'start model section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 4").Value = 1 Then
            modelSTR = TP_number + " - Model Shot"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & modelSTR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                modelURL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & modelSTR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", modelURL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\11_MODEL.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end model section
'start supports section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 8").Value = 1 Then
            psSTR = TP_number + " - Pipe Support Drawings"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & psSTR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                psURL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & psSTR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", psURL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\10_SUPPORT_DRAWINGS.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end supports section
'start instrument section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 21").Value = 1 Then
            instSTR = TP_number + " - Instrument List"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & instSTR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                instURL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & instSTR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", instURL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\07_INSTRUMENT.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end instrument section
'start blind list section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 7").Value = 1 Then
            blindSTR = TP_number + " - Blinds List"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & blindSTR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                blindURL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & blindSTR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", blindURL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\06_BLIND_LIST.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end blind list section
'start line list section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 6").Value = 1 Then
            llSTR = TP_number + " - Line List"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & llSTR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                llURL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & llSTR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", llURL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\04_LINE_LIST.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end line list section
'start PEC section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 13").Value = 1 Then
            pecSTR = TP_number + " - Pneumatic Exclusion Calculation"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & pecSTR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                pecURL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & pecSTR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", pecURL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\14_PEC.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end PEC section
'start PTR section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 14").Value = 1 Then
            ptrSTR = TP_number + " - Pneumatic Test Ramp Chart"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & ptrSTR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                ptrURL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & ptrSTR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", ptrURL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\15_PTR_CHART.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end PTR section
'start PCover section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 24").Value = 1 Then
            pcoverSTR = TP_number + " - Pneumatic Test Cover Sheet"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & pcoverSTR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                pcoverURL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & pcoverSTR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", pcoverURL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\15_PCover.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end PCover section
'start P01 section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 9").Value = 1 Then
            P01STR = TP_number + " - P01A"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & P01STR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                P01URL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & P01STR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", P01URL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\01_P01A.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end P01 section
'start P02 section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 10").Value = 1 Then
            P02STR = TP_number + " - P02A"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & P02STR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                P02URL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & P02STR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", P02URL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\05_P02A.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end P02 section
'start P04 section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 11").Value = 1 Then
            P04STR = TP_number + " - P04A"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & P04STR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                P04URL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & P04STR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", P04URL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\08_P04A.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end P04 section
'start P05 section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 12").Value = 1 Then
            P05STR = TP_number + " - P05A"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & P05STR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                P05URL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & P05STR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", P05URL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\09_P05A.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end P05 section
'start Q01 section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 15").Value = 1 Then
            Q01STR = TP_number + " - Q01A"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & Q01STR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                Q01URL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & Q01STR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", Q01URL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\94_Q01A.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end Q01 section
'start Q03 section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 16").Value = 1 Then
            Q03STR = TP_number + " - Q03A"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & Q03STR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                Q03URL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & Q03STR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", Q03URL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\95_Q03A.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end Q03 section
'start Q04 section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 17").Value = 1 Then
            Q04STR = TP_number + " - Q04A"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & Q04STR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                Q04URL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & Q04STR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", Q04URL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\14_Q04A.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end Q04 section
'start Q05 section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 18").Value = 1 Then
            Q05STR = TP_number + " - Q05A"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & Q05STR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                Q05URL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & Q05STR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", Q05URL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\15_Q05A.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end Q05 section
'start X01 section
    If Sheets("Test Pack Tracker").CheckBoxes("Check Box 19").Value = 1 Then
            X01STR = TP_number + " - Q05A"
            'check if link exists
            If Application.Evaluate("=IFERROR(INDEX(Links!F:F,MATCH(""" & X01STR & """,Links!E:E,0)),"""")") = "" Then
                'if link does not do nothing
            Else
                X01URL = Application.Evaluate("=INDEX(Links!F:F,MATCH(""" & X01STR & """,Links!E:E,0))")
                Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
                WinHttpReq.Open "GET", X01URL, False, "username", "password"
                WinHttpReq.send
                If WinHttpReq.Status = 200 Then
                    Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile TP_folder & TP_number & "\12_X01A.pdf", 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
                End If
            End If
    End If
'end X01 section
'MERGE PDF FILES SECTION
Const DestFile As String = "00_TPMergedFile.pdf" ' <-- change to suit
   
    Dim MyPath As String, MyFiles As String
    Dim a() As String, i As Long, f As String

    MyPath = TP_folder & TP_number & "\"
   
      ' Populate the array a() by PDF file names
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    ReDim a(1 To 2 ^ 14)
    f = Dir(MyPath & "*.pdf")
    While Len(f)
        If StrComp(f, DestFile, vbTextCompare) Then
            i = i + 1
            a(i) = f
        End If
        f = Dir()
    Wend
   
    ' Merge PDFs
    If i Then
        ReDim Preserve a(1 To i)
        MyFiles = Join(a, ",")
        Application.StatusBar = "Merging, please wait ..."
        Call MergePDFs(MyPath, MyFiles, DestFile)
        Application.StatusBar = False
        i = 0
    Else
        MsgBox "No PDF files found in" & vbLf & MyPath, vbExclamation, "Canceled"
    End If
    
'END MERGE PDF
    End If
' Flatten test pack
'Dim MyFile As String
'MyFile = Dir(MyPath)
'Do While MyFile <> ""
'If MyFile Like "*TPMergedFile.PDF" Or MyFile Like "*TPMergedFile.pdf" Then
'Fullpath = MyPath & MyFile
'Set App = CreateObject("AcroExch.app")
'Set avdoc = CreateObject("AcroExch.AVDoc")
'Set pdDoc = CreateObject("AcroExch.PDDoc")
'Set AForm = CreateObject("AFormAut.App")
'pdDoc.Open (Fullpath)
'Set avdoc = pdDoc.OpenAVDoc(Fullpath)
'   js = "this.flattenPages();"
'     '//execute the js code
'    AForm.Fields.ExecuteThisJavaScript js
'
'Set pdDoc = avdoc.GetPDDoc
'pdDoc.Save PDSaveFull, Fullpath
'pdDoc.Close
'Set AForm = Nothing
'Set avdoc = Nothing
'Set App = Nothing
'End If
'MyFile = Dir
'Loop

Next rng

End Sub

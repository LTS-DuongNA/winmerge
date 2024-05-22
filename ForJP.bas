Sub GenerateTerminalJsonFile_20240123()
    Dim seriNumberData As Long
    Dim seriNumber As String
    Dim terminalResult As String
    terminalResult = ""
    ' Define last colum and last row used
    
'2024.01.29 A列の最終行を取得するように変更
'    Dim intLastCellIndexInRow As Integer
'    intLastCellIndexInRow = ActiveCell.SpecialCells(xlLastCell).Column
'    Dim intLastCellIndexInColumn As Integer
'    intLastCellIndexInColumn = ActiveCell.SpecialCells(xlLastCell).Row
'    ' Get data from sheet
'    Dim terminalStart As Integer
'    terminalStart = 4
'
'    ' Get all S/N's json
'    terminalResult = terminalResult & "{" & vbNewLine & ""
'    For seriNumberData = 13 To intLastCellIndexInColumn
'        seriNumber = Cells(seriNumberData, 2).Value
'        Dim terminalRange As Range: Set terminalRange = Range(Cells(seriNumberData, terminalStart), Cells(seriNumberData, terminalStart + terminalDataCount - 1))
'
'        terminalResult = terminalResult & """" & Cells(seriNumberData, 2).Value & """" & ":"
'        terminalResult = terminalResult & GetTerminalNewFormat(terminalRange, CStr(seriNumberData))
'
'        If seriNumberData <> intLastCellIndexInColumn Then
        

    Dim xlLastRow As Long                               'Excel自体の最終行
    Dim LastRow As Long                                 '最終行
 
        xlLastRow = Cells(Rows.Count, 1).Row            'Excelの最終行を取得
        LastRow = Cells(xlLastRow, 1).End(xlUp).Row     'A列の最終行を取得

'2024.04.19　A列に1が立っているデータの最終行を取得
    For i = 9 To LastRow
        If Cells(i, 1) = 1 Then
            LastRow = Cells(i, 1).Row
    End If
    Next i

'2024.04.19 END


    Dim terminalStart As Integer
        terminalStart = 4


'2024.04.02　実行前にA列に1が立っているデータがなければマクロを終了
    If WorksheetFunction.CountIf(Worksheets("Terminal").Range("A9:A" & LastRow), 1) = 0 Then
        MsgBox ("出力列に「1」の立っている行がないため、処理を終了します。")
        Exit Sub
    End If



'    ' Get all S/N's json
    terminalResult = terminalResult & "{" & vbNewLine & ""
'    For seriNumberData = 13 To LastRow
    For seriNumberData = 9 To LastRow                   '開始位置を9行目に変更
    
'2024.04.02　A列に1が立っているデータを出力対象とする
    If Cells(seriNumberData, 1) = 1 Then
   

        seriNumber = Cells(seriNumberData, 3).Value
        Dim terminalRange As Range: Set terminalRange = Range(Cells(seriNumberData, terminalStart), Cells(seriNumberData, terminalStart + terminalDataCount - 1))

        terminalResult = terminalResult & """" & Cells(seriNumberData, 3).Value & """" & ":"
        terminalResult = terminalResult & GetTerminalNewFormat(terminalRange, CStr(seriNumberData))

        If seriNumberData <> LastRow Then
        
'2024.01.29 END

            terminalResult = terminalResult & ","
            Else
            terminalResult = terminalResult & ""
            
        End If
        terminalResult = terminalResult & vbNewLine & ""
    
    
    
'2024.04.02　A列に1が立っているデータを出力対象とする
        Cells(seriNumberData, 1) = "済"
    End If
    
    
    Next
    terminalResult = terminalResult & "}"
    ' Export JSON

'    Dim fileDirTerminal As String: fileDirTerminal = Application.GetSaveAsFilename("Terminal.json", fileFilter:="JSON Files (*.json), *.json")
'    WriteUTF8WithoutBOM terminalResult, fileDirTerminal
    
    
'キャンセルボタン押下時falseファイルができないように対応

    Dim fileDirtenant As String
        fileDirtenant = Application.GetSaveAsFilename("Terminal.json", fileFilter:="JSON Files (*.json), *.json")

            If fileDirtenant = "False" Then
              Exit Sub
            Else
                WriteUTF8WithoutBOM terminalResult, fileDirtenant
            End If
    
    
    MsgBox "Files: Terminal.json is created. Please take a look...!!!"
End Sub
Sub GenerateTenantJsonFile_20240123()
    Dim seriNumberData As Long
    Dim seriNumber As String
    Dim tenantResult As String
    tenantResult = ""
    ' Define last colum and last row used
    
'2024.01.29 A列の最終行を取得するように変更
'    Dim intLastCellIndexInRow As Integer
'    intLastCellIndexInRow = ActiveCell.SpecialCells(xlLastCell).Column
'    Dim intLastCellIndexInColumn As Integer
'    intLastCellIndexInColumn = ActiveCell.SpecialCells(xlLastCell).Row
'    ' Get data from sheet
'    Dim tenantStart As Integer
'    tenantStart = 4
'
'    ' Get all S/N's json
'    tenantResult = tenantResult & "{" & vbNewLine & ""
'    For seriNumberData = 13 To intLastCellIndexInColumn
'        seriNumber = Cells(seriNumberData, 2).Value
'        Dim tenantRange As Range: Set tenantRange = Range(Cells(seriNumberData, tenantStart), Cells(seriNumberData, tenantStart + tenantDataCount - 1))
'
'        tenantResult = tenantResult & """" & Cells(seriNumberData, 2).Value & """" & ":"
'        tenantResult = tenantResult & GetTentantNewFormat(tenantRange)
'
'        If seriNumberData <> intLastCellIndexInColumn Then
        
        
    Dim xlLastRow As Long                               'Excel自体の最終行
    Dim LastRow As Long                                 '最終行
 
        xlLastRow = Cells(Rows.Count, 1).Row            'Excelの最終行を取得
        
        
    '    LastRow = Cells(xlLastRow, 4).End(xlUp).Row     'D列の最終行を取得
        LastRow = Cells(xlLastRow, 1).End(xlUp).Row     '2024.04.02 A列の最終行を取得
    
    
'2024.04.19　A列に1が立っているデータの最終行を取得
    For i = 9 To LastRow
        If Cells(i, 1) = 1 Then
            LastRow = Cells(i, 1).Row
    End If
    Next i

'2024.04.19 END
    
    
    
    ' Get data from sheet
    Dim tenantStart As Integer
    tenantStart = 4
       
    
'2024.04.02　実行前にA列に1が立っているデータがなければマクロを終了
    If WorksheetFunction.CountIf(Worksheets("Tenant").Range("A9:A" & LastRow), 1) = 0 Then
        MsgBox ("出力列に「1」の立っている行がないため、処理を終了します。")
        Exit Sub
    End If
    
    
    ' Get all S/N's json
    tenantResult = tenantResult & "{" & vbNewLine & ""
'    For seriNumberData = 13 To LastRow
    For seriNumberData = 9 To LastRow                  '開始位置を9行目に変更
    
    
'2024.04.02　A列に1が立っているデータを出力対象とする
    If Cells(seriNumberData, 1) = 1 Then
        
        
'2024.04.02　1列ずらす対応
'        seriNumber = Cells(seriNumberData, 2).Value
        seriNumber = Cells(seriNumberData, 3).Value
        Dim tenantRange As Range: Set tenantRange = Range(Cells(seriNumberData, tenantStart), Cells(seriNumberData, tenantStart + tenantDataCount - 1))
        
        tenantResult = tenantResult & """" & Cells(seriNumberData, 3).Value & """" & ":"
        tenantResult = tenantResult & GetTentantNewFormat(tenantRange)

        If seriNumberData <> LastRow Then
'2024.01.29 END
        
        
        
        
            tenantResult = tenantResult & ","
        End If
        tenantResult = tenantResult & vbNewLine & ""
        
        
'2024.04.02　A列に1が立っているデータを出力対象とし、立っていた「1」を「済」に変更
        Cells(seriNumberData, 1) = "済"
    End If
    
    
    Next
    tenantResult = tenantResult & "}"
    ' Export JSON


    
'キャンセルボタン押下時falseファイルができないように対応
'    Dim fileDirtenant As String: fileDirtenant = Application.GetSaveAsFilename("Tenant.json", fileFilter:="JSON Files (*.json), *.json")
'    WriteUTF8WithoutBOM tenantResult, fileDirtenant

    Dim fileDirtenant As String
        fileDirtenant = Application.GetSaveAsFilename("Tenant.json", fileFilter:="JSON Files (*.json), *.json")

            If fileDirtenant = "False" Then
              Exit Sub
            Else
                WriteUTF8WithoutBOM tenantResult, fileDirtenant
            End If

    MsgBox "Files: Tenant.json is created. Please take a look...!!!"
End Sub
Sub GenerateStoreJsonFile_20240123()
    Dim storeResult As String
    storeResult = ""
    ' Define last colum and last row used
'2024.01.29 未使用のためコメントアウト
'    Dim intLastCellIndexInRow As Integer
'    intLastCellIndexInRow = ActiveCell.SpecialCells(xlLastCell).Column
'    Dim intLastCellIndexInColumn As Integer
'    intLastCellIndexInColumn = ActiveCell.SpecialCells(xlLastCell).Row
'2024.01.29 未使用のためコメントアウト END

' Get data from sheet
    
    Dim storeStart As Integer
    storeStart = 5
    
    ' Get all S/N's json
'2024.01.30 対象列変更
'    Dim storeRange As Range: Set storeRange = Range("H6", "H79")
    Dim storeRange As Range: Set storeRange = Range("I6", "I79")
    
    storeResult = storeResult & GetStoreNewFormat(storeRange)

    ' Export JSON
        
'キャンセルボタン押下時falseファイルができないように対応
'    Dim fileDirtenant As String: fileDirtenant = Application.GetSaveAsFilename("Store.json", fileFilter:="JSON Files (*.json), *.json")
'    WriteUTF8WithoutBOM tenantResult, fileDirtenant

    Dim fileDirtenant As String
        fileDirtenant = Application.GetSaveAsFilename("Store.json", fileFilter:="JSON Files (*.json), *.json")

            If fileDirtenant = "False" Then
              Exit Sub
            Else
                WriteUTF8WithoutBOM storeResult, fileDirtenant
            End If
                
    MsgBox "Files: store.json is created. Please take a look...!!!"
End Sub
Sub GenerateCompanyJsonFile_20240123()
    Dim companyResult As String
    companyResult = ""
    ' Define last colum and last row used
'2024.01.29 未使用のためコメントアウト
'    Dim intLastCellIndexInRow As Integer
'    intLastCellIndexInRow = ActiveCell.SpecialCells(xlLastCell).Column
'    Dim intLastCellIndexInColumn As Integer
'    intLastCellIndexInColumn = ActiveCell.SpecialCells(xlLastCell).Row
'2024.01.29 未使用のためコメントアウト END
    ' Get data from sheet
    Dim companyStart As Integer
    companyStart = 5
    
    ' Get all S/N's json
    Dim companyRange As Range: Set companyRange = Range("H6", "H24")
    companyResult = companyResult & GetCompanyNewFormat(companyRange)

    ' Export JSON
        
'キャンセルボタン押下時falseファイルができないように対応
'    Dim fileDirtenant As String: fileDirtenant = Application.GetSaveAsFilename("Company.json", fileFilter:="JSON Files (*.json), *.json")
'    WriteUTF8WithoutBOM tenantResult, fileDirtenant

    Dim fileDirtenant As String
        fileDirtenant = Application.GetSaveAsFilename("Company.json", fileFilter:="JSON Files (*.json), *.json")

            If fileDirtenant = "False" Then
              Exit Sub
            Else
                WriteUTF8WithoutBOM companyResult, fileDirtenant
            End If
        
        
    MsgBox "Files: Company.json is created. Please take a look...!!!"
End Sub

'Function WriteJsonFileWithUtf8()
'    Dim jsonText As String
'    jsonText = "{ ""message"": ""Hello world"" }"
'
'    ' Write the JSON string to a file with UTF-8 encoding
'    WriteUtf8TextToFile jsonText, "C:\Users\duong\OneDrive\M痒 t??h\duongna.json"
'End Function

Function WriteUtf8TextToFile(text As String, filePath As String)
    Dim utf8Stream As Object
    Set utf8Stream = CreateObject("ADODB.Stream")
    
    ' Open the stream
    utf8Stream.Charset = "UTF-8"
    utf8Stream.Open
    utf8Stream.WriteText text

    ' Save the stream to the file
    utf8Stream.SaveToFile filePath, 2 ' 2 represents adSaveCreateOverWrite

    ' Close the stream
    utf8Stream.Close
End Function
Function WriteUTF8WithoutBOM(ByVal strText As String, fileName As String)
  Dim UTFStream As Object, BinaryStream As Object
  With CreateObject("adodb.stream")
     .Type = 2: .Mode = 3: .Charset = "UTF-8"
     .LineSeparator = -1
     .Open: .WriteText strText, 1
     .Position = 3 'skip BOM' !!!
     Set BinaryStream = CreateObject("adodb.stream")
         BinaryStream.Type = 1
         BinaryStream.Mode = 3
         BinaryStream.Open
        .CopyTo BinaryStream
        .Flush
    .Close
  End With
    BinaryStream.SaveToFile fileName, 2
    BinaryStream.Flush
    BinaryStream.Close
End Function
Public Function Format(ParamArray arr() As Variant) As String
    Dim i As Long
    Dim temp As String

    temp = CStr(arr(0))
    For i = 1 To UBound(arr)
        temp = Replace(temp, "{" & i - 1 & "}", CStr(arr(i)))
    Next

    Format = temp
End Function
Function CleanSpace(ByVal strIn As String) As String
    strIn = Trim(strIn)
    Do While InStr(strIn, "  ")
        strIn = Replace(strIn, "  ", " ")
    Loop
    CleanSpace = strIn
End Function
Public Function GetTerminalNewFormat(rangeInput As Range, snData As String) As String
'    Dim lang, result, reportType As String
    Dim lang, result, reportType, pass_1, pass_2, choice As String
    Dim numberOfipPort, formatPosition As Integer
    numberOfipPort = 2
    formatPosition = 1

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Language")

    pass_1 = ThisWorkbook.Sheets("password").Range("B1")
    pass_2 = ThisWorkbook.Sheets("password").Range("B2")


'2024.01.29 AutoFitする必要ない
'    ws.Columns("G:H").EntireColumn.AutoFit
'2024.01.29 AutoFitする必要ない END

    lang = "," & vbNewLine & "    ""lang"": {" & vbNewLine
        For countLang = 3 To 100
        If ws.Cells(countLang, 7).Value <> "" And ws.Cells(countLang + 1, 7).Value <> "" Then
            lang = lang & "        " & """" & ws.Cells(countLang, 8).Value & """: """ & ws.Cells(countLang, 7).Value & """" & "," & vbNewLine
        End If
        If ws.Cells(countLang, 7).Value <> "" And ws.Cells(countLang + 1, 7).Value = "" Then
            lang = lang & "        " & """" & ws.Cells(countLang, 8).Value & """: """ & ws.Cells(countLang, 7).Value & """" & vbNewLine
            lang = lang & "    }" & vbNewLine
        End If
    Next

'2024.01.29 表示順入替
'    result = "{" & vbNewLine & "    ""regno"": """ & rangeInput.Cells(1, 1).Value & """," & vbNewLine & "    ""pos_group_no"": """ & rangeInput.Cells(1, 2).Value & """," & vbNewLine & "    ""tid"": """ & rangeInput.Cells(1, 3).Value & """," & vbNewLine & "    ""network_setting"": {" & vbNewLine & "        ""owner_ip"": """ & rangeInput.Cells(1, 4).Value & """," & vbNewLine & "        ""subnet_mask"": """ & rangeInput.Cells(1, 5).Value & """," & vbNewLine & "        ""default_gateway"": """ & rangeInput.Cells(1, 6).Value & """" & vbNewLine & "    }," & vbNewLine & "    ""pos_link_setting"": {" & vbNewLine & "        ""connection"": " & rangeInput.Cells(1, 7).Value & "," & vbNewLine & "        ""manual_money_input"": " & rangeInput.Cells(1, 8).Value & "" & vbNewLine & "    }," & vbNewLine & "    ""settlement_inspection_control"": {" & vbNewLine
'    result = result & "        ""parent_child_setting"": " & rangeInput.Cells(1, 9).Value & "," & vbNewLine & "        ""mode"": """ & rangeInput.Cells(1, 10).Value & """" & vbNewLine & "    }," & vbNewLine & "    ""credit_parameter"": {" & vbNewLine & "        ""credit_tid"": """ & rangeInput.Cells(1, 11).Value & """" & vbNewLine & "    }," & vbNewLine & "    ""emoney_parameter"": {" & vbNewLine & "        ""receiptPrintCompany"": " & rangeInput.Cells(1, 12).Value & "," & vbNewLine & "        ""waonGivePoint"": " & rangeInput.Cells(1, 13).Value & "," & vbNewLine & "        ""utid"": """ & rangeInput.Cells(1, 14).Value & """," & vbNewLine & "        ""activate_id"": """ & rangeInput.Cells(1, 15).Value & """," & vbNewLine & "        ""activation_password"": """ & rangeInput.Cells(1, 16).Value & """" & vbNewLine & "    }," & vbNewLine
'    result = result & "    ""payment_limit_amount"": " & rangeInput.Cells(1, 17).Value & "," & vbNewLine & "    ""thank_title"": """ & rangeInput.Cells(1, 18).Value & """," & vbNewLine & "    ""thank_title_union"": """ & rangeInput.Cells(1, 19).Value & """," & vbNewLine & "    ""update_app"": {" & vbNewLine & "        ""update_enable"": " & rangeInput.Cells(1, 20).Value & "," & vbNewLine & "        ""time_start"": """ & rangeInput.Cells(1, 21).Value & """," & vbNewLine & "        ""time_end"": """ & rangeInput.Cells(1, 22).Value & """" & vbNewLine & "    }," & vbNewLine & "    ""menu_control"": {" & vbNewLine & "        ""normal_menu_password"": """ & rangeInput.Cells(1, 23).Value & """," & vbNewLine & "        ""maintenance_password"": """ & rangeInput.Cells(1, 24).Value & """," & vbNewLine & "        ""settings_password"": """ & rangeInput.Cells(1, 25).Value
'    result = result & """" & vbNewLine & "    }," & vbNewLine & "    ""application_tid"": {" & vbNewLine & "        ""qrcode_parameter_tid"": """ & rangeInput.Cells(1, 26).Value & """," & vbNewLine & "        ""aeon_gift_parameter_tid"": """ & rangeInput.Cells(1, 27).Value & """," & vbNewLine & "        ""waonpoint_parameter_tid"": """ & rangeInput.Cells(1, 28).Value & """," & vbNewLine & "        ""aeon_card_parameter_tid"": """ & rangeInput.Cells(1, 29).Value & """" & vbNewLine & "    }," & vbNewLine & "    ""input_item_pattern"": {" & vbNewLine & "      ""settlement_sales_input_item_pattern"": """ & rangeInput.Cells(1, 30).Value & """" & vbNewLine
'    result = result & "    }," & vbNewLine
'    result = result & "    ""launcher"":"
    
'2024.04.02 開始位置をD列に変更
    result = "{" & vbNewLine & "    ""regno"": """ & rangeInput.Cells(1, 3).Value & """," & vbNewLine
    result = result & "    ""pos_group_no"": """ & rangeInput.Cells(1, 4).Value & """," & vbNewLine
    result = result & "    ""tid"": """ & rangeInput.Cells(1, 5).Value & """," & vbNewLine
    result = result & "    ""network_setting"": {" & vbNewLine & "        ""owner_ip"": """ & rangeInput.Cells(1, 6).Value & """," & vbNewLine
    result = result & "        ""subnet_mask"": """ & rangeInput.Cells(1, 7).Value & """," & vbNewLine
    result = result & "        ""default_gateway"": """ & rangeInput.Cells(1, 8).Value & """" & vbNewLine
    result = result & "    }," & vbNewLine & "    ""pos_link_setting"": {" & vbNewLine
    
        
'2024.01.30 選択肢日本語化
        If rangeInput.Cells(1, 9).Value = "連動する" Then
            choice = "true"
          Else
            choice = "false"
        End If
    result = result & "        ""connection"": " & choice & "," & vbNewLine
'    result = result & "        ""connection"": " & rangeInput.Cells(1, 7).Value & "," & vbNewLine

        If rangeInput.Cells(1, 10).Value = "入力する" Then
            choice = "true"
          Else
            choice = "false"
        End If
    result = result & "        ""manual_money_input"": " & choice & "" & vbNewLine
'    result = result & "        ""manual_money_input"": " & rangeInput.Cells(1, 8).Value & "" & vbNewLine
    result = result & "    }," & vbNewLine & "    ""settlement_inspection_control"": {" & vbNewLine
    
        If rangeInput.Cells(1, 12).Value = "親子設定あり" Then
            choice = "true"
          Else
            choice = "false"
        End If

    result = result & "        ""parent_child_setting"": " & choice & "," & vbNewLine
'    result = result & "        ""parent_child_setting"": " & rangeInput.Cells(1, 10).Value & "," & vbNewLine

        If rangeInput.Cells(1, 13).Value = "親機" Then
            choice = "base"
          Else
            choice = "sub"
        End If

    result = result & "        ""mode"": """ & choice & """" & vbNewLine & "    }," & vbNewLine
'    result = result & "        ""mode"": """ & rangeInput.Cells(1, 11).Value & """" & vbNewLine & "    }," & vbNewLine
    result = result & "    ""credit_parameter"": {" & vbNewLine & "        ""credit_tid"": """ & rangeInput.Cells(1, 14).Value & """" & vbNewLine
    result = result & "    }," & vbNewLine
    result = result & "    ""emoney_parameter"": {" & vbNewLine
    
        If rangeInput.Cells(1, 19).Value = "印字する" Then
            choice = "true"
          Else
            choice = "false"
        End If
    
    result = result & "        ""receiptPrintCompany"": " & choice & "," & vbNewLine
'    result = result & "        ""receiptPrintCompany"": " & rangeInput.Cells(1, 17).Value & "," & vbNewLine

        If rangeInput.Cells(1, 20).Value = "付与する" Then
            choice = "true"
          Else
            choice = "false"
        End If

    result = result & "        ""waonGivePoint"": " & choice & "," & vbNewLine
'    result = result & "        ""waonGivePoint"": " & rangeInput.Cells(1, 18).Value & "," & vbNewLine
    result = result & "        ""utid"": """ & rangeInput.Cells(1, 21).Value & """," & vbNewLine
    result = result & "        ""activate_id"": """ & rangeInput.Cells(1, 22).Value & """," & vbNewLine
    result = result & "        ""activation_password"": """ & rangeInput.Cells(1, 23).Value & """" & vbNewLine & "    }," & vbNewLine
    result = result & "    ""payment_limit_amount"": " & rangeInput.Cells(1, 24).Value & "," & vbNewLine
    result = result & "    ""thank_title"": """ & rangeInput.Cells(1, 25).Value & """," & vbNewLine
    result = result & "    ""thank_title_union"": """ & rangeInput.Cells(1, 26).Value & """," & vbNewLine
    result = result & "    ""update_app"": {" & vbNewLine
        
        If rangeInput.Cells(1, 27).Value = "有効" Then
            choice = "true"
          Else
            choice = "false"
        End If

    result = result & "        ""update_enable"": " & choice & "," & vbNewLine
'    result = result & "        ""update_enable"": " & rangeInput.Cells(1, 25).Value & "," & vbNewLine
'2024.01.30 選択肢日本語化 END

'2024.01.30 時刻分離
    result = result & "        ""time_start"": """
    result = result & rangeInput.Cells(1, 28).Value
    result = result & rangeInput.Cells(1, 29).Value
    result = result & rangeInput.Cells(1, 30).Value & " "
    result = result & rangeInput.Cells(1, 31).Value & ":"
    result = result & rangeInput.Cells(1, 32).Value
    result = result & """," & vbNewLine
       
    result = result & "        ""time_end"": """
    result = result & rangeInput.Cells(1, 33).Value
    result = result & rangeInput.Cells(1, 34).Value
    result = result & rangeInput.Cells(1, 35).Value & " "
    result = result & rangeInput.Cells(1, 36).Value & ":"
    result = result & rangeInput.Cells(1, 37).Value
'2024.01.30 時刻分離 END

    
    result = result & """" & vbNewLine & "    }," & vbNewLine
    result = result & "    ""menu_control"": {" & vbNewLine & "        ""normal_menu_password"": """ & rangeInput.Cells(1, 38).Value
    
'    result = result & """," & vbNewLine & "        ""maintenance_password"": """ & rangeInput.Cells(1, 37).Value & """," & vbNewLine
'    result = result & "        ""settings_password"": """ & rangeInput.Cells(1, 38).Value
    
    result = result & """," & vbNewLine & "        ""maintenance_password"": """ & pass_1 & """," & vbNewLine
    result = result & "        ""settings_password"": """ & pass_2 & """" & vbNewLine
    result = result & "    }," & vbNewLine & "    ""application_tid"": {" & vbNewLine
    result = result & "        ""qrcode_parameter_tid"": """ & rangeInput.Cells(1, 15).Value & """," & vbNewLine
    result = result & "        ""aeon_gift_parameter_tid"": """ & rangeInput.Cells(1, 16).Value & """," & vbNewLine
    result = result & "        ""waonpoint_parameter_tid"": """ & rangeInput.Cells(1, 17).Value & """," & vbNewLine
    result = result & "        ""aeon_card_parameter_tid"": """ & rangeInput.Cells(1, 18).Value & """" & vbNewLine
    
    
'2024.02.02
'    result = result & "    }," & vbNewLine & "    ""input_item_pattern"": {" & vbNewLine
    result = result & "    }," & vbNewLine & "    ""settlement_inspect"": {" & vbNewLine
    result = result & "        ""input_item_pattern"": {" & vbNewLine
      
    result = result & "            ""settlement_sales_input_item_pattern"": """ & rangeInput.Cells(1, 11).Value & """" & vbNewLine
    result = result & "        }" & vbNewLine
    result = result & "    }," & vbNewLine
    
    result = result & "    ""launcher"":"
'    reportType = ThisWorkbook.Sheets("Terminal").Range("AF" & snData).Value
    reportType = ThisWorkbook.Sheets("Terminal").Range("M" & snData).Value
'2024.01.29 表示順入替    END


    If reportType = "1" And LCase(ThisWorkbook.Sheets("Tenant").Range("D" & snData).Value) = "ar" Then
    result = result & GetReportForm1(ThisWorkbook.Sheets("Report").Range("C24:C44"))
    ElseIf reportType = "2" And LCase(ThisWorkbook.Sheets("Tenant").Range("D" & snData).Value) = "ar" Then
    result = result & GetReportForm2(ThisWorkbook.Sheets("Report").Range("C48:C72"))
    ElseIf reportType = "3" And LCase(ThisWorkbook.Sheets("Tenant").Range("D" & snData).Value) = "ar" Then
    result = result & GetReportForm3(ThisWorkbook.Sheets("Report").Range("C76:C99"))
    ElseIf reportType = "1" And LCase(ThisWorkbook.Sheets("Tenant").Range("D" & snData).Value) = "at" Then
    result = result & GetReportForm4(ThisWorkbook.Sheets("Report").Range("C103:C104"))
    
'20240415 ATでもARテナント1を利用する対応
    ElseIf reportType = "2" And LCase(ThisWorkbook.Sheets("Tenant").Range("D" & snData).Value) = "at" Then
    result = result & GetReportForm1(ThisWorkbook.Sheets("Report").Range("C24:C44"))
 '20240415 ATでもARテナント1を利用する対応 end
    
    
    ElseIf LCase(ThisWorkbook.Sheets("Tenant").Range("D" & snData).Value) = "am" Then
    result = result & GetReportFormAM(ThisWorkbook.Sheets("Report").Range("C6:C20"))
    Else
    result = result & GetReportForm4(ThisWorkbook.Sheets("Report").Range("C103:C104"))
    End If

    result = result & lang & "}"

    GetTerminalNewFormat = result
End Function
Public Function GetTentantNewFormat(rangeInput As Range) As String
'    Dim lang, result As String
    Dim lang, result, choice As String
    Dim numberOfipPort, formatPosition As Integer
    numberOfipPort = 2
    formatPosition = 1

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Language")

'2024.01.29 AutoFitする必要ない
'    ws.Columns("E:F").EntireColumn.AutoFit
'2024.01.29 AutoFitする必要ない END

    lang = "," & vbNewLine & "    ""lang"": {" & vbNewLine
        For countLang = 3 To 100
        If ws.Cells(countLang, 5).Value <> "" And ws.Cells(countLang + 1, 5).Value <> "" Then
            lang = lang & "        " & """" & ws.Cells(countLang, 6).Value & """: """ & ws.Cells(countLang, 5).Value & """" & "," & vbNewLine
        End If
        If ws.Cells(countLang, 5).Value <> "" And ws.Cells(countLang + 1, 5).Value = "" Then
            lang = lang & "        " & """" & ws.Cells(countLang, 6).Value & """: """ & ws.Cells(countLang, 5).Value & """" & vbNewLine
            lang = lang & "    }" & vbNewLine
        End If
    Next


'2024.01.29 表示順入替
'    result = "{" & vbNewLine & "    ""tenant"": {" & vbNewLine & "        ""system"": """ & rangeInput.Cells(1, 1).Value & """," & vbNewLine & "        ""tenant_code"": """ & rangeInput.Cells(1, 2).Value & """" & vbNewLine & "    }," & vbNewLine & "    ""payment_control"": {" & vbNewLine & "        ""credit_setting"": {" & vbNewLine & "            ""credit_enable"": " & rangeInput.Cells(1, 3).Value & "," & vbNewLine & "            ""union_revenue_stamp_thresold"": " & rangeInput.Cells(1, 4).Value & "" & vbNewLine & "        }," & vbNewLine & "        ""union_enable"": " & rangeInput.Cells(1, 5).Value & "," & vbNewLine & "        ""emoney_enable"": " & rangeInput.Cells(1, 6).Value & "," & vbNewLine & "        ""aeon_gift_enable"": " & rangeInput.Cells(1, 7).Value & "," & vbNewLine & "        ""digital_setting"": {" & vbNewLine & "            ""digital_shopping_coupons_enable"": " & rangeInput.Cells(1, 8).Value
'    result = result & "," & vbNewLine & "            ""number_of_digital_shopping_coupons_available"": " & rangeInput.Cells(1, 9).Value & "" & vbNewLine & "        }," & vbNewLine & "        ""cash_enable"": " & rangeInput.Cells(1, 10).Value & "," & vbNewLine & "        ""qr_code_enable"": " & rangeInput.Cells(1, 11).Value & "," & vbNewLine & "        ""waonpoint_enable"": " & rangeInput.Cells(1, 12).Value & "" & vbNewLine & "    }," & vbNewLine & "    ""settlement_inspection_control"": {" & vbNewLine & "        ""duty_free_shop_control"": " & rangeInput.Cells(1, 13).Value & "," & vbNewLine & "        ""zero_setting"": " & rangeInput.Cells(1, 14).Value & "" & vbNewLine & "    }," & vbNewLine & "    ""tenant_name"": """ & rangeInput.Cells(1, 15).Value & """," & vbNewLine & "    ""signature_note"": {" & vbNewLine & "        ""signature_note_without_password"": """ & rangeInput.Cells(1, 16).Value & """," & vbNewLine
'    result = result & "        ""signature_note_with_password"": """ & rangeInput.Cells(1, 17).Value & """" & vbNewLine & "    }," & vbNewLine & "    ""qrpay"": {" & vbNewLine & "        ""qr_pay_card"": """ & rangeInput.Cells(1, 18).Value & """," & vbNewLine & "        ""qr_pay_charge"": """ & rangeInput.Cells(1, 19).Value & """" & vbNewLine & "    }" & vbNewLine
'    result = result & lang & "}"

'2024.04.02 開始位置をD列からに変更

    result = "{" & vbNewLine & "    ""tenant"": {" & vbNewLine & "        ""system"": """ & rangeInput.Cells(1, 2).Value & """," & vbNewLine
    result = result & "        ""tenant_code"": """ & rangeInput.Cells(1, 3).Value & """" & "," & vbNewLine
    result = result & "        ""tenant_name"": """ & rangeInput.Cells(1, 4).Value & """" & "," & vbNewLine
    result = result & "        ""tenant_name_print_1st"": """ & rangeInput.Cells(1, 5).Value & """" & "," & vbNewLine
    result = result & "        ""tenant_name_print_2nd"": """ & rangeInput.Cells(1, 6).Value & """" & vbNewLine
    result = result & "    }," & vbNewLine
    result = result & "    ""payment_control"": {" & vbNewLine & "        ""credit_setting"": {" & vbNewLine
    
'2024.01.30 選択肢日本語
        If rangeInput.Cells(1, 7).Value = "使用する" Then
            choice = "true"
          Else
            choice = "false"
        End If

    result = result & "            ""credit_enable"": " & choice & "," & vbNewLine
'    result = result & "            ""credit_enable"": " & rangeInput.Cells(1, 4).Value & "," & vbNewLine
    
    
    result = result & "            ""union_revenue_stamp_thresold"": " & rangeInput.Cells(1, 9).Value & "" & vbNewLine
    result = result & "        }," & vbNewLine
        
        If rangeInput.Cells(1, 8).Value = "使用する" Then
            choice = "true"
          Else
            choice = "false"
        End If
    result = result & "        ""union_enable"": " & choice & "," & vbNewLine
'    result = result & "        ""union_enable"": " & rangeInput.Cells(1, 5).Value & "," & vbNewLine
    
        If rangeInput.Cells(1, 10).Value = "使用する" Then
            choice = "true"
          Else
            choice = "false"
        End If
    result = result & "        ""emoney_enable"": " & choice & "," & vbNewLine
'    result = result & "        ""emoney_enable"": " & rangeInput.Cells(1, 7).Value & "," & vbNewLine
        
        If rangeInput.Cells(1, 11).Value = "使用する" Then
            choice = "true"
          Else
            choice = "false"
        End If
    result = result & "        ""aeon_gift_enable"": " & choice & "," & vbNewLine
'    result = result & "        ""aeon_gift_enable"": " & rangeInput.Cells(1, 8).Value & "," & vbNewLine
    
    result = result & "        ""digital_setting"": {" & vbNewLine
        
        If rangeInput.Cells(1, 12).Value = "使用する" Then
            choice = "true"
          Else
            choice = "false"
        End If
    result = result & "            ""digital_shopping_coupons_enable"": " & choice & "," & vbNewLine
'    result = result & "            ""digital_shopping_coupons_enable"": " & rangeInput.Cells(1, 9).Value & "," & vbNewLine
    
    result = result & "            ""number_of_digital_shopping_coupons_available"": " & rangeInput.Cells(1, 13).Value & "" & vbNewLine
    result = result & "        }," & vbNewLine
    
        If rangeInput.Cells(1, 14).Value = "使用する" Then
            choice = "true"
          Else
            choice = "false"
        End If
    result = result & "        ""cash_enable"": " & choice & "," & vbNewLine
'    result = result & "        ""cash_enable"": " & rangeInput.Cells(1, 11).Value & "," & vbNewLine
    
        If rangeInput.Cells(1, 15).Value = "使用する" Then
            choice = "true"
          Else
            choice = "false"
        End If
    result = result & "        ""qr_code_enable"": " & choice & "," & vbNewLine
'    result = result & "        ""qr_code_enable"": " & rangeInput.Cells(1, 12).Value & "," & vbNewLine
    
        If rangeInput.Cells(1, 16).Value = "使用する" Then
            choice = "true"
          Else
            choice = "false"
        End If
    result = result & "        ""waonpoint_enable"": " & choice & "" & vbNewLine
'    result = result & "        ""waonpoint_enable"": " & rangeInput.Cells(1, 13).Value & "" & vbNewLine
    
    result = result & "    }," & vbNewLine & "    ""settlement_inspection_control"": {" & vbNewLine
    
    
        If rangeInput.Cells(1, 17).Value = "使用する" Then
            choice = "true"
          Else
            choice = "false"
        End If
    result = result & "        ""duty_free_shop_control"": " & choice & "," & vbNewLine
'    result = result & "        ""duty_free_shop_control"": " & rangeInput.Cells(1, 14).Value & "," & vbNewLine

        If rangeInput.Cells(1, 18).Value = "使用する" Then
            choice = "true"
          Else
            choice = "false"
        End If
    result = result & "        ""zero_setting"": " & choice & "" & vbNewLine
'    result = result & "        ""zero_setting"": " & rangeInput.Cells(1, 15).Value & "" & vbNewLine
'2024.01.30 選択肢日本語 END

    result = result & "    }," & vbNewLine
    result = result & "    ""signature_note"": {" & vbNewLine
    result = result & "        ""signature_note_without_password"": """ & rangeInput.Cells(1, 19).Value & """," & vbNewLine
    result = result & "        ""signature_note_with_password"": """ & rangeInput.Cells(1, 20).Value & """" & vbNewLine
    result = result & "    }," & vbNewLine & "    ""qrpay"": {" & vbNewLine
    result = result & "        ""qr_pay_card"": """ & rangeInput.Cells(1, 21).Value & """," & vbNewLine
    result = result & "        ""qr_pay_charge"": """ & rangeInput.Cells(1, 22).Value & """" & vbNewLine & "    }" & vbNewLine
    result = result & lang & "}"
'2024.01.29 表示順入替 END



    GetTentantNewFormat = result
End Function
Public Function GetStoreNewFormat(rangeInput As Range) As String
'    Dim lang, result As String
    Dim lang, result, choice As String
    Dim numberOfipPort, formatPosition As Integer
    numberOfipPort = 2
    formatPosition = 1

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Language")
'2024.01.29 AutoFitする必要ない
'    ws.Columns("C:D").EntireColumn.AutoFit
'2024.01.29 AutoFitする必要ない END

lang = "," & vbNewLine & "    ""lang"": {" & vbNewLine
        For countLang = 3 To 100
        If ws.Cells(countLang, 3).Value <> "" And ws.Cells(countLang + 1, 3).Value <> "" Then
            lang = lang & "        " & """" & ws.Cells(countLang, 4).Value & """: """ & ws.Cells(countLang, 3).Value & """" & "," & vbNewLine
        End If
        If ws.Cells(countLang, 3).Value <> "" And ws.Cells(countLang + 1, 3).Value = "" Then
            lang = lang & "        " & """" & ws.Cells(countLang, 4).Value & """: """ & ws.Cells(countLang, 3).Value & """" & vbNewLine
            lang = lang & "    }" & vbNewLine
        End If
    Next

    result = "{" & vbNewLine & "    ""connection_destination"": {" & vbNewLine & "        ""credit_sever_primary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(1, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(2, 1).Value & """" & vbNewLine & "        }," & vbNewLine
    result = result & "        ""credit_sever_second"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(3, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(4, 1).Value & """" & vbNewLine & "        }," & vbNewLine
    result = result & "        ""credit_sever_tertiary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(5, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(6, 1).Value & """" & vbNewLine & "        }," & vbNewLine
    result = result & "        ""qrcode_sever_primary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(7, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(8, 1).Value & """" & vbNewLine & "        }," & vbNewLine
    result = result & "        ""qrcode_sever_second"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(9, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(10, 1).Value & """" & vbNewLine & "        }," & vbNewLine
    result = result & "        ""qrcode_sever_tertiary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(11, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(12, 1).Value & """" & vbNewLine & "        }," & vbNewLine
    result = result & "        ""wp_sever_primary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(13, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(14, 1).Value & """" & vbNewLine & "        }," & vbNewLine
    result = result & "        ""wp_sever_second"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(15, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(16, 1).Value & """" & vbNewLine & "        }," & vbNewLine
    result = result & "        ""wp_sever_tertiary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(17, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(18, 1).Value & """" & vbNewLine
    result = result & "        }," & vbNewLine
    
    
'電マネ口座伝として有効とする
    result = result & "        ""emoney_sever_primary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(19, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(20, 1).Value & """" & vbNewLine
    result = result & "        }," & vbNewLine
    result = result & "        ""emoney_sever_second"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(21, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(22, 1).Value & """" & vbNewLine
    result = result & "        }," & vbNewLine
    result = result & "        ""emoney_sever_tertiary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(23, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(24, 1).Value & """" & vbNewLine
    
'電マネ口座伝として有効とする END
    
'電マネ口座伝として有効にともないcolumn変更
    result = result & "        }," & vbNewLine
    result = result & "        ""union_sever_primary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(25, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(26, 1).Value & """" & vbNewLine
    result = result & "        }," & vbNewLine
    result = result & "        ""union_sever_second"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(27, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(28, 1).Value & """" & vbNewLine
    result = result & "        }," & vbNewLine
    result = result & "        ""union_sever_tertiary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(29, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(30, 1).Value & """" & vbNewLine
    result = result & "        }," & vbNewLine & "        ""aeon_gift_sever_primary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(31, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(32, 1).Value & """" & vbNewLine
    result = result & "        }," & vbNewLine
    result = result & "        ""aeon_gift_sever_second"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(33, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(34, 1).Value & """" & vbNewLine
    result = result & "        }," & vbNewLine
    result = result & "        ""aeon_gift_sever_tertiary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(35, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(36, 1).Value & """" & vbNewLine
    result = result & "        }," & vbNewLine
    result = result & "        ""digital_gift_sever_primary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(37, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(38, 1).Value & """" & vbNewLine
    result = result & "        }," & vbNewLine
    result = result & "        ""digital_gift_sever_second"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(39, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(40, 1).Value & """" & vbNewLine
    result = result & "        }," & vbNewLine
    result = result & "        ""digital_gift_sever_tertiary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(41, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(42, 1).Value & """" & vbNewLine
    result = result & "        }," & vbNewLine
    result = result & "        ""inspect_sever_primary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(43, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(44, 1).Value & """" & vbNewLine
    result = result & "        }," & vbNewLine
    result = result & "        ""inspect_sever_second"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(45, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(46, 1).Value & """" & vbNewLine
    result = result & "        }," & vbNewLine & "        ""inspect_sever_tertiary"": {" & vbNewLine
    result = result & "            ""IP"": """ & rangeInput.Cells(47, 1).Value & """," & vbNewLine
    result = result & "            ""port"": """ & rangeInput.Cells(48, 1).Value & """" & vbNewLine
    result = result & "        }" & vbNewLine
    result = result & "    }," & vbNewLine & "    ""settlement_inspect"": {" & vbNewLine
    result = result & "        ""payment_receipt_control"": {" & vbNewLine
    result = result & "            ""number_of_reports_issued_at_checkout"": """ & rangeInput.Cells(49, 1).Value & """" & vbNewLine
    result = result & "        }" & vbNewLine
    result = result & "    }," & vbNewLine & "    ""parking_lot_cooperation"": {" & vbNewLine
    
    
'2024.01.30 選択肢日本語
        If rangeInput.Cells(50, 1).Value = "使用する" Then
            choice = "true"
          Else
            choice = "false"
        End If
    
    result = result & "        ""parking_discount"": """ & choice & """" & vbNewLine
'    result = result & "        ""parking_discount"": """ & rangeInput.Cells(50, 1).Value & """" & vbNewLine
    
    
    result = result & "    }," & vbNewLine & "    ""umie"": {" & vbNewLine
    
        If rangeInput.Cells(51, 1).Value = "umie" Then
            choice = "umie"
          Else
            choice = "other"
        End If

    result = result & "        ""umie_compatibility"": """ & choice & """" & vbNewLine
'   result = result & "        ""umie_compatibility"": """ & rangeInput.Cells(51, 1).Value & """" & vbNewLine
    
    
    result = result & "    }," & vbNewLine & "    ""mozo"": {" & vbNewLine
    
        If rangeInput.Cells(52, 1).Value = "MOZO" Then
            choice = "mozo"
          Else
            choice = "other"
        End If
    
    result = result & "        ""mozo_compatibility"": """ & choice & """" & vbNewLine
'    result = result & "        ""mozo_compatibility"": """ & rangeInput.Cells(52, 1).Value & """" & vbNewLine
'2024.01.30 選択肢日本語 END
    
    result = result & "    }," & vbNewLine & "    ""store_code"": {" & vbNewLine
    result = result & "        ""store_code_details"": """ & rangeInput.Cells(53, 1).Value & """" & vbNewLine
    result = result & "    }," & vbNewLine & "    ""store_name"": {" & vbNewLine
    result = result & "        ""store_name_details"": """ & rangeInput.Cells(54, 1).Value & """" & vbNewLine & "    }" & vbNewLine
    result = result & lang & "}"

    GetStoreNewFormat = result
End Function
Public Function GetCompanyNewFormat(rangeInput As Range) As String
'    Dim lang, result As String
    Dim lang, result, choice As String
    Dim numberOfipPort, formatPosition, countLang As Integer
    numberOfipPort = 2
    formatPosition = 1

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Language")
'2024.01.29 AutoFitする必要ない
'    ws.Columns("A:B").EntireColumn.AutoFit
'2024.01.29 AutoFitする必要ない END

lang = "," & vbNewLine & "    ""lang"": {" & vbNewLine
        For countLang = 3 To 100
        If ws.Cells(countLang, 1).Value <> "" And ws.Cells(countLang + 1, 1).Value <> "" Then
            lang = lang & "        " & """" & ws.Cells(countLang, 2).Value & """: """ & ws.Cells(countLang, 1).Value & """" & "," & vbNewLine
        End If
        If ws.Cells(countLang, 1).Value <> "" And ws.Cells(countLang + 1, 1).Value = "" Then
            lang = lang & "        " & """" & ws.Cells(countLang, 2).Value & """: """ & ws.Cells(countLang, 1).Value & """" & vbNewLine
            lang = lang & "    }" & vbNewLine
        End If
    Next

    ' result = "{" & vbNewLine & "   ""connection_destination_settings"": {" & vbNewLine & "       ""DNS_server_1st"": """ & rangeInput.Cells(1, 1).Value & """," & vbNewLine & "       ""DNS_server_2nd"": """ & rangeInput.Cells(2, 1).Value & """," & vbNewLine & "       ""AFS_payment_getway"": {" & vbNewLine & "       ""IP"": """ & rangeInput.Cells(3, 1).Value & """," & vbNewLine & "       ""port"": """ & rangeInput.Cells(4, 1).Value & """" & vbNewLine & "       }" & vbNewLine & "   }," & vbNewLine & "   ""document_storage_func"": {" & vbNewLine & "       ""store_slip"": " & rangeInput.Cells(5, 1).Value & "," & vbNewLine & "       ""connection_destination"": """ & rangeInput.Cells(6, 1).Value & """," & vbNewLine & "       ""port_number"": """ & rangeInput.Cells(7, 1).Value & """," & vbNewLine & "       ""port_connection_timer"": """ & rangeInput.Cells(8, 1).Value & """," & vbNewLine & "       ""receive_send_timer"": """ & rangeInput.Cells(9, 1).Value & ""","
    ' result = result & vbNewLine & "       ""service_classification"": """ & rangeInput.Cells(10, 1).Value & """," & vbNewLine & "       ""tenant_code"": """ & rangeInput.Cells(11, 1).Value & """" & vbNewLine & "   }," & vbNewLine & "   ""shutdown_confirmation"": {" & vbNewLine & "       ""shutdown_with_popup_confirmation"": " & rangeInput.Cells(12, 1).Value & "" & vbNewLine & "   }," & vbNewLine & "   ""automatic_shutdown_avaiable"": {" & vbNewLine & "       ""automatic_shutdown_after_payment"": " & rangeInput.Cells(13, 1).Value & "" & vbNewLine & "   }," & vbNewLine & "   ""points_parameters"": {" & vbNewLine & "       ""terminal_setting_information"": {" & vbNewLine & "       ""company_code"": """ & rangeInput.Cells(14, 1).Value & """" & vbNewLine & "       ""requester_channel_id_code"": """ & rangeInput.Cells(15, 1).Value & """" & vbNewLine & "       }" & vbNewLine & "   }," & vbNewLine & "   ""receipt_message"": {"
    ' result = result & vbNewLine & "       ""receipt_message_setting"":  """ & rangeInput.Cells(16, 1).Value & """" & vbNewLine & "   }," & vbNewLine & "   ""payment_control"": {" & vbNewLine & "       ""credit_control"": {" & vbNewLine & "       ""specifiable_bonus_month_summer"": " & rangeInput.Cells(17, 1).Value & "," & vbNewLine & "       ""specifiable_bonus_month_winter"": " & rangeInput.Cells(18, 1).Value & "" & vbNewLine & "       ""credit_membership_mask_character"": """ & rangeInput.Cells(19, 1).Value & """," & vbNewLine & "       ""split_bonus_amout_threshold"": " & rangeInput.Cells(20, 1).Value & "" & vbNewLine & "       }" & vbNewLine & "   }"
    ' result = result & lang & "}"

    result = "{" & vbNewLine & "  ""connection_destination_settings"": {" & vbNewLine & "    ""DNS_server_1st"": """ & rangeInput.Cells(1, 1).Value & """," & vbNewLine & "    ""DNS_server_2nd"": """ & rangeInput.Cells(2, 1).Value & """," & vbNewLine & "    ""AFS_payment_getway"": {" & vbNewLine & "      ""IP"": """ & rangeInput.Cells(3, 1).Value & """," & vbNewLine & "      ""port"": """ & rangeInput.Cells(4, 1).Value & """" & vbNewLine & "    }" & vbNewLine & "  }," & vbNewLine & "  ""document_storage_func"": {" & vbNewLine
    
    
'2024.01.30 選択肢日本語
        If rangeInput.Cells(5, 1).Value = "使用する" Then
            choice = "true"
          Else
            choice = "false"
        End If
    result = result & "    ""store_slip"": " & choice & "," & vbNewLine
'    result = result & "    ""store_slip"": " & rangeInput.Cells(5, 1).Value & "," & vbNewLine
        
    result = result & "    ""connection_destination"": """ & rangeInput.Cells(6, 1).Value & """," & vbNewLine & "    ""port_number"": """ & rangeInput.Cells(7, 1).Value & """," & vbNewLine & "    ""port_connection_timer"": """ & rangeInput.Cells(8, 1).Value & """," & vbNewLine & "    ""receive_send_timer"": """
    result = result & rangeInput.Cells(9, 1).Value & """," & vbNewLine & "    ""service_classification"": """ & rangeInput.Cells(10, 1).Value & """," & vbNewLine & "    ""tenant_code"": """ & rangeInput.Cells(11, 1).Value & """" & vbNewLine & "  }," & vbNewLine & "  ""shutdown_confirmation"": {" & vbNewLine
    
    
        If rangeInput.Cells(12, 1).Value = "表示あり" Then
            choice = "true"
          Else
            choice = "false"
        End If
    result = result & "    ""shutdown_with_popup_confirmation"": " & choice & "  }," & vbNewLine
'    result = result & "    ""shutdown_with_popup_confirmation"": " & rangeInput.Cells(12, 1).Value & "" & vbNewLine & "  }," & vbNewLine
    
    
    result = result & "  ""automatic_shutdown_avaiable"": {" & vbNewLine
    
        If rangeInput.Cells(13, 1).Value = "シャットダウンあり" Then
            choice = "true"
          Else
            choice = "false"
        End If
    result = result & "    ""automatic_shutdown_after_payment"": " & choice & "" & vbNewLine
'    result = result & "    ""automatic_shutdown_after_payment"": " & rangeInput.Cells(13, 1).Value & "" & vbNewLine
'2024.01.30 選択肢日本語 END
    
    result = result & "  }," & vbNewLine & "  ""points_parameters"": {" & vbNewLine & "    ""terminal_setting_information"": {" & vbNewLine & "      ""company_code"": """ & rangeInput.Cells(14, 1).Value & """" & vbNewLine & "    }" & vbNewLine & "  }," & vbNewLine & "  ""payment_control"": {" & vbNewLine & "    ""credit_control"": {"
    result = result & vbNewLine & "      ""specifiable_bonus_month_summer"": " & rangeInput.Cells(15, 1).Value & "," & vbNewLine & "      ""specifiable_bonus_month_winter"": " & rangeInput.Cells(16, 1).Value & "" & vbNewLine & "    }" & vbNewLine & "  }" & vbNewLine
    result = result & lang & "}"

    GetCompanyNewFormat = result
End Function
Public Function GetReportForm1(rangeInput As Range) As String
    Dim i As Long
    Dim mask, header_1, header_2, unit_1, unit_2, unit_3 As String
    header_1 = ThisWorkbook.Sheets("Report").Range("D108")
    header_2 = ThisWorkbook.Sheets("Report").Range("D109")
    unit_1 = ThisWorkbook.Sheets("Report").Range("D110")
    unit_2 = ThisWorkbook.Sheets("Report").Range("D111")
    unit_3 = ThisWorkbook.Sheets("Report").Range("D112")
    
    mask = "    [" & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""header_1""," & vbNewLine & "     ""title"": """ & header_1 & """," & vbNewLine & "     ""type"": ""automatic""," & vbNewLine & "     ""func"": []," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 1," & vbNewLine & "     ""field_query"": null" & vbNewLine & "     }," & vbNewLine & "     {" & vbNewLine & "     ""id"": ""header_2""," & vbNewLine & "     ""title"": """ & header_2 & """," & vbNewLine & "     ""type"": ""automatic""," & vbNewLine & "     ""func"": []," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 2," & vbNewLine & "     ""field_query"": null" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id1""," & vbNewLine & "     ""title"": ""{0}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F24:G24")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(1, 3)) & "]," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 3," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(1, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id2""," & vbNewLine & "     ""title"": ""{1}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F25:G25")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(2, 3)) & "]," & vbNewLine & "     ""unit"": {" & vbNewLine & "         ""title"": """ & unit_2 & """," & vbNewLine & "         ""type"": ""count""" & vbNewLine & "     }," & vbNewLine & "     ""order"": 4," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(2, 1), 2) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id3""," & vbNewLine & "     ""title"": ""{2}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F26:G26")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(3, 3)) & "]," & vbNewLine & "     ""unit"": {" & vbNewLine & "         ""title"": """ & unit_1 & """," & vbNewLine & "         ""type"": ""count""" & vbNewLine & "     }," & vbNewLine & "     ""order"": 5," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(3, 1), 2) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id4""," & vbNewLine & "     ""title"": ""{3}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F27:G27")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(4, 3)) & "]," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 6," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(4, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id5""," & vbNewLine & "     ""title"": ""{4}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F28:G28")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(5, 3)) & "]," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 7," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(5, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id6""," & vbNewLine & "     ""title"": ""{5}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F29:G29")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(6, 3)) & "]," & vbNewLine & "     ""unit"": {" & vbNewLine & "         ""title"": """"," & vbNewLine & "         ""type"": ""count""" & vbNewLine & "     }," & vbNewLine & "     ""order"": 8," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(6, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id7""," & vbNewLine & "     ""title"": ""{6}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F30:G30")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(7, 3)) & "]," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 9," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(7, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id8""," & vbNewLine & "     ""title"": ""{7}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F31:G31")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(8, 3)) & "]," & vbNewLine & "     ""unit"": {" & vbNewLine & "         ""title"": """"," & vbNewLine & "         ""type"": ""count""" & vbNewLine & "     }," & vbNewLine & "     ""order"": 10," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(8, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id9""," & vbNewLine & "     ""title"": ""{8}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F32:G32")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(9, 3)) & "]," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 11," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(9, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id10""," & vbNewLine & "     ""title"": ""{9}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F33:G33")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(10, 3)) & "]," & vbNewLine & "     ""unit"": {" & vbNewLine & "         ""title"": """"," & vbNewLine & "         ""type"": ""count""" & vbNewLine & "     }," & vbNewLine & "     ""order"": 12," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(10, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id11""," & vbNewLine & "     ""title"": ""{10}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F34:G34")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(11, 3)) & "]," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 13," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(11, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id12""," & vbNewLine & "     ""title"": ""{11}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F35:G35")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(12, 3)) & "]," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 14," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(12, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id13""," & vbNewLine & "     ""title"": ""{12}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F36:G36")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(13, 3)) & "]," & vbNewLine & "     ""unit"": {" & vbNewLine & "         ""title"": ""PT""," & vbNewLine & "         ""type"": ""point""" & vbNewLine & "     }," & vbNewLine & "     ""order"": 15," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(13, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id14""," & vbNewLine & "     ""title"": ""{13}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F37:G37")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(14, 3)) & "]," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 16," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(14, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id15""," & vbNewLine & "     ""title"": ""{14}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F38:G38")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(15, 3)) & "]," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 17," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(15, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id16""," & vbNewLine & "     ""title"": ""{15}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F39:G39")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(16, 3)) & "]," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 18," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(16, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id17""," & vbNewLine & "     ""title"": ""{16}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F40:G40")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(17, 3)) & "]," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 19," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(17, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id18""," & vbNewLine & "     ""title"": ""{17}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F41:G41")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(18, 3)) & "]," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 20," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(18, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id19""," & vbNewLine & "     ""title"": ""{18}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F42:G42")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(19, 3)) & "]," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 21," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(19, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id20""," & vbNewLine & "     ""title"": ""{19}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F43:G43")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(20, 3)) & "]," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 22," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(20, 1), 1) & "" & vbNewLine & "     }," & vbNewLine
    mask = mask & "     {" & vbNewLine & "     ""id"": ""id21""," & vbNewLine & "     ""title"": ""{20}""," & vbNewLine & "     ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F44:G44")) & "," & vbNewLine & "     ""func"": [" & GetCalculationFormArray(rangeInput.Cells(21, 3)) & "]," & vbNewLine & "     ""unit"": null," & vbNewLine & "     ""order"": 23," & vbNewLine & "     ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(21, 1), 1) & "" & vbNewLine & "     }" & vbNewLine & " ]"

    For i = 0 To rangeInput.Count
        mask = Replace$(mask, "{" & i & "}", rangeInput.Cells(i + 1, 1).Value)
    Next
    GetReportForm1 = mask
End Function

Public Function GetReportForm2(rangeInput As Range) As String
    Dim i As Long
    Dim mask, header_1, header_2, unit_1, unit_2, unit_3 As String
    header_1 = ThisWorkbook.Sheets("Report").Range("D108")
    header_2 = ThisWorkbook.Sheets("Report").Range("D109")
    unit_1 = ThisWorkbook.Sheets("Report").Range("D110")
    unit_2 = ThisWorkbook.Sheets("Report").Range("D111")
    unit_3 = ThisWorkbook.Sheets("Report").Range("D112")
    
    mask = "  [" & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""header_1""," & vbNewLine & "      ""title"": """ & header_1 & """," & vbNewLine & "      ""type"": ""automatic""," & vbNewLine & "      ""func"": []," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 1," & vbNewLine & "      ""field_query"": null" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""header_2""," & vbNewLine & "      ""title"": """ & header_2 & """," & vbNewLine & "      ""type"": ""automatic""," & vbNewLine & "      ""func"": []," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 2," & vbNewLine & "      ""field_query"": null" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id1""," & vbNewLine & "      ""title"": ""{0}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F48:G48")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(1, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 3," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(1, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id2""," & vbNewLine & "      ""title"": ""{1}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F49:G49")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(2, 3)) & "]," & vbNewLine & "      ""unit"": {" & vbNewLine & "        ""title"": """ & unit_2 & """," & vbNewLine & "        ""type"": ""count""" & vbNewLine & "      }," & vbNewLine & "      ""order"": 4," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(2, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id3""," & vbNewLine & "      ""title"": ""{2}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F50:G50")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(3, 3)) & "]," & vbNewLine & "      ""unit"": {" & vbNewLine & "        ""title"": """ & unit_1 & """," & vbNewLine & "        ""type"": ""count""" & vbNewLine & "      }," & vbNewLine & "      ""order"": 5," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(3, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id4""," & vbNewLine & "      ""title"": ""{3}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F51:G51")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(4, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 6," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(4, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id5""," & vbNewLine & "      ""title"": ""{4}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F52:G52")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(5, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 7," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(5, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id6""," & vbNewLine & "      ""title"": ""{5}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F53:G53")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(6, 3)) & "]," & vbNewLine & "      ""unit"": {" & vbNewLine & "        ""title"": """"," & vbNewLine & "        ""type"": ""count""" & vbNewLine & "      }," & vbNewLine & "      ""order"": 8," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(6, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id7""," & vbNewLine & "      ""title"": ""{6}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F54:G54")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(7, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 9," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(7, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id8""," & vbNewLine & "      ""title"": ""{7}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F55:G55")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(8, 3)) & "]," & vbNewLine & "      ""unit"": {" & vbNewLine & "        ""title"": """"," & vbNewLine & "        ""type"": ""count""" & vbNewLine & "      }," & vbNewLine & "      ""order"": 10," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(8, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id9""," & vbNewLine & "      ""title"": ""{8}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F56:G56")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(9, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 11," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(9, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id10""," & vbNewLine & "      ""title"": ""{9}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F57:G57")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(10, 3)) & "]," & vbNewLine & "      ""unit"": {" & vbNewLine & "        ""title"": """"," & vbNewLine & "        ""type"": ""count""" & vbNewLine & "      }," & vbNewLine & "      ""order"": 12," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(10, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id11""," & vbNewLine & "      ""title"": ""{10}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F58:G58")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(11, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 13," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(11, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id12""," & vbNewLine & "      ""title"": ""{11}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F59:G59")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(12, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 14," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(12, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id13""," & vbNewLine & "      ""title"": ""{12}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F60:G60")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(13, 3)) & "]," & vbNewLine & "      ""unit"": {" & vbNewLine & "        ""title"": ""PT""," & vbNewLine & "        ""type"": ""point""" & vbNewLine & "      }," & vbNewLine & "      ""order"": 15," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(13, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id14""," & vbNewLine & "      ""title"": ""{13}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F61:G61")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(14, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 16," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(14, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id15""," & vbNewLine & "      ""title"": ""{14}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F62:G62")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(15, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 17," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(15, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id16""," & vbNewLine & "      ""title"": ""{15}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F63:G63")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(16, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 18," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(16, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id17""," & vbNewLine & "      ""title"": ""{16}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F64:G64")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(17, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 19," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(17, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id18""," & vbNewLine & "      ""title"": ""{17}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F65:G65")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(18, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 20," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(18, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id19""," & vbNewLine & "      ""title"": ""{18}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F66:G66")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(19, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 21," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(19, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id20""," & vbNewLine & "      ""title"": ""{19}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F67:G67")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(20, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 22," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(20, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id21""," & vbNewLine & "      ""title"": ""{20}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F68:G68")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(21, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 23," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(21, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id22""," & vbNewLine & "      ""title"": ""{21}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F69:G69")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(22, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 24," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(22, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id23""," & vbNewLine & "      ""title"": ""{22}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F70:G70")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(23, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 25," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(23, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id24""," & vbNewLine & "      ""title"": ""{23}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F71:G71")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(24, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 26," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(24, 1), 2) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id25""," & vbNewLine & "      ""title"": ""{24}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F72:G72")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(25, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 27," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(25, 1), 2) & "" & vbNewLine
    mask = mask & "    }" & vbNewLine & "  ]"

    For i = 0 To rangeInput.Count
        mask = Replace$(mask, "{" & i & "}", rangeInput.Cells(i + 1, 1).Value)
    Next
    GetReportForm2 = mask
End Function
Public Function GetReportForm3(rangeInput As Range) As String
    Dim i As Long
    Dim mask, header_1, header_2, unit_1, unit_2, unit_3 As String
    header_1 = ThisWorkbook.Sheets("Report").Range("D108")
    header_2 = ThisWorkbook.Sheets("Report").Range("D109")
    unit_1 = ThisWorkbook.Sheets("Report").Range("D110")
    unit_2 = ThisWorkbook.Sheets("Report").Range("D111")
    unit_3 = ThisWorkbook.Sheets("Report").Range("D112")
    
    mask = "  [" & vbNewLine & "    {" & vbNewLine & "      ""id"": ""header_1""," & vbNewLine & "      ""title"": """ & header_1 & """," & vbNewLine & "      ""type"": ""automatic""," & vbNewLine & "      ""func"": []," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 1," & vbNewLine & "      ""field_query"": null" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""header_2""," & vbNewLine & "      ""title"": """ & header_2 & """," & vbNewLine & "      ""type"": ""automatic""," & vbNewLine & "      ""func"": []," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 2," & vbNewLine & "      ""field_query"": null" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id1""," & vbNewLine & "      ""title"": ""{0}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F76:G76")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(1, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 3," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(1, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id2""," & vbNewLine & "      ""title"": ""{1}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F77:G77")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(2, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 4," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(2, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id3""," & vbNewLine & "      ""title"": ""{2}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F78:G78")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(3, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 5," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(3, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id4""," & vbNewLine & "      ""title"": ""{3}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F79:G79")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(4, 3)) & "]," & vbNewLine & "      ""unit"": {" & vbNewLine & "        ""title"": """ & unit_2 & """," & vbNewLine & "        ""type"": ""count""" & vbNewLine & "      }," & vbNewLine & "      ""order"": 6," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(4, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id5""," & vbNewLine & "      ""title"": ""{4}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F80:G80")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(5, 3)) & "]," & vbNewLine & "      ""unit"": {" & vbNewLine & "        ""title"": """ & unit_1 & """," & vbNewLine & "        ""type"": ""count""" & vbNewLine & "      }," & vbNewLine & "      ""order"": 7," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(5, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id6""," & vbNewLine & "      ""title"": ""{5}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F81:G81")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(6, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 8," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(6, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id7""," & vbNewLine & "      ""title"": ""{6}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F82:G82")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(7, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 9," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(7, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id8""," & vbNewLine & "      ""title"": ""{7}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F83:G83")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(8, 3)) & "]," & vbNewLine & "      ""unit"": {" & vbNewLine & "        ""title"": """"," & vbNewLine & "        ""type"": ""count""" & vbNewLine & "      }," & vbNewLine & "      ""order"": 10," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(8, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id9""," & vbNewLine & "      ""title"": ""{8}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F84:G84")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(9, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 11," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(9, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id10""," & vbNewLine & "      ""title"": ""{9}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F85:G85")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(10, 3)) & "]," & vbNewLine & "      ""unit"": {" & vbNewLine & "        ""title"": """"," & vbNewLine & "        ""type"": ""count""" & vbNewLine & "      }," & vbNewLine & "      ""order"": 12," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(10, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id11""," & vbNewLine & "      ""title"": ""{10}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F86:G86")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(11, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 13," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(11, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id12""," & vbNewLine & "      ""title"": ""{11}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F87:G87")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(12, 3)) & "]," & vbNewLine & "      ""unit"": {" & vbNewLine & "        ""title"": """"," & vbNewLine & "        ""type"": ""count""" & vbNewLine & "      }," & vbNewLine & "      ""order"": 14," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(12, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id13""," & vbNewLine & "      ""title"": ""{12}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F88:G88")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(13, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 15," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(13, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id14""," & vbNewLine & "      ""title"": ""{13}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F89:G89")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(14, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 16," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(14, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id15""," & vbNewLine & "      ""title"": ""{14}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F90:G90")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(15, 3)) & "]," & vbNewLine & "      ""unit"": {" & vbNewLine & "        ""title"": ""PT""," & vbNewLine & "        ""type"": ""point""" & vbNewLine & "      }," & vbNewLine & "      ""order"": 17," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(15, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id16""," & vbNewLine & "      ""title"": ""{15}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F91:G91")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(16, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 18," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(16, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id17""," & vbNewLine & "      ""title"": ""{16}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F92:G92")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(17, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 19," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(17, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id18""," & vbNewLine & "      ""title"": ""{17}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F93:G93")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(18, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 20," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(18, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id19""," & vbNewLine & "      ""title"": ""{18}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F94:G94")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(19, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 21," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(19, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id20""," & vbNewLine & "      ""title"": ""{19}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F95:G95")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(20, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 22," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(20, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id21""," & vbNewLine & "      ""title"": ""{20}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F96:G96")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(21, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 23," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(21, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id22""," & vbNewLine & "      ""title"": ""{21}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F97:G97")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(22, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 24," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(22, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id23""," & vbNewLine & "      ""title"": ""{22}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F98:G98")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(23, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 25," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(23, 1), 3) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id24""," & vbNewLine & "      ""title"": ""{23}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F99:G99")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(24, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 26," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(24, 1), 3) & "" & vbNewLine & "    }" & vbNewLine & "  ]"

    For i = 0 To rangeInput.Count
        mask = Replace$(mask, "{" & i & "}", rangeInput.Cells(i + 1, 1).Value)
    Next
    GetReportForm3 = mask
End Function
Public Function GetReportForm4(rangeInput As Range) As String
    Dim i As Long
    Dim mask, header_1, header_2, unit_1, unit_2, unit_3 As String
    header_1 = ThisWorkbook.Sheets("Report").Range("D108")
    header_2 = ThisWorkbook.Sheets("Report").Range("D109")
    unit_1 = ThisWorkbook.Sheets("Report").Range("D110")
    unit_2 = ThisWorkbook.Sheets("Report").Range("D111")
    unit_3 = ThisWorkbook.Sheets("Report").Range("D112")
    
    mask = "  [" & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""header_1""," & vbNewLine & "      ""title"": """ & header_1 & """," & vbNewLine & "      ""type"": ""automatic""," & vbNewLine & "      ""func"": []," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 1," & vbNewLine & "      ""field_query"": null" & vbNewLine & "    }," & vbNewLine & "    {" & vbNewLine & "      ""id"": ""header_2""," & vbNewLine & "      ""title"": """ & header_2 & """," & vbNewLine & "      ""type"": ""automatic""," & vbNewLine & "      ""func"": []," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 2," & vbNewLine & "      ""field_query"": null" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id1""," & vbNewLine & "      ""title"": ""{0}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F103:G103")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(1, 3)) & "]," & vbNewLine & "      ""unit"": {" & vbNewLine & "        ""title"": ""PT""," & vbNewLine & "        ""type"": ""point""" & vbNewLine & "      }," & vbNewLine & "      ""order"": 3," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(1, 1), 4) & "" & vbNewLine & "    }," & vbNewLine
    mask = mask & "    {" & vbNewLine & "      ""id"": ""id2""," & vbNewLine & "      ""title"": ""{1}""," & vbNewLine & "      ""type"": " & GetReportTypeUnit(ThisWorkbook.Sheets("Report").Range("F104:G104")) & "," & vbNewLine & "      ""func"": [" & GetCalculationFormArray(rangeInput.Cells(2, 3)) & "]," & vbNewLine & "      ""unit"": null," & vbNewLine & "      ""order"": 4," & vbNewLine & "      ""field_query"": " & GetFieldQueryByTittleAndType(rangeInput.Cells(2, 1), 4) & "" & vbNewLine & "    }" & vbNewLine & "  ]"

    For i = 0 To rangeInput.Count
        mask = Replace$(mask, "{" & i & "}", rangeInput.Cells(i + 1, 1).Value)
    Next
    GetReportForm4 = mask
End Function
Public Function GetReportFormAM(rangeInput As Range) As String
    Dim i As Long
    Dim mask, header_1, header_2, unit_1, unit_2, unit_3 As String
    header_1 = ThisWorkbook.Sheets("Report").Range("D108")
    header_2 = ThisWorkbook.Sheets("Report").Range("D109")
    unit_1 = ThisWorkbook.Sheets("Report").Range("D110")
    unit_2 = ThisWorkbook.Sheets("Report").Range("D111")
    unit_3 = ThisWorkbook.Sheets("Report").Range("D112")
    
    mask = "    [" & vbNewLine & "        {" & vbNewLine & "            ""id"": ""header_1""," & vbNewLine & "            ""title"": """ & header_1 & """," & vbNewLine & "            ""type"": ""automatic""," & vbNewLine & "            ""func"": [" & vbNewLine & "            {" & vbNewLine & "                ""id"": ""id4""," & vbNewLine & "                ""type"": ""subtract""" & vbNewLine & "            }," & vbNewLine
    mask = mask & "            {" & vbNewLine & "                ""id"": ""id1""," & vbNewLine & "                ""type"": ""add""" & vbNewLine & "            }," & vbNewLine & "            {" & vbNewLine & "                ""id"": ""id2""," & vbNewLine & "                ""type"": ""subtract""" & vbNewLine & "            }," & vbNewLine & "            {" & vbNewLine & "                ""id"": ""id3""," & vbNewLine & "                ""type"": ""subtract""" & vbNewLine & "            }"
    mask = mask & vbNewLine & "            ]," & vbNewLine & "            ""unit"": null," & vbNewLine & "            ""order"": 1," & vbNewLine & "            ""field_query"": null" & vbNewLine & "        }," & vbNewLine & "        {" & vbNewLine & "            ""id"": ""header_2""," & vbNewLine & "            ""title"": """ & header_2 & """," & vbNewLine & "            ""type"": ""automatic""," & vbNewLine & "            ""func"": [" & vbNewLine & "            {" & vbNewLine & "                ""id"": ""id1""," & vbNewLine & "                ""type"": ""add""" & vbNewLine & "            }," & vbNewLine & "            {" & vbNewLine & "                ""id"": ""id14""," & vbNewLine & "                ""type"": ""add""" & vbNewLine & "            }," & vbNewLine & "            {" & vbNewLine & "                ""id"": ""id8""," & vbNewLine & "                ""type"": ""subtract""" & vbNewLine
    mask = mask & "            }," & vbNewLine & "            {" & vbNewLine & "                ""id"": ""id9""," & vbNewLine & "                ""type"": ""subtract""" & vbNewLine & "            }," & vbNewLine & "            {" & vbNewLine & "                ""id"": ""id10""," & vbNewLine & "                ""type"": ""subtract""" & vbNewLine & "            }," & vbNewLine & "            {" & vbNewLine & "                ""id"": ""id11""," & vbNewLine & "                ""type"": ""subtract""" & vbNewLine & "            }," & vbNewLine & "            {" & vbNewLine & "                ""id"": ""id12""," & vbNewLine & "                ""type"": ""subtract""" & vbNewLine & "            }," & vbNewLine & "            {" & vbNewLine & "                ""id"": ""id13""," & vbNewLine & "                ""type"": ""subtract""" & vbNewLine & "            }," & vbNewLine & "            {" & vbNewLine
    mask = mask & "                ""id"": ""id15""," & vbNewLine & "                ""type"": ""subtract""" & vbNewLine & "            }" & vbNewLine & "            ]," & vbNewLine & "            ""unit"": null," & vbNewLine & "            ""order"": 2," & vbNewLine & "            ""field_query"": null" & vbNewLine & "        }," & vbNewLine
    mask = mask & "        {" & vbNewLine & "            ""id"": ""id1""," & vbNewLine & "            ""title"": ""{0}""," & vbNewLine & "            ""type"": ""manual""," & vbNewLine & "            ""func"": [" & GetCalculationFormArray(rangeInput.Cells(1, 3)) & "]," & vbNewLine & "            ""unit"": null," & vbNewLine & "            ""order"": 3," & vbNewLine & "            ""print_order"": 1," & vbNewLine & "            ""field_query"": ""total_sales_amount""" & vbNewLine & "        }," & vbNewLine
    mask = mask & "        {" & vbNewLine & "            ""id"": ""id2""," & vbNewLine & "            ""title"": ""{1}""," & vbNewLine & "            ""type"": ""manual""," & vbNewLine & "            ""func"": [" & GetCalculationFormArray(rangeInput.Cells(2, 3)) & "]," & vbNewLine & "            ""unit"": null," & vbNewLine & "            ""order"": 4," & vbNewLine & "            ""print_order"": 2," & vbNewLine & "            ""field_query"": ""tax_amount""" & vbNewLine & "        }," & vbNewLine
    mask = mask & "        {" & vbNewLine & "            ""id"": ""id3""," & vbNewLine & "            ""title"": ""{2}""," & vbNewLine & "            ""type"": ""manual""," & vbNewLine & "            ""func"": [" & GetCalculationFormArray(rangeInput.Cells(3, 3)) & "]," & vbNewLine & "            ""unit"": null," & vbNewLine & "            ""order"": 5," & vbNewLine & "            ""print_order"": 3," & vbNewLine & "            ""field_query"": ""sales_deduction_amount""" & vbNewLine & "        }," & vbNewLine
    mask = mask & "        {" & vbNewLine & "            ""id"": ""id4""," & vbNewLine & "            ""title"": ""{3}""," & vbNewLine & "            ""type"": ""manual""," & vbNewLine & "            ""func"": [" & GetCalculationFormArray(rangeInput.Cells(4, 3)) & "]," & vbNewLine & "            ""order"": 6," & vbNewLine & "            ""print_order"": 6," & vbNewLine & "            ""field_query"": ""net_sales_amount""" & vbNewLine & "        }," & vbNewLine
    mask = mask & "        {" & vbNewLine & "            ""id"": ""id5""," & vbNewLine & "            ""title"": ""{4}""," & vbNewLine & "            ""type"": ""manual""," & vbNewLine & "            ""func"": [" & GetCalculationFormArray(rangeInput.Cells(5, 3)) & "]," & vbNewLine & "            ""unit"": null," & vbNewLine & "            ""order"": 7," & vbNewLine & "            ""print_order"": 7," & vbNewLine & "            ""field_query"": ""payment_amount""" & vbNewLine & "        }," & vbNewLine
    mask = mask & "        {" & vbNewLine & "            ""id"": ""id6""," & vbNewLine & "            ""title"": ""{5}""," & vbNewLine & "            ""type"": ""manual""," & vbNewLine & "            ""func"": [" & GetCalculationFormArray(rangeInput.Cells(6, 3)) & "]," & vbNewLine & "            ""unit"": {" & vbNewLine & "            ""type"": """"," & vbNewLine & "            ""title"": """"" & vbNewLine & "            }," & vbNewLine & "            ""order"": 8," & vbNewLine & "            ""print_order"": 8," & vbNewLine & "            ""field_query"": ""number_of_customers""" & vbNewLine & "        }," & vbNewLine
    mask = mask & "        {" & vbNewLine & "            ""id"": ""id7""," & vbNewLine & "            ""title"": ""{6}""," & vbNewLine & "            ""type"": ""manual""," & vbNewLine & "            ""func"": [" & GetCalculationFormArray(rangeInput.Cells(7, 3)) & "]," & vbNewLine & "            ""unit"": {" & vbNewLine & "            ""title"": """"," & vbNewLine & "            ""type"": """"" & vbNewLine & "            }," & vbNewLine & "            ""order"": 9," & vbNewLine & "            ""print_order"": 15," & vbNewLine & "            ""field_query"": ""settlement_number""" & vbNewLine & "        }," & vbNewLine
    mask = mask & "        {" & vbNewLine & "            ""id"": ""id8""," & vbNewLine & "            ""title"": ""{7}""," & vbNewLine & "            ""type"": ""manual""," & vbNewLine & "            ""func"": [" & GetCalculationFormArray(rangeInput.Cells(8, 3)) & "]," & vbNewLine & "            ""unit"": null," & vbNewLine & "            ""order"": 10," & vbNewLine & "            ""print_order"": 9," & vbNewLine & "            ""field_query"": ""cash_sales_amount""" & vbNewLine & "        }," & vbNewLine
    mask = mask & "        {" & vbNewLine & "            ""id"": ""id9""," & vbNewLine & "            ""title"": ""{8}""," & vbNewLine & "            ""type"": ""manual""," & vbNewLine & "            ""func"": [" & GetCalculationFormArray(rangeInput.Cells(9, 3)) & "]," & vbNewLine & "            ""unit"": null," & vbNewLine & "            ""order"": 11," & vbNewLine & "            ""print_order"": 10," & vbNewLine & "            ""field_query"": ""specified_credit_sales_amount""" & vbNewLine & "        }," & vbNewLine
    mask = mask & "        {" & vbNewLine & "            ""id"": ""id10""," & vbNewLine & "            ""title"": ""{9}""," & vbNewLine & "            ""type"": ""manual""," & vbNewLine & "            ""func"": [" & GetCalculationFormArray(rangeInput.Cells(10, 3)) & "]," & vbNewLine & "            ""unit"": null," & vbNewLine & "            ""order"": 12," & vbNewLine & "            ""print_order"": 11," & vbNewLine & "            ""field_query"": ""specified_electronic_money_sales_amount""" & vbNewLine & "        }," & vbNewLine
    mask = mask & "        {" & vbNewLine & "            ""id"": ""id11""," & vbNewLine & "            ""title"": ""{10}""," & vbNewLine & "            ""type"": ""manual""," & vbNewLine & "            ""func"": [" & GetCalculationFormArray(rangeInput.Cells(11, 3)) & "]," & vbNewLine & "            ""unit"": null," & vbNewLine & "            ""order"": 13," & vbNewLine & "            ""print_order"": 12," & vbNewLine & "            ""field_query"": ""specified_code_sales_amount""" & vbNewLine & "        }," & vbNewLine
    mask = mask & "        {" & vbNewLine & "            ""id"": ""id12""," & vbNewLine & "            ""title"": ""{11}""," & vbNewLine & "            ""type"": ""manual""," & vbNewLine & "            ""func"": [" & GetCalculationFormArray(rangeInput.Cells(12, 3)) & "]," & vbNewLine & "            ""unit"": null," & vbNewLine & "            ""order"": 14," & vbNewLine & "            ""print_order"": 13," & vbNewLine & "            ""field_query"": ""specified_gift_voucher_sales_amount""" & vbNewLine & "        }," & vbNewLine
    mask = mask & "        {" & vbNewLine & "            ""id"": ""id13""," & vbNewLine & "            ""title"": ""{12}""," & vbNewLine & "            ""type"": ""manual""," & vbNewLine & "            ""func"": [" & GetCalculationFormArray(rangeInput.Cells(13, 3)) & "]," & vbNewLine & "            ""unit"": null," & vbNewLine & "            ""order"": 15," & vbNewLine & "            ""print_order"": 4," & vbNewLine & "            ""field_query"": ""credit_sales_amount""" & vbNewLine & "        }," & vbNewLine
    mask = mask & "        {" & vbNewLine & "            ""id"": ""id14""," & vbNewLine & "            ""title"": ""{13}""," & vbNewLine & "            ""type"": ""manual""," & vbNewLine & "            ""func"": [" & GetCalculationFormArray(rangeInput.Cells(14, 3)) & "]," & vbNewLine & "            ""unit"": null," & vbNewLine & "            ""order"": 16," & vbNewLine & "            ""print_order"": 5," & vbNewLine & "            ""field_query"": ""credit_deposit_amount""" & vbNewLine & "        }," & vbNewLine
    mask = mask & "        {" & vbNewLine & "            ""id"": ""id15""," & vbNewLine & "            ""title"": ""{14}""," & vbNewLine & "            ""type"": ""manual""," & vbNewLine & "            ""func"": [" & GetCalculationFormArray(rangeInput.Cells(15, 3)) & "]," & vbNewLine & "            ""unit"": {" & vbNewLine & "            ""type"": ""pt""," & vbNewLine & "            ""title"": ""PT""" & vbNewLine & "            }," & vbNewLine & "            ""order"": 17," & vbNewLine & "            ""print_order"": 14," & vbNewLine & "            ""field_query"": ""point_utilization_amount""" & vbNewLine & "        }" & vbNewLine & "    ]"
    
    For i = 0 To rangeInput.Count
        mask = Replace$(mask, "{" & i & "}", rangeInput.Cells(i + 1, 1).Value)
    Next
    GetReportFormAM = mask
End Function

Public Function GetReportTypeUnit(rangeInput As Range) As String
    Dim output As String
    Dim f, g As Boolean
    f = (Not IsEmpty(rangeInput.Cells(1, 1).Value)) And (rangeInput.Cells(1, 1) = ThisWorkbook.Sheets("Report").Range("F1"))
    g = (Not IsEmpty(rangeInput.Cells(1, 2).Value)) And (rangeInput.Cells(1, 2) = ThisWorkbook.Sheets("Report").Range("F1"))
    If f And g Then
        output = """dynamic"""
    ElseIf f Then
        output = """query"""
    ElseIf g Then
        output = """manual"""
    Else: output = """manual"""
    End If
    GetReportTypeUnit = output
End Function
Public Function GetFieldQueryByTittleAndType(inputTittle As String, inputType As Integer) As String
    Dim output As String
    If inputType = 0 Then
        'am
        output = GetFieldQueryByTittle(inputTittle, ThisWorkbook.Sheets("tittle_query_map").Range("C6:D20"))
    ElseIf inputType = 1 Then
        'ar type 1
        output = GetFieldQueryByTittle(inputTittle, ThisWorkbook.Sheets("tittle_query_map").Range("C24:D44"))
    ElseIf inputType = 2 Then
        'ar type 2
        output = GetFieldQueryByTittle(inputTittle, ThisWorkbook.Sheets("tittle_query_map").Range("C48:D72"))
    ElseIf inputType = 3 Then
        'ar type 3
        output = GetFieldQueryByTittle(inputTittle, ThisWorkbook.Sheets("tittle_query_map").Range("C76:D99"))
    ElseIf inputType = 4 Then
        'at
        output = GetFieldQueryByTittle(inputTittle, ThisWorkbook.Sheets("tittle_query_map").Range("C103:D104"))
    End If
GetFieldQueryByTittleAndType = output
End Function

Public Function GetFieldQueryByTittle(inputTittle As String, rangeInput As Range) As String
    Dim output As String
    For i = 0 To rangeInput.Count
        If (rangeInput.Cells(i, 1).Value = inputTittle) Then
            If (IsEmpty(rangeInput.Cells(i, 2))) Then
                output = "null"
            Else: output = """" & rangeInput.Cells(i, 2) & """"
            End If
            Exit For
        End If
    Next
    GetFieldQueryByTittle = output
End Function
Public Function GetCalculationFormArray(inputString As String) As String
    Dim output, saveNumber, currentChar, firstNumber As String
    output = ""
    If (Len(inputString) > 0) Then
        
        Dim beforeEqual As Boolean
        Dim note As String
        beforeEqual = True
        currentChar = Left(inputString, 1)
        saveNumber = ""
        While (IsNumeric(currentChar))
            saveNumber = saveNumber & currentChar
            inputString = Right(inputString, Len(inputString) - 1)
            currentChar = Left(inputString, 1)
        Wend
        firstNumber = saveNumber
        saveNumber = ""
        
        If (currentChar = "=") Then
            beforeEqual = False
            note = "+"
        Else
            note = currentChar
        End If
        
        inputString = Right(inputString, Len(inputString) - 1)
        
        While (Len(inputString) > 0)
            currentChar = Left(inputString, 1)
            If (IsNumeric(currentChar)) Then
                saveNumber = saveNumber & currentChar
            ElseIf (currentChar = "=") Then
                If (note = "+") Then
                    output = output & GetCalculationFormUnit(saveNumber, "subtract", False)
                ElseIf (note = "-") Then
                    output = output & GetCalculationFormUnit(saveNumber, "add", False)
                End If
                saveNumber = ""
                note = "+"
                beforeEqual = False
            Else
                If (beforeEqual) Then
                    If (Len(saveNumber) > 0) Then
                        If (note = "+") Then
                            output = output & GetCalculationFormUnit(saveNumber, "subtract", False)
                        ElseIf (note = "-") Then
                            output = output & GetCalculationFormUnit(saveNumber, "add", False)
                        End If
                    End If
                Else
                    If (Len(saveNumber) > 0) Then
                        If (note = "+") Then
                            output = output & GetCalculationFormUnit(saveNumber, "add", False)
                        ElseIf (note = "-") Then
                            output = output & GetCalculationFormUnit(saveNumber, "subtract", False)
                        End If
                    End If
                End If
                saveNumber = ""
                note = currentChar
            End If
            inputString = Right(inputString, Len(inputString) - 1)
        Wend
        If (note = "+") Then
            output = output & GetCalculationFormUnit(saveNumber, "add", True)
        ElseIf (note = "-") Then
            output = output & GetCalculationFormUnit(saveNumber, "subtract", True)
        End If
        output = output & vbNewLine & "      "
    End If
    GetCalculationFormArray = output
End Function

Public Function GetCalculationFormUnit(ByVal inputId As String, typeInput As String, isLast As Boolean) As String
    Dim output As String
    output = ""
    output = output & vbNewLine & "        {" & vbNewLine
    output = output & "          ""id"": ""id" & inputId & """," & vbNewLine
    output = output & "          ""type"": """ & typeInput & """" & vbNewLine
    output = output & "        }"
    
    If Not isLast Then
        output = output & ","
    End If
    GetCalculationFormUnit = output
End Function








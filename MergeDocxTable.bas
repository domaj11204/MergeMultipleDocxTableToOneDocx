Attribute VB_Name = "MergeDocxTable"
Sub 多檔案彙整至總表()
    Dim masterDoc As Document
    Dim sourceDoc As Document
    Dim fileDialog As fileDialog
    Dim vFile As Variant
    Dim t As Integer, r As Integer, c As Integer
    Dim masterTable As Table, sourceTable As Table
    Dim masterText As String, sourceText As String
    Dim userChoice As VbMsgBoxResult
    
    ' 變數：備份用
    Dim fso As Object
    Dim backupPath As String
    
    ' 設定當前文件為「總檔」
    Set masterDoc = ActiveDocument
    
    ' =========================備份=========================
    On Error Resume Next
    
    
    ' 存檔
    masterDoc.Save
    
    ' 設定備份檔名 (加上時間戳記：Backup_YYYYMMDD_hhmiss_檔名.docx)
    backupPath = masterDoc.Path & Application.PathSeparator & _
                 "Backup_" & Format(Now, "yyyymmdd_hhmmss") & "_" & masterDoc.Name
    
    ' 備份
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile masterDoc.FullName, backupPath
    
    If Err.Number = 0 Then
        ' 備份成功，不干擾使用者，繼續執行
        ' Debug.Print "備份已建立：" & backupPath
    Else
        ' 備份失敗警告 (例如權限不足)
        MsgBox "警告：自動備份失敗！" & vbCrLf & _
               "建議您手動備份後再執行。", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    '=======================備份完畢=======================
    
    ' 檔案選擇視窗
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    fileDialog.Title = "請選擇要彙整進來的來源檔案 (可多選)"
    fileDialog.AllowMultiSelect = True
    
    If fileDialog.Show = -1 Then
        Application.ScreenUpdating = False ' 關閉畫面刷新
        
        ' 依序處理來源檔案
        For Each vFile In fileDialog.SelectedItems
            ' 在背景唯讀開啟來源檔
            Set sourceDoc = Documents.Open(FileName:=vFile, Visible:=False, ReadOnly:=True)
            If sourceDoc.Tables.Count <> masterDoc.Tables.Count Then
                 MsgBox "發現多餘表格！" & vbCrLf & vbCrLf & _
                "來源檔案：" & sourceDoc.Name & vbCrLf & _
                " 擁有 " & sourceDoc.Tables.Count & " 個表格，" & _
                "但總檔只有 " & masterDoc.Tables.Count & " 個。" & vbCrLf & vbCrLf & _
                "將忽略該檔案後續多出的表格。", _
                vbExclamation, "表格數量不符警告"
            End If
            
            ' 3. 遍歷該檔案的所有表格
            For t = 1 To sourceDoc.Tables.Count
                
                ' 多餘的Table則跳過
                If t > masterDoc.Tables.Count Then Exit For
                
                Set sourceTable = sourceDoc.Tables(t)
                Set masterTable = masterDoc.Tables(t)
                
                If sourceTable.Rows.Count <> masterTable.Rows.Count Then
                     MsgBox "列數不符！" & vbCrLf & vbCrLf & _
                    "來源檔案：" & sourceDoc.Name & vbCrLf & _
                    "第" & t & "個Table有" & sourceTable.Rows.Count & "列" & _
                    "總檔為 " & masterTable.Rows.Count & "列。" & vbCrLf & vbCrLf & _
                    "將忽略來源檔案多餘的列。", _
                    vbExclamation, "列數量不符警告"
                End If
                If sourceTable.Columns.Count <> masterTable.Columns.Count Then
                     MsgBox "欄數不符！" & vbCrLf & vbCrLf & _
                    "來源檔案：" & sourceDoc.Name & vbCrLf & _
                    "第" & t & "個Table有" & sourceTable.Columns.Count & "欄" & _
                    "總檔為 " & masterTable.Columns.Count & "欄。" & vbCrLf & vbCrLf & _
                    "將忽略來源檔案多餘的欄。", _
                    vbExclamation, "欄數量不符警告"
                End If
                ' 4. 遍歷每一個儲存格
                For r = 1 To sourceTable.Rows.Count
                    ' 忽略多餘的列
                    If r > masterTable.Rows.Count Then Exit For
                    
                    For c = 1 To sourceTable.Columns.Count
                        ' 忽略多餘的欄
                        If c > masterTable.Columns.Count Then Exit For
                        
                        ' 取得文字並清理 (移除控制字元)
                        sourceText = CleanCellText(sourceTable.Cell(r, c).Range.Text)
                        masterText = CleanCellText(masterTable.Cell(r, c).Range.Text)
                        
                        ' Master 空白 -> 自動填入
                        If Len(sourceText) > 0 And Len(masterText) = 0 Then
                            masterTable.Cell(r, c).Range.Text = sourceText
                            
                        ' 衝突，跳出通知
                        ElseIf Len(sourceText) > 0 And Len(masterText) > 0 And sourceText <> masterText Then
                            
                            ' 刷新畫面
                            Application.ScreenUpdating = True
                            
                            userChoice = MsgBox("發現資料衝突！" & vbCrLf & vbCrLf & _
                                         "檔案來源：" & sourceDoc.Name & vbCrLf & _
                                         "位置：第 " & t & " 個表格，第 " & r & " 列，第 " & c & " 欄" & vbCrLf & _
                                         "--------------------------------" & vbCrLf & _
                                         "總檔：" & masterText & vbCrLf & _
                                         "來源：" & sourceText & vbCrLf & _
                                         "--------------------------------" & vbCrLf & _
                                         "是否要「覆蓋」總檔內容？" & vbCrLf & _
                                         "(「是」覆蓋，「否」保留原樣)", _
                                         vbYesNo + vbExclamation, "資料衝突確認")
                            
                            If userChoice = vbYes Then
                                masterTable.Cell(r, c).Range.Text = sourceText
                            End If
                            ' 關閉刷新畫面
                            Application.ScreenUpdating = False
                            
                        End If
                    Next c
                Next r
            Next t
            
            ' 關閉來源檔
            sourceDoc.Close SaveChanges:=wdDoNotSaveChanges
        Next vFile
        
        Application.ScreenUpdating = True
        MsgBox "彙整完成！所有表格已更新。", vbInformation
        
    Else
        MsgBox "未選取任何檔案。"
    End If
    
End Sub

' --- 輔助函式：清除 Word 表格儲存格末尾的特殊符號 ---
Function CleanCellText(txt As String) As String
    Dim temp As String
    temp = txt
    ' 移除 Word 表格儲存格結尾的控制碼 (換行符號 Chr(13) + 鈴聲符號 Chr(7))
    If Len(temp) > 0 Then
        ' 移除最後的控制字元直到剩下純文字
        Do While Right(temp, 1) = Chr(13) Or Right(temp, 1) = Chr(7)
            temp = Left(temp, Len(temp) - 1)
            If Len(temp) = 0 Then Exit Do
        Loop
    End If
    CleanCellText = Trim(temp)
End Function




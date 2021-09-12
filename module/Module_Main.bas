Attribute VB_Name = "Module_Main"
Sub main()
 
    '各ワークシートを変数にset
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TO")
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("CC")
    Dim ws3 As Worksheet
    Set ws3 = ThisWorkbook.Sheets("BCC")
  
    'Outlookオブジェクトの作成
    Dim OutlookObj As Outlook.Application
    Set OutlookObj = New Outlook.Application
    
    '各シート１列目の最終行取得
    Dim r As Long, lastRow As Long, lastRow2 As Long
    '20210912UPD_ws.Cells(1, 1).End(xlDown).Rowだと１行の場合、最終行が正確に取得できないため独自Function使用
    'lastRow = ws.Cells(1, 1).End(xlDown).Row
    'lastRow2 = ws2.Cells(1, 1).End(xlDown).Row
    'lastRow3 = ws3.Cells(1, 1).End(xlDown).Row
    lastRow = GetLastRow(ws, "A")
    lastRow2 = GetLastRow(ws2, "A")
    lastRow3 = GetLastRow(ws3, "A")
   
    'メールアイテムオブジェクト作成
    Dim mailItemObj As Outlook.MailItem
    Set mailItemObj = OutlookObj.CreateItem(olMailItem)

    '各シートの宛先、CC、BCCのアドレスリストを変数に格納
    Dim ToAll As String, CCAll As String, BCCAll As String
    For r = 2 To lastRow
        ToAll = ToAll + ws.Cells(r, 1).Value + ";" '宛先リスト
    Next r
    For r = 2 To lastRow2
        CCAll = CCAll + ws2.Cells(r, 1).Value + ";" 'CCリスト
    Next r
    For r = 2 To lastRow3
        BCCAll = BCCAll + ws3.Cells(r, 1).Value + ";" 'BCCリスト
    Next r
    
    'メールアイテム作成
    With mailItemObj
        .To = ToAll 'Toを設定
        .CC = CCAll 'CCを設定
        .BCC = BCCAll 'BCCを設定
    End With
    
    '下書きメールアイテムを表示
    mailItemObj.Display
    '下書きフォルダに保存
    'mailItemObj.Save
    '直接送信
    'mailItemObj.Send

    '次のメールアイテムを作成するため破棄
    Set mailItemObj = Nothing
 
End Sub

'最終行取得Function
Function GetLastRow(ByVal sheet As Worksheet, ByVal col As String) As Long
   Dim lastRow As Long
   
   With sheet
       If .Range(col & 1).Value = "" Then
           '1行目の値がブランクの場合は最終行は0
           lastRow = 0
       ElseIf .Range(col & 2).Value = "" Then
           '2行目の値がブランクの場合は最終行は1行
           lastRow = 1
       Else
           lastRow = .Range(col & 1).End(xlDown).Row
       End If
   End With
   
   GetLastRow = lastRow
End Function

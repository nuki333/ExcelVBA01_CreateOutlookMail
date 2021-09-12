Attribute VB_Name = "Module_Main"
Sub main()
 
    '�e���[�N�V�[�g��ϐ���set
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TO")
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("CC")
    Dim ws3 As Worksheet
    Set ws3 = ThisWorkbook.Sheets("BCC")
  
    'Outlook�I�u�W�F�N�g�̍쐬
    Dim OutlookObj As Outlook.Application
    Set OutlookObj = New Outlook.Application
    
    '�e�V�[�g�P��ڂ̍ŏI�s�擾
    Dim r As Long, lastRow As Long, lastRow2 As Long
    '20210912UPD_ws.Cells(1, 1).End(xlDown).Row���ƂP�s�̏ꍇ�A�ŏI�s�����m�Ɏ擾�ł��Ȃ����ߓƎ�Function�g�p
    'lastRow = ws.Cells(1, 1).End(xlDown).Row
    'lastRow2 = ws2.Cells(1, 1).End(xlDown).Row
    'lastRow3 = ws3.Cells(1, 1).End(xlDown).Row
    lastRow = GetLastRow(ws, "A")
    lastRow2 = GetLastRow(ws2, "A")
    lastRow3 = GetLastRow(ws3, "A")
   
    '���[���A�C�e���I�u�W�F�N�g�쐬
    Dim mailItemObj As Outlook.MailItem
    Set mailItemObj = OutlookObj.CreateItem(olMailItem)

    '�e�V�[�g�̈���ACC�ABCC�̃A�h���X���X�g��ϐ��Ɋi�[
    Dim ToAll As String, CCAll As String, BCCAll As String
    For r = 2 To lastRow
        ToAll = ToAll + ws.Cells(r, 1).Value + ";" '���惊�X�g
    Next r
    For r = 2 To lastRow2
        CCAll = CCAll + ws2.Cells(r, 1).Value + ";" 'CC���X�g
    Next r
    For r = 2 To lastRow3
        BCCAll = BCCAll + ws3.Cells(r, 1).Value + ";" 'BCC���X�g
    Next r
    
    '���[���A�C�e���쐬
    With mailItemObj
        .To = ToAll 'To��ݒ�
        .CC = CCAll 'CC��ݒ�
        .BCC = BCCAll 'BCC��ݒ�
    End With
    
    '���������[���A�C�e����\��
    mailItemObj.Display
    '�������t�H���_�ɕۑ�
    'mailItemObj.Save
    '���ڑ��M
    'mailItemObj.Send

    '���̃��[���A�C�e�����쐬���邽�ߔj��
    Set mailItemObj = Nothing
 
End Sub

'�ŏI�s�擾Function
Function GetLastRow(ByVal sheet As Worksheet, ByVal col As String) As Long
   Dim lastRow As Long
   
   With sheet
       If .Range(col & 1).Value = "" Then
           '1�s�ڂ̒l���u�����N�̏ꍇ�͍ŏI�s��0
           lastRow = 0
       ElseIf .Range(col & 2).Value = "" Then
           '2�s�ڂ̒l���u�����N�̏ꍇ�͍ŏI�s��1�s
           lastRow = 1
       Else
           lastRow = .Range(col & 1).End(xlDown).Row
       End If
   End With
   
   GetLastRow = lastRow
End Function

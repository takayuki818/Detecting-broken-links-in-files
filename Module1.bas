Attribute VB_Name = "Module1"
Option Explicit
Sub �����N�؂�m�F()
    Dim �t�@�C���ꏊ As String
    Dim ��ƃu�b�N As Workbook
    Dim �V�[�g As Worksheet
    Dim ��`���T���� As String, ���͋K���T���� As String
    Dim �͈� As Variant
    '�_�C�A���O�{�b�N�X����Ώۃt�@�C���I��
    �t�@�C���ꏊ = Application.GetOpenFilename("Excel �u�b�N,*.xls?")
    If �t�@�C���ꏊ = "False" Then Exit Sub
    '�����N�X�V���s��Ȃ��w��Ńu�b�N���J��
    Set ��ƃu�b�N = Workbooks.Open(�t�@�C���ꏊ, UpdateLinks:=0)
    
    '���O��`�ӏ��̒T��
    ��`���T���� = ��`�������N�؂�T��(��ƃu�b�N)
    Select Case ��`���T����
        Case "": ��`���T���� = "���O�̒�`�̃����N�؂�ӏ����F����"
        Case Else: ��`���T���� = "���O�̒�`�̃����N�؂�ӏ����" & ��`���T����
    End Select
    
    '���͋K���ݒ�ӏ��̒T��
    For Each �V�[�g In ��ƃu�b�N.Worksheets
        On Error Resume Next
        '�V�[�g���̓��͋K�����ݒ肳��Ă���S�Z����ϐ��i�[�����݂��Ȃ��ꍇ�G���[
        Set �͈� = �V�[�g.Cells.SpecialCells(xlCellTypeAllValidation)
        If Not (�͈� Is Nothing) Then
            ���͋K���T���� = ���͋K���T���� & �V�[�g�����͋K�������N�؂�T��(�V�[�g)
        End If
        On Error GoTo 0
    Next
    Select Case ���͋K���T����
        Case "": ���͋K���T���� = "���X�g�^���͋K���̃����N�؂�ӏ����F����"
        Case Else: ���͋K���T���� = "���X�g�^���͋K���̃����N�؂�ӏ����" & ���͋K���T����
    End Select
    
    MsgBox ��`���T���� & vbCrLf & vbCrLf & ���͋K���T����
    ��ƃu�b�N.Close SaveChanges:=False
    Set ��ƃu�b�N = Nothing
End Sub
Function ��`�������N�؂�T��(��ƃu�b�N As Workbook)
    Dim ��`�� As Variant
    For Each ��`�� In ��ƃu�b�N.Names
        If InStr(��`��.RefersTo, "#REF") > 0 Or InStr(��`��.RefersTo, ".xl") > 0 Then
            ��`�������N�؂�T�� = ��`�������N�؂�T�� & vbCrLf & ��`��.Name & " : " & ��`��.RefersTo
        End If
    Next
End Function
Function �V�[�g�����͋K�������N�؂�T��(�V�[�g As Worksheet) As String
    Dim �͈� As Variant
    For Each �͈� In �V�[�g.Cells.SpecialCells(xlCellTypeAllValidation)
        '���͋K���̎�ނ��u���X�g�v�`���̏ꍇ
        If �͈�.Validation.Type = xlValidateList Then
            If InStr(�͈�.Validation.Formula1, "#REF") > 0 Or InStr(�͈�.Validation.Formula1, ".xl") > 0 Then
                �V�[�g�����͋K�������N�؂�T�� = �V�[�g�����͋K�������N�؂�T�� & vbCrLf & �V�[�g.Name & " : " & �͈�.Address(False, False) & �͈�.Validation.Formula1
            End If
        End If
    Next
End Function

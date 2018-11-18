Attribute VB_Name = "Module1"
Option Explicit

'�L����������
Function IsYugen(ByVal numerator As Variant, ByVal denominator As Variant) As Boolean
    IsYugen = IsFiniteDecimal(numerator, denominator)
End Function

'�L����������
Private Function IsFiniteDecimal(ByVal numerator As Variant, ByVal denominator As Variant) As Boolean

    '���ꕪ�q�Ƃ������ɂ���
    Dim slidecount As Integer
    slidecount = MaxInt(GetUnderInteger(numerator), GetUnderInteger(denominator))
    numerator = Slide(numerator, slidecount)
    denominator = Slide(denominator, slidecount)
    
    '��
    Dim gcd As Long
    gcd = GetGCD(numerator, denominator)
    '���ꂾ���ɂ��Ă����i���q�͂���Ȍ�̔���ɂ͊֌W�Ȃ��j
    denominator = denominator / gcd
    
    '�񕪂���������Ђ�����2��5�Ŋ����Ă���
    denominator = warisusume(denominator, 2)
    denominator = warisusume(denominator, 5)
    
    '���Ŋ���؂������ʂ�1�i��������-1�j�Ȃ�Anumerator / denominator�͗L������
    IsFiniteDecimal = denominator = 1 Or denominator = -1
End Function

'�����_�ȉ��̌�����Ԃ�
Private Function GetUnderInteger(ByVal v As Variant) As Integer
    Dim keta As String
    keta = CStr(CDec(v))
    
    Dim pos As Integer
    pos = InStr(keta, ".")
    
    If (pos = 0) Then
        GetUnderInteger = 0
    Else
        GetUnderInteger = Len(keta) - pos
    End If
End Function

'�傫������Ԃ�
Private Function MaxInt(ByVal v1 As Integer, ByVal v2 As Integer) As Integer
    If v1 > v2 Then
        MaxInt = v1
    Else
        MaxInt = v2
    End If
End Function

'�����_�̈ʒu��keta�����E�Ɉړ�
Private Function Slide(ByVal v As Variant, ByVal keta As Integer) As Long
    Slide = CDec(v) * 10 ^ keta
End Function

'�ő���񐔂�Ԃ�
Private Function GetGCD(ByVal v1 As Long, ByVal v2 As Long) As Long
    Dim amari As Long
    
    amari = v1 Mod v2
    If amari = 0 Then
        GetGCD = v2
    Else
        GetGCD = GetGCD(v2, amari)
    End If
End Function

'����؂�Ȃ��Ȃ�܂łЂ����犄���Ă���
Private Function warisusume(ByVal v As Long, ByVal n As Long) As Long
    Do
        If v Mod n = 0 Then
            v = v / n
        Else
            Exit Do
        End If
    Loop
    warisusume = v
End Function

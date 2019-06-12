Attribute VB_Name = "Module1"
Option Explicit

'セルの値（数値）取得
'文字列で取得したのをDecimalに変換しなおす。
Function CellTextToValue(ByVal cell As Variant) As Variant
    Dim s As String
    s = cell.Text
    CellTextToValue = CDec(s)
End Function

'有限小数判定
Function IsYugen(ByVal numeratorcell As Variant, ByVal denominatorcell As Variant) As Boolean
    IsYugen = IsFiniteDecimal(CellTextToValue(numeratorcell), CellTextToValue(denominatorcell))
End Function

'有限小数判定
Private Function IsFiniteDecimal(ByVal numerator As Variant, ByVal denominator As Variant) As Boolean

    '0対策
    If denominator = 0 Then denominator = 1

    '分母分子とも整数にする
    Dim slidecount As Integer
    slidecount = MaxInt(GetUnderInteger(numerator), GetUnderInteger(denominator))
    numerator = Slide(numerator, slidecount)
    denominator = Slide(denominator, slidecount)
    
    '約分
    Dim gcd As Long
    gcd = GetGCD(numerator, denominator)
    '分母だけにしておく（分子はこれ以後の判定には関係ない）
    denominator = denominator / gcd
    
    '約分した分母をひたすら2と5で割っていく
    denominator = warisusume(denominator, 2)
    denominator = warisusume(denominator, 5)
    
    '↑で割り切った結果が1（もしくは-1）なら、numerator / denominatorは有限小数
    IsFiniteDecimal = denominator = 1 Or denominator = -1
End Function

'小数点以下の桁数を返す
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

'大きい方を返す
Private Function MaxInt(ByVal v1 As Integer, ByVal v2 As Integer) As Integer
    If v1 > v2 Then
        MaxInt = v1
    Else
        MaxInt = v2
    End If
End Function

'小数点の位置をketaだけ右に移動
Private Function Slide(ByVal v As Variant, ByVal keta As Integer) As Long
    Slide = CDec(v) * 10 ^ keta
End Function

'最大公約数を返す
Private Function GetGCD(ByVal v1 As Long, ByVal v2 As Long) As Long
    Dim amari As Long
    
    amari = v1 Mod v2
    If amari = 0 Then
        GetGCD = v2
    Else
        GetGCD = GetGCD(v2, amari)
    End If
End Function

'割り切れなくなるまでひたすら割っていく
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

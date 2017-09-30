Public Structure Rational
    Private _gcd As Long '用于约分的最大公约数
    Private Property fenzi As Long
    Private Property fenmu As Long
    Public Sub New(numerator As Long, denominator As Long)
        '分子分母都是整数
        '自动约分
        If denominator < 0 Then '负数统一为分子上带负号
            denominator *= -1
            numerator *= -1
        End If
        '约分过程，可以考虑先判断是否需要约分
        _gcd = gcd(numerator, denominator)
        fenzi = numerator / _gcd
        fenmu = denominator / _gcd
    End Sub

    Public Sub New(num As Double)
        '浮点数化为分数
        fenmu = 1
        If Fix(num) = num Then
            fenzi = num
        Else
            Do While Math.Abs(Fix(num) - num) >= 0.0000000001
                num *= 10
                fenmu *= 10
            Loop
            _gcd = gcd(num, fenmu)
            fenzi = num / _gcd
            fenmu /= _gcd
        End If
    End Sub

    Public ReadOnly Property Space As Integer
        '计算打印时在控件上占用的像素数(待淘汰)
        Get
            Dim l1 As Integer = Math.Abs(fenzi).ToString.Length, l2 As Integer = fenmu.ToString.Length
            If fenzi >= 0 Then
                If fenmu = 1 Then
                    Return l1 * 6
                Else
                    Return Math.Max(l1, l2) * 6 + 2
                End If
            Else
                If fenmu = 1 Then
                    Return l1 * 6 + 6
                Else
                    Return Math.Max(l1, l2) * 6 + 9
                End If
            End If
        End Get
    End Property

    Shared Operator +(a As Rational, b As Rational) As Rational
        If a.fenmu = b.fenmu Then Return New Rational(a.fenzi + b.fenzi, a.fenmu)
        '分母不同时通分过程待修改
        Return New Rational(a.fenzi * b.fenmu + b.fenzi * a.fenmu, a.fenmu * b.fenmu)
    End Operator

    Shared Operator -(a As Rational, b As Rational) As Rational
        If a.fenmu = b.fenmu Then Return New Rational(a.fenzi - b.fenzi, a.fenmu)
        '分母不同时通分过程待修改
        Return New Rational(a.fenzi * b.fenmu - b.fenzi * a.fenmu, a.fenmu * b.fenmu)
    End Operator

    Shared Operator -(a As Rational) As Rational
        a.fenzi *= -1
    End Operator

    Shared Operator *(a As Rational, b As Rational) As Rational
        '分子分母应先约分
        Return New Rational(a.fenzi * b.fenzi, a.fenmu * b.fenmu)
    End Operator

    Shared Operator /(a As Rational, b As Rational) As Rational
        '分子分母应先约分
        Return New Rational(a.fenzi * b.fenmu, a.fenmu * b.fenzi)
    End Operator

    Shared Operator =(a As Rational, b As Integer) As Boolean
        If a.fenzi = b AndAlso a.fenmu = 1 Then Return True
        Return False
    End Operator

    Shared Operator =(a As Rational, b As Rational) As Boolean
        If a.fenzi = b.fenzi AndAlso a.fenmu = b.fenmu Then Return True
        Return False
    End Operator

    Shared Operator <>(a As Rational, b As Rational) As Boolean
        If a.fenzi <> b.fenzi OrElse a.fenmu <> b.fenmu Then Return True
        Return False
    End Operator

    Shared Operator <>(a As Rational, b As Integer) As Boolean
        If a.fenmu <> 1 OrElse a.fenzi <> b Then Return True
        Return False
    End Operator

    Public Overrides Function ToString() As String
        If fenmu = 1 Then
            Return fenzi
        Else
            Return fenzi & "/" & fenmu
        End If
    End Function

    Private Shared Function gcd(a As Long, b As Long) As Long
        If b Then Return gcd(b, a Mod b)
        Return Math.Abs(a)
    End Function

    Public Sub Print(control As Control, x As Integer, y As Integer)
        '将该有理数打印在指定控件的指定位置上
        '目前打印的大小固定，以后考虑可以改变字号
        Dim l As New Label With {.Location = New Point(x, y), .Size = New Size(Space, 29)}
        control.Controls.Add(l)
        l.Refresh()
        Dim l1 As Integer = Math.Abs(fenzi).ToString.Length, l2 As Integer = fenmu.ToString.Length
        Dim g As Graphics = l.CreateGraphics, number As New Font("Times New Roman", 9), brush As New SolidBrush(Color.Black)
        If fenzi >= 0 Then
            If fenmu = 1 Then
                g.DrawString(fenzi, number, brush, -2, 7)
            Else
                g.DrawLine(Pens.Black, 0, 14, Space, 14)
                If l1 >= l2 Then
                    With g
                        .DrawString(fenzi, number, brush, -1, -1)
                        .DrawString(fenmu, number, brush, (l1 - l2) * 3 - 1, 15)
                    End With
                Else
                    With g
                        .DrawString(fenzi, number, brush, (l2 - l1) * 3 - 1, -1)
                        .DrawString(fenmu, number, brush, -1, 15)
                    End With
                End If
            End If
        Else
            If fenmu = 1 Then
                With g
                    .DrawLine(Pens.Black, 0, 14, 4, 14)
                    .DrawString(-fenzi, number, brush, 4, 7)
                End With
            Else
                With g
                    .DrawLine(Pens.Black, 0, 14, 4, 14)
                    .DrawLine(Pens.Black, 7, 14, Space, 14)
                End With
                If l1 >= l2 Then
                    With g
                        .DrawString(-fenzi, number, brush, 6, -1)
                        .DrawString(fenmu, number, brush, (l1 - l2) * 3 + 6, 15)
                    End With
                Else
                    With g
                        .DrawString(-fenzi, number, brush, (l2 - l1) * 3 + 6, -1)
                        .DrawString(fenmu, number, brush, 6, 15)
                    End With
                End If
            End If
        End If
        g.Dispose()
    End Sub

    Public Function Round(digits As Integer) As String
        Return FormatNumber(fenzi / fenmu, digits)
    End Function
End Structure
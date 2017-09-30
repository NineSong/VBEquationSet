Public Class Form1

#Region "Variables" '定义变量
    Dim x, y, z As Rational '三个未知数
    Dim num() As Long = {1, 1, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0} '相当于增广矩阵
    Dim g As Graphics, letter As New Font("Times New Roman", 10, FontStyle.Italic), brush As New SolidBrush(Color.Black) '用于显示结果
#End Region

#Region "Functions" '处理事件
    '控制输入(待简化)
    Private Sub Control(sender As TextBox, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress, TextBox2.KeyPress, TextBox3.KeyPress, TextBox4.KeyPress, TextBox5.KeyPress, TextBox6.KeyPress, TextBox7.KeyPress, TextBox8.KeyPress, TextBox9.KeyPress, TextBox10.KeyPress, TextBox11.KeyPress, TextBox12.KeyPress
        Dim key = e.KeyChar, text = sender.Text
        If Char.IsDigit(key) OrElse key = Chr(8) Then
            If text = "-" AndAlso key = "0" Then e.Handled = True
            If text = "0" Then sender.Text = ""
        ElseIf key = "-" AndAlso text = "" Then
        ElseIf key = Chr(13) Then
            e.Handled = True
            SendKeys.Send("{Tab}")
        Else
            e.Handled = True
            Beep()
        End If
    End Sub
    '保存A(系数)
    Private Sub Format1(sender As TextBox, e As System.EventArgs) Handles TextBox1.Leave, TextBox2.Leave, TextBox3.Leave, TextBox5.Leave, TextBox6.Leave, TextBox7.Leave, TextBox9.Leave, TextBox10.Leave, TextBox11.Leave
        With sender
            Select Case .Text
                Case ""
                Case "1" : .Text = ""
                Case "-1"
                    num(.TabIndex - 1) = -1
                    .Text = "-"
                Case "-" : num(.TabIndex - 1) = -1
                Case Else : num(.TabIndex - 1) = .Text
            End Select
        End With
    End Sub
    '保存b(等号右边)
    Private Sub Format2(sender As TextBox, e As System.EventArgs) Handles TextBox4.Leave, TextBox8.Leave, TextBox12.Leave
        With sender
            If .Text = "" OrElse .Text = "-" Then .Text = "0"
            num(.TabIndex - 1) = .Text
        End With
    End Sub
    '解线性方程组
    Private Sub Solve() Handles Button1.Click
        Dim D, Dx, Dy, Dz As Long
        If RadioButton1.Checked Then '二元
            If TextBox4.Text = "" Then TextBox4.Text = "0"
            If TextBox8.Text = "" Then TextBox8.Text = "0"
            Try
                D = num(0) * num(4) - num(3) * num(1)
                If D <> 0 Then
                    GroupBox3.Tag = "有解"
                    Dx = num(2) * num(4) - num(5) * num(1)
                    Dy = num(0) * num(5) - num(3) * num(2)
                    x = New Rational(Dx, D)
                    y = New Rational(Dy, D)
                    Print(x, y)
                Else
                    With GroupBox3
                        .Tag = "无解"
                        .CreateGraphics.DrawString("/", letter, brush, 188, 28)
                    End With
                End If
            Catch ex As OverflowException
                Clear()
                MsgBox("数据溢出", 64)
            End Try
        ElseIf RadioButton2.Checked Then '三元
            If TextBox4.Text = "" Then TextBox4.Text = "0"
            If TextBox8.Text = "" Then TextBox8.Text = "0"
            If TextBox12.Text = "" Then TextBox12.Text = "0"
            Try
                D = num(0) * num(5) * num(10) + num(4) * num(9) * num(2) + num(8) * num(1) * num(6) - num(8) * num(5) * num(2) - num(4) * num(1) * num(10) - num(0) * num(9) * num(6)
                If D <> 0 Then
                    GroupBox3.Tag = "有解"
                    Dx = num(3) * num(5) * num(10) + num(7) * num(9) * num(2) + num(11) * num(1) * num(6) - num(11) * num(5) * num(2) - num(7) * num(1) * num(10) - num(3) * num(9) * num(6)
                    Dy = num(0) * num(7) * num(10) + num(4) * num(11) * num(2) + num(8) * num(3) * num(6) - num(8) * num(7) * num(2) - num(4) * num(3) * num(10) - num(0) * num(11) * num(6)
                    Dz = num(0) * num(5) * num(11) + num(4) * num(9) * num(3) + num(8) * num(1) * num(7) - num(8) * num(5) * num(3) - num(4) * num(1) * num(11) - num(0) * num(9) * num(7)
                    x = New Rational(Dx, D)
                    y = New Rational(Dy, D)
                    z = New Rational(Dz, D)
                    Print(x, y, z)
                Else
                    With GroupBox3
                        .Tag = "无解"
                        .CreateGraphics.DrawString("/", letter, brush, 188, 28)
                    End With
                End If
            Catch ex As OverflowException
                Clear()
                MsgBox("数据溢出", 64)
            End Try
        ElseIf RadioButton3.Checked Then '自定义(高斯消元法)_待改进
            Dim m, n As Integer, mat(,) As Rational
            m = InputBox("请输入行数")
            n = InputBox("请输入列数")
            ReDim mat(m - 1, n - 1)
            With RichTextBox1
                If .Text <> "" Then
                    For i = 1 To 60
                        .Text &= "-"
                    Next
                    .Text &= vbCrLf & vbCrLf
                End If
                .Text &= "A = " & vbCrLf
                Try
                    For i = 0 To m - 1
                        For j = 0 To n - 1
                            Do While mat(i, j) = Nothing
                                Dim response = InputBox("请输入第" & i + 1 & "行，第" & j + 1 & "列的数", "输入数据", "0")
                                If IsNumeric(response) Then
                                    mat(i, j) = New Rational(response)
                                ElseIf response = "" Then
                                    .Text &= vbCrLf & vbCrLf
                                    Exit Sub
                                Else
                                    MsgBox("请输入数字！", 64)
                                End If
                            Loop
                            .Text &= mat(i, j).ToString & "     "
                        Next
                        .Text &= vbCrLf
                    Next
                    Dim finished As Boolean = False
                    Dim _i, rank As Integer '_i：主元向右的偏移量，(i,i+_i)为理想的主元
                    Dim c As Rational '消元时用到的数
                    Dim pivot(n - 1) As Point
                    Dim row_of_pivot(n - 1), non_pivot(n - 1) As Integer
                    Dim result(n - 2) As Rational
                    For i = 0 To m - 2 '从第一行开始消元
                        If i + _i = n Then Exit For '如果当前行无主元则消元结束
                        If mat(i, i + _i) = 0 Then '如果主元的位置为0，则开始行交换
                            For j = i + _i To n - 1 '从主元所在位置开始到最后一列
                                For k = i + 1 To m - 1 '从主元下面一行开始一直到最后一行
                                    If mat(k, j) <> 0 Then '如果在第k行的同一列找到了非零的元素
                                        Dim temp As Rational
                                        For l = 0 To n - 1
                                            temp = mat(i, l)
                                            mat(i, l) = mat(k, l)
                                            mat(k, l) = temp
                                        Next
                                        finished = True '找到主元
                                        Exit For '行交换成功
                                    End If
                                Next
                                If finished Then Exit For '如果行交换成功就跳出，开始消元
                                _i += 1 '主元的偏移量+1
                                If j <> n - 1 AndAlso mat(i, j + 1) <> 0 Then Exit For '如果行交换失败，但是主元右边的位置非零，则找到主元
                            Next '如果主元右边的位置还是0，那么就从主元右边一列开始寻找主元
                        End If
                        If i + _i = n Then Exit For
                        row_of_pivot(i + _i) = i + 1 '该列为主列，记录该列中主元的行号
                        rank += 1 '矩阵的秩+1
                        pivot(rank - 1) = New Point(i, i + _i) '保存主元的位置
                        For j = i + 1 To m - 1
                            If mat(j, i + _i) <> 0 Then
                                c = mat(j, i + _i) / mat(i, i + _i)
                                For k = i + _i To n - 1
                                    mat(j, k) -= mat(i, k) * c
                                Next
                            End If
                        Next
                    Next
                    finished = False
                    If m + _i - 1 < n Then
                        For i = m + _i - 1 To n - 2
                            If mat(m - 1, i) <> 0 Then
                                rank += 1
                                row_of_pivot(i) = m
                                pivot(rank - 1) = New Point(m - 1, i)
                                finished = True
                                Exit For
                            End If
                        Next
                        If Not finished Then
                            If mat(m - 1, n - 1) <> 0 Then
                                rank += 1
                                pivot(rank - 1) = New Point(m - 1, n - 1)
                            End If
                        End If
                    End If
                    If rank >= 1 Then
                        For i = rank - 1 To 0 Step -1
                            c = mat(pivot(i).X, pivot(i).Y)
                            If c <> 1 Then '如果该主元不为1
                                For j = pivot(i).Y To n - 1 '从主元所在的列到最后一列
                                    mat(pivot(i).X, j) /= c '该行每个数都除以该主元
                                Next
                            End If
                            If pivot(i).X <> 0 Then '如果该主元不在第一行
                                For j = pivot(i).X - 1 To 0 Step -1 '从主元所在行到第一行
                                    c = mat(j, pivot(i).Y)
                                    For k = pivot(i).Y To n - 1 '从主元所在列到最后一列
                                        mat(j, k) -= mat(pivot(i).X, k) * c
                                    Next
                                Next
                            End If
                        Next
                    End If
                    .Text &= vbCrLf
                    If rank <> 0 Then
                        If pivot(rank - 1).Y = n - 1 Then
                            .Text &= "无解"
                        Else
                            .Text &= "x = ["
                            If rank = n - 1 Then
                                For i = 0 To rank - 1
                                    .Text &= mat(i, n - 1).ToString & ","
                                Next
                                .Text = .Text.Substring(0, .TextLength - 1) & "]^T"
                            Else
                                Dim count As Integer
                                For i = 0 To n - 2
                                    If row_of_pivot(i) > 0 Then
                                        .Text &= mat(row_of_pivot(i) - 1, n - 1).ToString & ","
                                    Else
                                        non_pivot(count) = i
                                        count += 1
                                        .Text &= "0,"
                                    End If
                                Next
                                .Text = .Text.Substring(0, .TextLength - 1) & "]^T + "
                                For i = 0 To n - rank - 2
                                    .Text &= "c" & i + 1 & "["
                                    For j = 0 To n - rank - 2
                                        result(non_pivot(j)) = New Rational(0)
                                    Next
                                    result(non_pivot(i)) = New Rational(1)
                                    For j = 0 To rank - 1
                                        result(pivot(j).Y) = -mat(pivot(j).X, non_pivot(i))
                                    Next
                                    For j = 0 To n - 2
                                        .Text &= result(j).ToString & ","
                                    Next
                                    .Text = .Text.Substring(0, .TextLength - 1) & "]^T + "
                                Next
                                .Text = .Text.Substring(0, .TextLength - 3)
                            End If
                        End If
                    End If
                Catch ex As OverflowException
                    MsgBox("数据溢出", 64)
                End Try
                .Text &= vbCrLf & vbCrLf
            End With
        End If
    End Sub
    '还原到初始状态
    Private Sub Clear() Handles Button2.Click
        GroupBox3.Tag = Nothing
        If Not RadioButton3.Checked Then
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox8.Text = ""
            If RadioButton1.Checked Then
                num = {1, 1, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0}
            ElseIf RadioButton2.Checked Then
                num = {1, 1, 1, 0, 1, 1, 1, 0, 1, 1, 1, 0}
                TextBox6.Text = ""
                TextBox7.Text = ""
                TextBox9.Text = ""
                TextBox10.Text = ""
                TextBox11.Text = ""
                TextBox12.Text = ""
            End If
            With GroupBox3
                .Controls.Clear()
                .Refresh()
            End With
            TextBox1.Focus()
        Else
            With RichTextBox1
                .Text = ""
                .Focus()
            End With
        End If
    End Sub
    '重绘
    Private Sub Repaint() Handles GroupBox1.Paint
        If GroupBox3.Tag = "有解" Then
            If RadioButton1.Checked Then
                Print(x, y)
            ElseIf RadioButton2.Checked Then
                Print(x, y, z)
            End If
        ElseIf GroupBox3.Tag = "无解" Then
            With GroupBox3
                .Refresh()
                .CreateGraphics.DrawString("/", letter, brush, 188, 28)
            End With
        End If
    End Sub

#End Region

#Region "UI"
    '变为二元的界面
    Private Sub Mode1() Handles RadioButton1.MouseClick
        Me.Size = New Size(416, 228)
        GroupBox2.Size = New Size(376, 79)
        GroupBox3.Location = New Point(12, 122)
        '----------------------------------------------------------
        Label1.Location = New Point(130, 20)
        With Label2
            .Location = New Point(223, 20)
            .Text = "y ="
        End With
        With Label3
            .Location = New Point(130, 49)
            .Text = "x +"
        End With
        With Label4
            .Location = New Point(223, 49)
            .Text = "y ="
        End With
        Label5.Visible = False
        Label6.Visible = False
        '-----------------------------------------------------------
        TextBox1.Location = New Point(65, 20)
        TextBox2.Location = New Point(158, 20)
        With TextBox3
            .Location = New Point(65, 49)
            .TabIndex = 4
        End With
        With TextBox4
            .Location = New Point(249, 20)
            .TabIndex = 3
        End With
        TextBox5.Location = New Point(158, 49)
        With TextBox6
            .TabIndex = 8
            .TabStop = False
            .Visible = False
        End With
        With TextBox7
            .TabStop = False
            .Visible = False
        End With
        With TextBox8
            .Location = New Point(249, 49)
            .TabIndex = 6
        End With
        TextBox9.TabStop = False
        TextBox10.TabStop = False
        TextBox11.TabStop = False
        TextBox12.TabStop = False
        '-----------------------------------------------------------
        RichTextBox1.Visible = False
        Clear()
    End Sub
    '变为三元的界面
    Private Sub Mode2() Handles RadioButton2.MouseClick
        Me.Size = New Size(416, 259)
        GroupBox2.Size = New Size(376, 111)
        GroupBox3.Location = New Point(12, 154)
        '-----------------------------------------------------
        Label1.Location = New Point(82, 20)
        With Label2
            .Location = New Point(175, 20)
            .Text = "y +"
        End With
        With Label3
            .Location = New Point(268, 20)
            .Text = "z ="
        End With
        With Label4
            .Location = New Point(82, 49)
            .Text = "x +"
        End With
        With Label5
            .Location = New Point(175, 49)
            .Visible = True
        End With
        With Label6
            .Location = New Point(268, 49)
            .Visible = True
        End With
        '------------------------------------------------------
        TextBox1.Location = New Point(18, 20)
        TextBox2.Location = New Point(110, 20)
        With TextBox3
            .Location = New Point(203, 20)
            .TabIndex = 3
        End With
        With TextBox4
            .Location = New Point(296, 20)
            .TabIndex = 4
        End With
        TextBox5.Location = New Point(18, 49)
        With TextBox6
            .Location = New Point(110, 49)
            .TabStop = True
            .TabIndex = 6
            .Visible = True
        End With
        With TextBox7
            .Location = New Point(203, 49)
            .TabStop = True
            .Visible = True
        End With
        With TextBox8
            .Location = New Point(296, 49)
            .TabIndex = 8
        End With
        TextBox9.TabStop = True
        TextBox10.TabStop = True
        TextBox11.TabStop = True
        TextBox12.TabStop = True
        '--------------------------------------------------------
        RichTextBox1.Visible = False
        Clear()
    End Sub
    '变为自定义的界面
    Private Sub Mode3() Handles RadioButton3.MouseClick
        Me.Size = New Size(416, 270)
        RichTextBox1.Visible = True
        Clear()
    End Sub

#End Region

#Region "Results"
    '显示二元的结果
    Private Sub Print(x As Rational, y As Rational)
        GroupBox3.Controls.Clear()
        x.Print(GroupBox3, 179 - ((x.Space + y.Space) >> 1), 17)
        y.Print(GroupBox3, 221 + ((x.Space - y.Space) >> 1), 17)
        Dim l As New Label With {.Size = New Size(66 + x.Space + y.Space, 29), .Location = New Point(155 - ((x.Space + y.Space) >> 1), 17)}
        GroupBox3.Controls.Add(l)
        l.Refresh()
        g = l.CreateGraphics
        With g
            .DrawString("x = ", letter, brush, 0, 7)
            .DrawString(",   y = ", letter, brush, x.Space + 30, 7)
            .Dispose()
        End With
    End Sub
    '显示三元的结果
    Private Sub Print(x As Rational, y As Rational, z As Rational)
        GroupBox3.Controls.Clear()
        x.Print(GroupBox3, 158 - ((x.Space + y.Space + z.Space) >> 1), 17)
        y.Print(GroupBox3, 200 + ((x.Space - y.Space - z.Space) >> 1), 17)
        z.Print(GroupBox3, 242 + ((x.Space + y.Space - z.Space) >> 1), 17)
        Dim l As New Label With {.Size = New Size(108 + x.Space + y.Space + z.Space, 29), .Location = New Point(134 - (x.Space + y.Space + z.Space) / 2, 17)}
        GroupBox3.Controls.Add(l)
        l.Refresh()
        g = l.CreateGraphics
        With g
            .DrawString("x = ", letter, brush, 0, 7)
            .DrawString(",   y = ", letter, brush, x.Space + 30, 7)
            .DrawString(",   z = ", letter, brush, x.Space + y.Space + 72, 7)
            .Dispose()
        End With
    End Sub
    '当鼠标悬停在结果框时显示详细信息
    Private Sub Info() Handles GroupBox3.MouseHover
        With ToolTip1
            If GroupBox3.Tag = "有解" Then
                If RadioButton1.Checked Then
                    .SetToolTip(GroupBox3, "x = " & x.ToString & " = " & x.Round(2) & vbCrLf & "y = " & y.ToString & " = " & y.Round(2))
                ElseIf RadioButton2.Checked Then
                    .SetToolTip(GroupBox3, "x = " & x.ToString & " = " & x.Round(2) & vbCrLf & "y = " & y.ToString & " = " & y.Round(2) & vbCrLf & "z = " & z.ToString & " = " & z.Round(2))
                End If
            Else
                .SetToolTip(GroupBox3, "")
            End If
        End With
    End Sub

#End Region

End Class
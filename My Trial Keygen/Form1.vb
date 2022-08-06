Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim DataList As New ArrayList
        Dim KeyList As New ArrayList
        Dim Pids As String
        TextBox2.Text = ""
        Dim i, j As Integer
        Dim k As Integer = 1
        Dim h As Int64
        Dim m As Int64 = 1

        Dim str As String

        If TextBox1.TextLength > 0 Then
            Pids = Trim(Replace(TextBox1.Text, "-", ""))

            For i = 1 To Len(Pids)
                str = Asc(Mid(Pids, i, 1))
                DataList.Add(str)
            Next

            For j = 0 To DataList.Count - 1
                h = CInt(DataList(j))
                m *= h
                If k Mod 4 = 0 Then
                    KeyList.Add(Hex(m / 2).ToString)
                    m = 1
                    k = 0
                End If
                k += 1
            Next

            For i = 0 To KeyList.Count - 1
                TextBox2.Text += KeyList(i)
                If i < KeyList.Count - 1 Then
                    TextBox2.Text += "-"
                End If
            Next

            TextBox2.Focus()
        Else
            MsgBox("Please Enter Product ID", MsgBoxStyle.Critical, "ID Error")
            TextBox1.Focus()

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox2.TextLength > 0 Then
            My.Computer.Clipboard.SetText(TextBox2.Text)
        Else
            MsgBox("Please Generate Key", MsgBoxStyle.Critical, "Error")
        End If

    End Sub
End Class

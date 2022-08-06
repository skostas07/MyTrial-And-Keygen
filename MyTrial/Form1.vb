Imports System.Management

Public Class Form1

    Private ProID As String = "MYTRX" '5 CHR special info
    Private Pids As String = ""
    Private Sids As String = ""
    Private RegisterKlasor As String = "Software\\" & My.Application.Info.CompanyName & "\\" & My.Application.Info.ProductName
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RegWrite("Product", My.Application.Info.ProductName)
        RegWrite("Company", My.Application.Info.CompanyName)
        RegWrite("Version", My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision)
        GenerateID()
        CheckReg()

    End Sub

    Sub CheckReg()
        Dim RegID As String = ""
        Dim RegSerial As String = ""
        RegID = RegRead("ProductID")
        RegSerial = RegRead("Serial")

        If Len(RegID) > 0 And Len(RegSerial) > 0 Then
            KeyCheck2(RegID, RegSerial)
        End If
    End Sub


    Sub KeyCheck2(RegPid As String, RegSid As String)

        Dim DataListC As New ArrayList
        Dim KeyListC As New ArrayList
        Sids = ""
        Dim i, j As Integer
        Dim k As Integer = 1
        Dim h As Int64
        Dim m As Int64 = 1
        Dim str As String
        RegPid = Replace(RegPid, "-", "")
        If Len(RegPid) > 0 Then

            For i = 1 To Len(RegPid)
                str = Asc(Mid(RegPid, i, 1))
                DataListC.Add(str)
            Next

            For j = 0 To DataListC.Count - 1
                h = CInt(DataListC(j))
                m *= h
                If k Mod 4 = 0 Then
                    KeyListC.Add(Hex(m / 2).ToString)
                    Sids += Hex(m / 2).ToString
                    m = 1
                    k = 0
                End If
                k += 1
            Next

            Dim Sidsu As String = ""
            For i = 0 To KeyListC.Count - 1
                Sidsu += KeyListC(i)
                If i < KeyListC.Count - 1 Then
                    Sidsu += "-"
                End If
            Next


            If RegSid = Sidsu Then
                Label4.Text = "Registered!"
                Label4.ForeColor = Color.LightGreen

            End If


        End If
    End Sub
    Sub KeyCheck()

        Dim DataListC As New ArrayList
        Dim KeyListC As New ArrayList
        Sids = ""
        Dim i, j As Integer
        Dim k As Integer = 1
        Dim h As Int64
        Dim m As Int64 = 1
        Dim str As String

        If Len(Pids) > 0 Then

            For i = 1 To Len(Pids)
                str = Asc(Mid(Pids, i, 1))
                DataListC.Add(str)
            Next

            For j = 0 To DataListC.Count - 1
                h = CInt(DataListC(j))
                m *= h
                If k Mod 4 = 0 Then
                    KeyListC.Add(Hex(m / 2).ToString)
                    Sids += Hex(m / 2).ToString
                    m = 1
                    k = 0
                End If
                k += 1
            Next

            Dim Sidsu As String = ""
            For i = 0 To KeyListC.Count - 1
                Sidsu += KeyListC(i)
                If i < KeyListC.Count - 1 Then
                    Sidsu += "-"
                End If
            Next


            If TextBox2.Text = Sidsu Then
                MsgBox("Lisence OK!", MsgBoxStyle.Information, My.Application.Info.ProductName)
                RegWrite("DateB", Now())

                RegWrite("Serial", Sidsu)
                Label4.Text = "Registered!"
                Label4.ForeColor = Color.LightGreen

            Else
                MsgBox("Check Serial!", MsgBoxStyle.Critical, "Serial Key Error")
                TextBox2.Focus()
            End If

        Else
            MsgBox("Product ID Error", MsgBoxStyle.Critical, "ID Error")

        End If
    End Sub
    Sub GenerateID()
        'generate product id

        'get system drive volume serial
        Dim SysCDrive, GetinfoC As String
        Dim Serials As String = ""
        Dim Err As Boolean
        SysCDrive = Environment.GetEnvironmentVariable("SystemDrive")
        Serials += ProID
        Dim Searcher As New ManagementObjectSearcher("SELECT * FROM Win32_LogicalDisk Where DeviceID = '" & SysCDrive & "'")
        For Each wmi_HD As ManagementObject In Searcher.Get()
            GetinfoC = wmi_HD("VolumeSerialNumber")
            If GetinfoC.Length > 0 Then
                Serials += GetinfoC
            Else
                Err = True
            End If
        Next

        'get hdd0 & hdd1 serial
        Dim Searcher2 As New ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive")
        Dim i As Integer = 0
        Dim HddSeri As String
        Dim HddSeri2 As String = ""
        For Each wmi_HD As ManagementObject In Searcher2.Get()
            HddSeri = Mid(Trim(wmi_HD("SerialNumber")), 1, 4)
            If i = 0 Then
                If Len(HddSeri) < 4 Then
                    HddSeri = "C321"
                End If
            ElseIf i = 1 Then
                If Len(HddSeri) < 4 Then
                    HddSeri = "D123"
                End If
            End If
            i += 1
            HddSeri2 += HddSeri
        Next

        If Len(HddSeri2) = 8 Then
            Serials += HddSeri2
        Else
            Err = True
        End If

        'get motherboard info
        Dim Searcher3 As New ManagementObjectSearcher("SELECT * FROM Win32_BaseBoard")
        Dim AnaSeri As String = ""
        i = 0
        For Each wmi_HD As ManagementObject In Searcher3.Get()
            If i = 0 Then
                AnaSeri = Trim(wmi_HD("Product"))
                AnaSeri = Replace(AnaSeri, "-", "")
                AnaSeri = Replace(AnaSeri, "/", "")
                AnaSeri = Replace(AnaSeri, "\", "")
                AnaSeri = Mid(AnaSeri, 1, 4)
            End If
            i += 1
        Next
        If Len(AnaSeri) = 4 Then
            Serials += AnaSeri
        Else
            Err = True
        End If

        'Serials += ProID
        Serials = Serials.ToUpper

        If Err = True Then
            MsgBox("Somthings Wrong!", MsgBoxStyle.Critical, "Error")
        Else
            'generate id
            Dim NumList As New ArrayList()
            Dim CalcList As New ArrayList()
            Dim CalcList2 As New ArrayList()
            Dim Str As Integer

            For i = 1 To Len(Serials)
                Str = Asc(Mid(Serials, i, 1))
                NumList.Add(Str)
            Next

            Dim j As Integer
            Dim k As Integer = 1
            Dim h As Int64
            Dim m As Int64 = 1

            For j = 0 To NumList.Count - 1
                h = CInt(NumList(j))
                m *= h
                If k Mod 4 = 0 Then
                    CalcList.Add(CInt(m / 32).ToString)
                    m = 1
                    k = 0
                End If
                k += 1
            Next

            h = 0
            m = 0
            For j = 0 To CalcList.Count - 1
                h = CalcList(j)
                m += h
                If k Mod 2 = 0 Then
                    CalcList2.Add(Hex(m / 2).ToString)
                    Pids += Hex(m / 2).ToString
                    m = 1
                    k = 0
                End If
                k += 1
            Next

            For i = 0 To CalcList2.Count - 1
                TextBox1.Text += CalcList2(i)
                If i < CalcList2.Count - 1 Then
                    TextBox1.Text += "-"
                End If
            Next

            'MsgBox(Pids)
            RegWrite("ProductID", TextBox1.Text)
            RegWrite("DateA", Now())

        End If

    End Sub
    Private Sub Form1_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        TabControl1.Size = New Size(CInt(Me.Size.Width - 40), CInt(Me.Size.Height - 80))
        Label1.Location = New Point(CInt((Me.Size.Width / 2) - Label1.Width / 2 - 20), Label1.Location.Y)
        Label2.Location = New Point(CInt((Me.Size.Width / 2) - Label2.Width / 2 - 20), Label2.Location.Y)
        Label3.Location = New Point(CInt((Me.Size.Width / 2) - Label3.Width / 2 - 20), Label3.Location.Y)
        Label4.Location = New Point(CInt((Me.Size.Width / 2) - Label4.Width / 2 - 20), Label4.Location.Y)
        TextBox1.Location = New Point(CInt((Me.Size.Width / 2) - TextBox1.Width / 2 - 20), TextBox1.Location.Y)
        TextBox2.Location = New Point(CInt((Me.Size.Width / 2) - TextBox2.Width / 2 - 20), TextBox2.Location.Y)
        Button1.Location = New Point(CInt((Me.Size.Width / 2) - Button1.Width / 2 - 20), Button1.Location.Y)
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox("Desinged by " & vbCrLf & My.Application.Info.CompanyName & vbCrLf & My.Application.Info.Description, MsgBoxStyle.Information, "My Trial About")
    End Sub

    Private Sub LisenceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LisenceToolStripMenuItem.Click
        TabControl1.SelectedTab = TabPage2
    End Sub

    Private Function RegWrite(Key As String, Value As String) As Boolean
        Dim regVersion = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegisterKlasor, True)
        If regVersion Is Nothing Then
            regVersion = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegisterKlasor)
        End If
        If regVersion IsNot Nothing Then
            regVersion.SetValue(Key, Value)
            regVersion.Close()
            RegWrite = True
        End If
    End Function
    Private Function RegRead(Key As String) As String
        Dim regVersion = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegisterKlasor, True)
        If regVersion Is Nothing Then
            RegRead = "No Key!"
        End If
        Dim intVersion As String = ""
        If regVersion IsNot Nothing Then
            intVersion = regVersion.GetValue(Key)
            regVersion.Close()
            If intVersion = "" Or intVersion Is Nothing Then
                RegRead = "No Data!"
            Else
                'RegRead = Key & " : " & intVersion
                RegRead = intVersion
            End If
        End If
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        KeyCheck()

    End Sub

    Private Sub DeleteLisenceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteLisenceToolStripMenuItem.Click
        RegWrite("Serial", "0")
        MsgBox("Lisence Deleted", MsgBoxStyle.Information, "Lisence")
        Label4.Text = "Unregistered"
        Label4.ForeColor = Color.Crimson
    End Sub
End Class

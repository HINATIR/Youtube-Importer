Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports System.Xml


Public Class Form2
    Dim lines As List(Of String) = New List(Of String)()
    Dim uName As String = Environment.UserName
    Dim appPath As String = My.Application.Info.DirectoryPath
    Dim ODIR As String = appPath + "/OutputDir.txt"


    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If File.Exists("option.xml") Then
            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load("option.xml")
            'フォルダパス
            TextBox2.Text = If(Not xmlDoc.DocumentElement.SelectSingleNode("options/folpath") Is Nothing, xmlDoc.DocumentElement.SelectSingleNode("options/folpath").InnerText, "")
            ComboBox4.SelectedIndex = Integer.Parse(If(Not xmlDoc.DocumentElement.SelectSingleNode("options/language") Is Nothing, xmlDoc.DocumentElement.SelectSingleNode("options/language").InnerText, ""))
            ComboBox3.SelectedIndex = Integer.Parse(If(Not xmlDoc.DocumentElement.SelectSingleNode("options/defab") Is Nothing, xmlDoc.DocumentElement.SelectSingleNode("options/defab").InnerText, ""))
            CheckBox3.Checked = Boolean.Parse(If(Not xmlDoc.DocumentElement.SelectSingleNode("options/hidetab") Is Nothing, xmlDoc.DocumentElement.SelectSingleNode("options/hidetab").InnerText, ""))
            Label15.Text = If(Not xmlDoc.DocumentElement.SelectSingleNode("options/update") Is Nothing, xmlDoc.DocumentElement.SelectSingleNode("options/update").InnerText, "")
        Else
            Label15.Text = ""
            Dim p2 As New Process()
            p2.StartInfo.FileName = Environment.GetEnvironmentVariable("ComSpec")
            p2.StartInfo.UseShellExecute = False
            p2.StartInfo.RedirectStandardOutput = True
            p2.StartInfo.RedirectStandardInput = False
            p2.StartInfo.CreateNoWindow = True
            p2.StartInfo.Arguments = "/c yt-dlp --version "
            p2.Start()
            Dim results2 As String = p2.StandardOutput.ReadToEnd()
            p2.WaitForExit()
            'p.Close()
            Label15.Text = results2
        End If
        Dim sr As New StreamReader(appPath & "\language\" & ComboBox4.Text & ".txt")
        Dim str As String = ""
        lines.Clear()

        Do While sr.EndOfStream = False
            lines.Add(sr.ReadLine)
        Loop
        sr.Close()


        Button1.Text = lines(18)
        Button2.Text = lines(8)
        Button3.Text = lines(6)
        Button4.Text = lines(4)
        Button5.Text = lines(5)
        Button6.Text = lines(4)
        Button7.Text = lines(5)
        Button8.Text = lines(23)
        CheckBox1.Text = lines(24)
        CheckBox2.Text = lines(25)
        CheckBox3.Text = lines(22)

        Me.ComboBox1.Items.Clear()
        With Me.ComboBox1
            .Items.Add(lines(33))
            .Items.Add(lines(34))
            .Items.Add(lines(35))
            .Items.Add(lines(36))
            .Items.Add(lines(37))
            .Items.Add(lines(38))
            .Items.Add(lines(39))
        End With

        Me.ComboBox2.Items.Clear()
        With Me.ComboBox2
            .Items.Add(lines(40))
            .Items.Add(lines(41))
            .Items.Add(lines(42))
        End With
        GroupBox1.Text = lines(9)
        Label1.Text = lines(3)
        Label2.Text = lines(15)
        Label3.Text = lines(13)
        Label4.Text = lines(7)
        Label5.Text = lines(14)

        Label7.Text = lines(16)
        Label8.Text = lines(19)
        Label9.Text = lines(21)
        Label10.Text = lines(10)
        Label11.Text = lines(11)


        Label14.Text = lines(20)
        LinkLabel1.Text = lines(17)
        TabPage1.Text = lines(0)
        TabPage2.Text = lines(1)
        TabPage3.Text = lines(2)
    End Sub
    Public Class SampleClass
        Public Number As Integer
        Public Message As String
    End Class
    Private Sub Form2_FormClosing(sender As Object, e As EventArgs) Handles MyBase.FormClosing
        Dim sw As New StreamWriter("option.xml", False,
        Encoding.GetEncoding("utf-8"))
        sw.Write("<?xml version=""1.0"" encoding=""utf-8"" ?>" & Chr(13) & "<configration>" & Chr(13) & "<options>" & Chr(13) & "<folpath>" & TextBox2.Text & "</folpath>" & Chr(13) & "<language>" & ComboBox4.SelectedIndex & "</language>" & Chr(13) & "<defab>" & ComboBox3.SelectedIndex & "</defab>" & Chr(13) & "<hidetab>" & CheckBox3.Checked & "</hidetab>" & Chr(13) & "<update>" & Label15.Text & "</update>" & Chr(13) & "</options>" & Chr(13) & "</configration>")
        sw.Close()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim fbd As New FolderBrowserDialog
        fbd.Description = lines(26)
        fbd.RootFolder = Environment.SpecialFolder.Desktop
        fbd.SelectedPath = "C:\Windows\" + uName + "\Desktop"
        If fbd.ShowDialog(Me) = DialogResult.OK Then
            TextBox2.Text = fbd.SelectedPath
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'サムネDL
        If (ComboBox2.SelectedIndex = 2) Then
            TextBox3.Text = "--write-thumbnail --skip-download --convert-thumbnails png "
        Else
            '秒数指定
            If (TextBox6.Text <> "" And TextBox7.Text <> "") Then
                TextBox3.Text = "--download-sections *" + TextBox6.Text + "-" + TextBox7.Text + " "
            ElseIf (TextBox6.Text = "" Or TextBox7.Text = "") Then
                If (TextBox6.Text = "" And TextBox7.Text = "") Then
                Else
                    MsgBox(lines(28))
                    Return
                End If
            End If
            If (ComboBox2.SelectedIndex = 0) Then
                '画質設定
                If (ComboBox1.SelectedIndex = -1) Then
                    MsgBox(lines(29))
                    Return
                End If
                If (ComboBox1.SelectedIndex = 0) Then

                End If
                If (ComboBox1.SelectedIndex = 1) Then
                    TextBox3.Text = TextBox3.Text + "-f " + Chr(34) + "bv+ba/b" + Chr(34) + " --merge-output-format mp4 "
                End If
                If (ComboBox1.SelectedIndex = 2) Then
                    TextBox3.Text = TextBox3.Text + "-f " + Chr(34) + "137+ba" + Chr(34) + " --merge-output-format mp4 "
                End If
                If (ComboBox1.SelectedIndex = 3) Then
                    TextBox3.Text = TextBox3.Text + "-f " + Chr(34) + "136+ba" + Chr(34) + " --merge-output-format mp4 "
                End If
                If (ComboBox1.SelectedIndex = 4) Then
                    TextBox3.Text = TextBox3.Text + "-f " + Chr(34) + "135+ba" + Chr(34) + " --merge-output-format mp4 "
                End If
                If (ComboBox1.SelectedIndex = 5) Then
                    TextBox3.Text = TextBox3.Text + "-f " + Chr(34) + "134+ba" + Chr(34) + " --merge-output-format mp4 "
                End If
                If (ComboBox1.SelectedIndex = 6) Then
                    TextBox3.Text = TextBox3.Text + "-f " + Chr(34) + TextBox4.Text + Chr(34) + " --merge-output-format mp4 "
                End If

            End If
        End If
        '音声DL
        If (ComboBox2.SelectedIndex = 1) Then
            TextBox3.Text = TextBox3.Text + "-x --extract-audio --audio-format mp3 "
        End If
        '保存場所
        If (TextBox2.Text = "") Then
        Else
            TextBox3.Text = TextBox3.Text + "-o " + TextBox2.Text + "\%(title)s.%(ext)s "
        End If
        '動画URLの検問
        If (TextBox1.Text = "") Then
            MsgBox(lines(30))
            Return
        End If
        '画質設定の検問
        If (ComboBox2.SelectedIndex = -1) Then
            MsgBox(lines(31))
            Return
        End If
        If (CheckBox1.Checked = True) Then
            'ブラウザ設定検問
            If (ComboBox3.SelectedIndex = -1) Then
                MsgBox(lines(32))
                Return
            End If
        End If
        If (CheckBox2.Checked = True) Then
            'クッキー用でURL開く
            Process.Start(TextBox1.Text)
        End If
        'hidetab
        Dim tabop As Integer = ProcessWindowStyle.Normal
        If (CheckBox3.Checked = True) Then
            tabop = ProcessWindowStyle.Hidden
        Else
            tabop = ProcessWindowStyle.Normal
        End If
        If (CheckBox1.Checked = True) Then
            Dim DL As New ProcessStartInfo With {.FileName = Environment.GetEnvironmentVariable("ComSpec"), .WindowStyle = tabop, .Arguments = "/c yt-dlp --cookies-from-browser " + ComboBox3.Text + " " + TextBox3.Text + TextBox1.Text
            }
            Dim p1 As Process = Process.Start(DL)
        Else
            Dim DL As New ProcessStartInfo With {.FileName = Environment.GetEnvironmentVariable("ComSpec"), .WindowStyle = tabop, .Arguments = "/c yt-dlp " + TextBox3.Text + TextBox1.Text
            }
            Dim p1 As Process = Process.Start(DL)
        End If
        TextBox3.Text = ""
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        RichTextBox1.Text = ""
        Dim p As New Process()
        p.StartInfo.FileName = Environment.GetEnvironmentVariable("ComSpec")
        p.StartInfo.UseShellExecute = False
        p.StartInfo.RedirectStandardOutput = True
        p.StartInfo.RedirectStandardInput = False
        p.StartInfo.CreateNoWindow = True
        If (TextBox5.Text = "") Then
            MsgBox(lines(27))
        End If
        p.StartInfo.Arguments = "/c yt-dlp -F " + TextBox5.Text
        p.Start()
        Dim results As String = p.StandardOutput.ReadToEnd()
        p.WaitForExit()
        'p.Close()
        RichTextBox1.Text = RichTextBox1.Text + results
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Clipboard.ContainsText() Then
            TextBox1.Text = Clipboard.GetText()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Clipboard.ContainsText() Then
            TextBox5.Text = Clipboard.GetText()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TextBox1.Text = ""
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox5.Text = ""
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If (ComboBox1.SelectedIndex = 6) Then
            TextBox4.Enabled = True
        Else
            TextBox4.Enabled = False
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If (ComboBox2.SelectedIndex = 0) Then
            ComboBox1.Enabled = True
        Else
            ComboBox1.Enabled = False
            TextBox4.Enabled = False
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("https://discord.com/invite/28DTTBJ4w7")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        RichTextBox2.Text = ""
        Dim p As New Process()
        p.StartInfo.FileName = Environment.GetEnvironmentVariable("ComSpec")
        p.StartInfo.UseShellExecute = False
        p.StartInfo.RedirectStandardOutput = True
        p.StartInfo.RedirectStandardInput = False
        p.StartInfo.CreateNoWindow = True
        p.StartInfo.Arguments = "/c yt-dlp -U"
        p.Start()
        Dim results As String = p.StandardOutput.ReadToEnd()
        p.WaitForExit()
        'p.Close()
        RichTextBox2.Text = RichTextBox2.Text + results

        Label15.Text = ""
        Dim p2 As New Process()
        p2.StartInfo.FileName = Environment.GetEnvironmentVariable("ComSpec")
        p2.StartInfo.UseShellExecute = False
        p2.StartInfo.RedirectStandardOutput = True
        p2.StartInfo.RedirectStandardInput = False
        p2.StartInfo.CreateNoWindow = True
        p2.StartInfo.Arguments = "/c yt-dlp --version "
        p2.Start()
        Dim results2 As String = p2.StandardOutput.ReadToEnd()
        p2.WaitForExit()
        'p.Close()
        Label15.Text = results2
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Process.Start("https://github.com/HINATIR/ytdlp-desktop/releases/")
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Dim sr As New StreamReader(appPath & "\language\" & ComboBox4.Text & ".txt")
        Dim str As String = ""
        lines.Clear()

        Do While sr.EndOfStream = False
            lines.Add(sr.ReadLine)
        Loop
        sr.Close()


        Button1.Text = lines(18)
        Button2.Text = lines(8)
        Button3.Text = lines(6)
        Button4.Text = lines(4)
        Button5.Text = lines(5)
        Button6.Text = lines(4)
        Button7.Text = lines(5)
        Button8.Text = lines(23)
        CheckBox1.Text = lines(24)
        CheckBox2.Text = lines(25)
        CheckBox3.Text = lines(22)

        Me.ComboBox1.Items.Clear()
        With Me.ComboBox1
            .Items.Add(lines(33))
            .Items.Add(lines(34))
            .Items.Add(lines(35))
            .Items.Add(lines(36))
            .Items.Add(lines(37))
            .Items.Add(lines(38))
            .Items.Add(lines(39))
        End With

        Me.ComboBox2.Items.Clear()
        With Me.ComboBox2
            .Items.Add(lines(40))
            .Items.Add(lines(41))
            .Items.Add(lines(42))
        End With
        GroupBox1.Text = lines(9)
        Label1.Text = lines(3)
        Label2.Text = lines(15)
        Label3.Text = lines(13)
        Label4.Text = lines(7)
        Label5.Text = lines(14)

        Label7.Text = lines(16)
        Label8.Text = lines(19)
        Label9.Text = lines(21)
        Label10.Text = lines(10)
        Label11.Text = lines(11)


        Label14.Text = lines(20)
        LinkLabel1.Text = lines(17)
        TabPage1.Text = lines(0)
        TabPage2.Text = lines(1)
        TabPage3.Text = lines(2)
    End Sub
End Class
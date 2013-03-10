Public Class settings
    Dim resolution As String
    Dim showfilenames As String
    Dim align As String
    Dim ctext As ArrayList = New ArrayList

    Private Sub settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = System.IO.File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\taskbarsettings.txt")
        Dim loaded As Array = System.IO.File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\taskbarsettings.txt").Split(Environment.NewLine)
        Dim settings As String = loaded(0)
        Dim set_ar As Array = settings.Split(";")
        resolution = set_ar(1).ToString.Split("=")(1)
        showfilenames = set_ar(0).ToString.Split("=")(1)
        align = set_ar(2).ToString.Split("=")(1)

        'init ctext array
        Dim ar As Array = TextBox1.Text.Split(vbNewLine)
        For i As Integer = 0 To ar.Length - 1
            ctext.Add(ar(i))
        Next

        If showfilenames = "true" Then
            CheckBox1.Checked = True
        Else
            CheckBox1.Checked = False
        End If
        AddHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged

        ComboBox1.SelectedIndex = ComboBox1.Items.IndexOf(resolution)
        ComboBox2.SelectedIndex = ComboBox2.Items.IndexOf(align)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        System.IO.File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\taskbarsettings.txt", TextBox1.Text)
        Application.Restart()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)
        TextBox1.Text = ""
        If showfilenames = "true" Then
            showfilenames = "false"
        Else
            showfilenames = "true"
        End If

        'update textbox1
        ctext(0) = "showfilenames=" & showfilenames.ToString.ToLower & ";resolution=" & resolution & ";align=" & align
        Dim current As Integer = 0
        For Each item In ctext
            current += 1
            If current = ctext.Count Then
                TextBox1.Text += item.ToString
            Else
                TextBox1.Text += item.ToString + vbNewLine
            End If
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox1.Text = ""
        resolution = Convert.ToInt32(ComboBox1.Items(ComboBox1.SelectedIndex).ToString)

        'update textbox1
        ctext(0) = "showfilenames=" & showfilenames.ToString.ToLower & ";resolution=" & resolution & ";align=" & align
        Dim current As Integer = 0
        For Each item In ctext
            current += 1
            If current = ctext.Count Then
                TextBox1.Text += item.ToString
            Else
                TextBox1.Text += item.ToString + vbNewLine
            End If
        Next
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        TextBox1.Text = ""
        align = ComboBox2.Items(ComboBox2.SelectedIndex).ToString

        'update textbox1
        ctext(0) = "showfilenames=" & showfilenames.ToString.ToLower & ";resolution=" & resolution & ";align=" & align
        Dim current As Integer = 0
        For Each item In ctext
            current += 1
            If current = ctext.Count Then
                TextBox1.Text += item.ToString
            Else
                TextBox1.Text += item.ToString + vbNewLine
            End If
        Next
    End Sub
End Class
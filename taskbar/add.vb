Public Class add

    Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As IntPtr, ByVal lpszExeFileName As String, ByVal nIconIndex As Integer) As IntPtr

    Dim currentimgpath As String
    Dim filepath As String
    Dim piccount As Integer

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If OpenFileDialog2.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = OpenFileDialog2.FileName
            filepath = OpenFileDialog2.FileName
            Dim hIcon As IntPtr
            hIcon = ExtractIcon(Me.Handle, filepath, 0)
            If hIcon <> 0 And hIcon <> 1 Then
                Dim ic As Icon = Icon.FromHandle(hIcon)
                Dim img As Image = ic.ToBitmap
                currentimgpath = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) & "\" & System.IO.Path.GetFileNameWithoutExtension(filepath) & ".png"
                img.Save(currentimgpath, Imaging.ImageFormat.Png)
                'save new item to taskbarsettings
                System.IO.File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\taskbarsettings.txt", vbNewLine & piccount & "#" & currentimgpath & "#" & filepath)
                Application.Restart()
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("Could not save. Choose a file please.")
        Else
            System.IO.File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\taskbarsettings.txt", vbNewLine & piccount & "#" & currentimgpath & "#" & filepath)
        End If
        Application.Restart()
    End Sub

    Private Sub add_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        piccount = System.IO.File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\taskbarsettings.txt").Split(Environment.NewLine).Length
    End Sub
End Class
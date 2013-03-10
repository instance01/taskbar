Imports System.Runtime.InteropServices
Imports System.Text
Imports System.IO

Public Class Form1
    
    Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As IntPtr, ByVal lpszExeFileName As String, ByVal nIconIndex As Integer) As IntPtr

    Dim drag As Boolean
    Dim mousex As Integer
    Dim mousey As Integer

    Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseDown
        drag = True
        mousex = Windows.Forms.Cursor.Position.X - Me.Left
        mousey = Windows.Forms.Cursor.Position.Y - Me.Top
    End Sub

    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseMove
        If drag Then
            Me.Top = Windows.Forms.Cursor.Position.Y - mousey
            Me.Left = Windows.Forms.Cursor.Position.X - mousex
        End If
    End Sub

    Private Sub Form1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseUp
        drag = False
    End Sub

    Dim piccount As Integer
    Dim timer(10) As Timer
    Dim backtimer(10) As Timer
    Dim picturebox(10) As PictureBox
    Dim deletepic(10) As PictureBox
    Dim filename(10) As Label

    'MAX 10 SUPPORTED!!

    Dim loaded_settings As ArrayList = New ArrayList

    Dim resolution As Integer = 50
    Dim showfilenames As String = "false"
    Dim align As String = "top"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Panel1.BackColor = Color.FromArgb(100, 10, 10, 10)

        'loading of settings file
        Dim loaded As Array = File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\taskbarsettings.txt").Split(Environment.NewLine)
        Dim settings As String = loaded(0)
        Dim set_ar As Array = settings.Split(";")
        resolution = Convert.ToInt32(set_ar(1).ToString.Split("=")(1))
        showfilenames = set_ar(0).ToString.Split("=")(1)
        align = set_ar(2).ToString.Split("=")(1)

        For a As Integer = 1 To loaded.Length - 1
            loaded_settings.Add(loaded(a))
        Next
        piccount = loaded_settings.Count

        Me.Height = resolution + 40
        Me.Width = (resolution + 20) * piccount + 10
        If align = "top" Then
            Me.Location = New Point(Screen.PrimaryScreen.Bounds.Width / 2 - (Me.Width / 2), 0)
        ElseIf align = "bottom" Then
            Me.Location = New Point(Screen.PrimaryScreen.Bounds.Width / 2 - (Me.Width / 2), Screen.PrimaryScreen.Bounds.Height - Me.Height)
        End If


        Dim picx As Integer = 10
        For i As Integer = 0 To piccount - 1
            'init timers
            timer(i) = New Timer
            timer(i).Interval = 20
            timer(i).Tag = i

            AddHandler timer(i).Tick, AddressOf animateTimer

            backtimer(i) = New Timer
            backtimer(i).Interval = 20
            backtimer(i).Tag = i

            AddHandler backtimer(i).Tick, AddressOf animateTimerBack

            'add pictureboxes
            picturebox(i) = New PictureBox
            picturebox(i).Location = New Point(picx + 5, 35)
            picturebox(i).Width = resolution
            picturebox(i).Height = resolution
            picturebox(i).Image = Image.FromFile(loaded_settings(i).ToString.Split("#")(1).ToString)
            picturebox(i).SizeMode = PictureBoxSizeMode.Zoom
            picturebox(i).Name = i
            picturebox(i).Cursor = Cursors.Hand
            picturebox(i).BackColor = Color.Black
            Me.Controls.Add(picturebox(i))

            AddHandler picturebox(i).Click, AddressOf picturebox_click
            AddHandler picturebox(i).MouseEnter, AddressOf picturebox_enter
            AddHandler picturebox(i).MouseLeave, AddressOf picturebox_leave

            picx += (resolution + 20)

            'add minicontrolboxes for removal of pictureboxes
            deletepic(i) = New PictureBox
            deletepic(i).Location = New Point(picturebox(i).Location.X + picturebox(i).Width, picturebox(i).Location.Y - 10)
            deletepic(i).Width = 15
            deletepic(i).Height = 15
            deletepic(i).Image = My.Resources.delete
            deletepic(i).SizeMode = PictureBoxSizeMode.Zoom
            deletepic(i).Name = i
            deletepic(i).Cursor = Cursors.Hand
            deletepic(i).BackColor = Color.Black
            Me.Controls.Add(deletepic(i))
            deletepic(i).BringToFront()

            AddHandler deletepic(i).Click, AddressOf delpicbox_click

            If showfilenames = "true" Then
                filename(i) = New Label
                filename(i).Location = New Point(picturebox(i).Location.X, picturebox(i).Location.Y + picturebox(i).Height)
                filename(i).Text = Path.GetFileNameWithoutExtension(loaded_settings(i).ToString.Split("#")(2).ToString)
                'filename(i).Width = picturebox(i).Width
                filename(i).Width = getFullTextWidth(filename(i))
                filename(i).Height = 15
                filename(i).ForeColor = Color.Orange
                filename(i).BackColor = Color.FromArgb(255, 10, 10, 10)
                Me.Controls.Add(filename(i))
                filename(i).BringToFront()
            End If
        Next

    End Sub

    Private Function getFullTextWidth(lbl As Label)
        Using g As Graphics = lbl.CreateGraphics()
            Return g.MeasureString(lbl.Text, lbl.Font).Width
        End Using
    End Function


    Private Sub delpicbox_click(sender As Object, e As EventArgs)
        Dim thispic As PictureBox = sender
        Dim current As Integer = Convert.ToInt32(thispic.Name)
        loaded_settings.RemoveAt(current)
        'save it
        Dim savestr As String = "showfilenames=" & showfilenames & ";resolution=" & resolution & ";align=" & align & vbNewLine
        Dim currentitem As Integer = 0
        For Each item In loaded_settings
            currentitem += 1
            If currentitem = loaded_settings.Count Then
                savestr += item.ToString
            Else
                savestr += item.ToString + vbNewLine
            End If
        Next
        File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\taskbarsettings.txt", savestr)
        Application.Restart()
    End Sub

    Private Sub picturebox_click(sender As Object, e As EventArgs)
        Dim thispic As PictureBox = sender
        Dim current As Integer = Convert.ToInt32(thispic.Name)
        Process.Start(loaded_settings(current).ToString.Split("#")(2).ToString)
    End Sub

    Private Sub picturebox_enter(sender As Object, e As EventArgs)
        Dim thispic As PictureBox = sender
        timer(Convert.ToInt32(thispic.Name)).Start()
    End Sub

    Private Sub picturebox_leave(sender As Object, e As EventArgs)
        Dim thispic As PictureBox = sender
        timer(Convert.ToInt32(thispic.Name)).Stop()
        backtimer(Convert.ToInt32(thispic.Name)).Start()
    End Sub

    Private Sub animateTimer(sender As Object, e As EventArgs)
        Dim thistimer As Timer = sender
        picturebox(thistimer.Tag).Width += 2
        picturebox(thistimer.Tag).Height += 2
        picturebox(thistimer.Tag).Location -= New Point(1, 1)
        If picturebox(thistimer.Tag).Width > (resolution + 10) Then
            thistimer.Stop()
        End If
    End Sub

    Private Sub animateTimerBack(sender As Object, e As EventArgs)
        Dim thistimer As Timer = sender
        picturebox(thistimer.Tag).Width -= 2
        picturebox(thistimer.Tag).Height -= 2
        picturebox(thistimer.Tag).Location += New Point(1, 1)
        If picturebox(thistimer.Tag).Width < resolution Then
            thistimer.Stop()
            picturebox(thistimer.Tag).Width = resolution
            picturebox(thistimer.Tag).Height = resolution
            picturebox(thistimer.Tag).Location -= New Point(1, 1)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        settings.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        add.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Application.Exit()
    End Sub

    'drag drop thingy
    Private Sub Form1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub
    Private Sub Form1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            Dim i As Integer
            ' Assign the files to an array. 
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            ' Loop through the array and add the files to the list. 
            For i = 0 To MyFiles.Length - 1
                If Path.GetExtension(MyFiles(i)).ToLower = ".exe" Then
                    'extract image from exe and save it
                    Dim hIcon As IntPtr
                    Dim currentimgpath As String
                    hIcon = ExtractIcon(Me.Handle, MyFiles(i), 0)
                    If hIcon <> 0 And hIcon <> 1 Then
                        Dim ic As Icon = Icon.FromHandle(hIcon)
                        Dim img As Image = ic.ToBitmap
                        currentimgpath = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) & "\" & Path.GetFileNameWithoutExtension(MyFiles(i)) & ".png"
                        img.Save(currentimgpath, Imaging.ImageFormat.Png)
                        'save new item to taskbarsettings
                        File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\taskbarsettings.txt", vbNewLine & piccount & "#" & currentimgpath & "#" & MyFiles(i))
                        Application.Restart()
                    End If
                ElseIf Path.GetExtension(MyFiles(i)).ToLower = ".lnk" Then
                    Dim TargetFilename As String = ResolveShortcut(MyFiles(i))
                    Dim hIcon As IntPtr
                    Dim currentimgpath As String
                    hIcon = ExtractIcon(Me.Handle, TargetFilename, 0)
                    If hIcon <> 0 And hIcon <> 1 Then
                        Dim ic As Icon = Icon.FromHandle(hIcon)
                        Dim img As Image = ic.ToBitmap
                        currentimgpath = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) & "\" & Path.GetFileNameWithoutExtension(MyFiles(i)) & ".png"
                        img.Save(currentimgpath, Imaging.ImageFormat.Png)
                        'save new item to taskbarsettings
                        File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\taskbarsettings.txt", vbNewLine & piccount & "#" & currentimgpath & "#" & TargetFilename)
                        Application.Restart()
                    End If
                Else
                    MsgBox("This is not a valid executable or link.")
                    Return
                End If
            Next
        End If
    End Sub




    'Thats not by Rnetwork
    'heres the link: http://stackoverflow.com/questions/13511897/getting-the-full-path-of-a-lnk-file


    <DllImport("shfolder.dll", CharSet:=CharSet.Auto)>
    Friend Shared Function SHGetFolderPath(hwndOwner As IntPtr, nFolder As Integer, hToken As IntPtr, dwFlags As Integer, lpszPath As StringBuilder) As Integer
    End Function

    <Flags()>
    Private Enum SLGP_FLAGS
        ''' <summary>Retrieves the standard short (8.3 format) file name</summary>
        SLGP_SHORTPATH = &H1
        ''' <summary>Retrieves the Universal Naming Convention (UNC) path name of the file</summary>
        SLGP_UNCPRIORITY = &H2
        ''' <summary>Retrieves the raw path name. A raw path is something that might not exist and may include environment variables that need to be expanded</summary>
        SLGP_RAWPATH = &H4
    End Enum

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Private Structure WIN32_FIND_DATAW
        Public dwFileAttributes As UInteger
        Public ftCreationTime As Long
        Public ftLastAccessTime As Long
        Public ftLastWriteTime As Long
        Public nFileSizeHigh As UInteger
        Public nFileSizeLow As UInteger
        Public dwReserved0 As UInteger
        Public dwReserved1 As UInteger
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public cFileName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=14)>
        Public cAlternateFileName As String
    End Structure

    <Flags()>
    Private Enum SLR_FLAGS
        ''' <summary>
        ''' Do not display a dialog box if the link cannot be resolved. When SLR_NO_UI is set,
        ''' the high-order word of fFlags can be set to a time-out value that specifies the
        ''' maximum amount of time to be spent resolving the link. The function returns if the
        ''' link cannot be resolved within the time-out duration. If the high-order word is set
        ''' to zero, the time-out duration will be set to the default value of 3,000 milliseconds
        ''' (3 seconds). To specify a value, set the high word of fFlags to the desired time-out
        ''' duration, in milliseconds.
        ''' </summary>
        SLR_NO_UI = &H1
        ''' <summary>Obsolete and no longer used</summary>
        SLR_ANY_MATCH = &H2
        ''' <summary>If the link object has changed, update its path and list of identifiers.
        ''' If SLR_UPDATE is set, you do not need to call IPersistFile::IsDirty to determine
        ''' whether or not the link object has changed.</summary>
        SLR_UPDATE = &H4
        ''' <summary>Do not update the link information</summary>
        SLR_NOUPDATE = &H8
        ''' <summary>Do not execute the search heuristics</summary>
        SLR_NOSEARCH = &H10
        ''' <summary>Do not use distributed link tracking</summary>
        SLR_NOTRACK = &H20
        ''' <summary>Disable distributed link tracking. By default, distributed link tracking tracks
        ''' removable media across multiple devices based on the volume name. It also uses the
        ''' Universal Naming Convention (UNC) path to track remote file systems whose drive letter
        ''' has changed. Setting SLR_NOLINKINFO disables both types of tracking.</summary>
        SLR_NOLINKINFO = &H40
        ''' <summary>Call the Microsoft Windows Installer</summary>
        SLR_INVOKE_MSI = &H80
    End Enum

    ''' <summary>The IShellLink interface allows Shell links to be created, modified, and resolved</summary>
    <ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214F9-0000-0000-C000-000000000046")>
    Private Interface IShellLinkW
        ''' <summary>Retrieves the path and file name of a Shell link object</summary>
        Sub GetPath(<Out(), MarshalAs(UnmanagedType.LPWStr)> pszFile As StringBuilder, cchMaxPath As Integer, ByRef pfd As WIN32_FIND_DATAW, fFlags As SLGP_FLAGS)
        ''' <summary>Retrieves the list of item identifiers for a Shell link object</summary>
        Sub GetIDList(ByRef ppidl As IntPtr)
        ''' <summary>Sets the pointer to an item identifier list (PIDL) for a Shell link object.</summary>
        Sub SetIDList(pidl As IntPtr)
        ''' <summary>Retrieves the description string for a Shell link object</summary>
        Sub GetDescription(<Out(), MarshalAs(UnmanagedType.LPWStr)> pszName As StringBuilder, cchMaxName As Integer)
        ''' <summary>Sets the description for a Shell link object. The description can be any application-defined string</summary>
        Sub SetDescription(<MarshalAs(UnmanagedType.LPWStr)> pszName As String)
        ''' <summary>Retrieves the name of the working directory for a Shell link object</summary>
        Sub GetWorkingDirectory(<Out(), MarshalAs(UnmanagedType.LPWStr)> pszDir As StringBuilder, cchMaxPath As Integer)
        ''' <summary>Sets the name of the working directory for a Shell link object</summary>
        Sub SetWorkingDirectory(<MarshalAs(UnmanagedType.LPWStr)> pszDir As String)
        ''' <summary>Retrieves the command-line arguments associated with a Shell link object</summary>
        Sub GetArguments(<Out(), MarshalAs(UnmanagedType.LPWStr)> pszArgs As StringBuilder, cchMaxPath As Integer)
        ''' <summary>Sets the command-line arguments for a Shell link object</summary>
        Sub SetArguments(<MarshalAs(UnmanagedType.LPWStr)> pszArgs As String)
        ''' <summary>Retrieves the hot key for a Shell link object</summary>
        Sub GetHotkey(ByRef pwHotkey As Short)
        ''' <summary>Sets a hot key for a Shell link object</summary>
        Sub SetHotkey(wHotkey As Short)
        ''' <summary>Retrieves the show command for a Shell link object</summary>
        Sub GetShowCmd(ByRef piShowCmd As Integer)
        ''' <summary>Sets the show command for a Shell link object. The show command sets the initial show state of the window.</summary>
        Sub SetShowCmd(iShowCmd As Integer)
        ''' <summary>Retrieves the location (path and index) of the icon for a Shell link object</summary>
        Sub GetIconLocation(<Out(), MarshalAs(UnmanagedType.LPWStr)> pszIconPath As StringBuilder, cchIconPath As Integer, ByRef piIcon As Integer)
        ''' <summary>Sets the location (path and index) of the icon for a Shell link object</summary>
        Sub SetIconLocation(<MarshalAs(UnmanagedType.LPWStr)> pszIconPath As String, iIcon As Integer)
        ''' <summary>Sets the relative path to the Shell link object</summary>
        Sub SetRelativePath(<MarshalAs(UnmanagedType.LPWStr)> pszPathRel As String, dwReserved As Integer)
        ''' <summary>Attempts to find the target of a Shell link, even if it has been moved or renamed</summary>
        Sub Resolve(hwnd As IntPtr, fFlags As SLR_FLAGS)
        ''' <summary>Sets the path and file name of a Shell link object</summary>
        Sub SetPath(<MarshalAs(UnmanagedType.LPWStr)> pszFile As String)

    End Interface

    <ComImport(), Guid("0000010c-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IPersist
        <PreserveSig()>
        Sub GetClassID(ByRef pClassID As Guid)
    End Interface


    <ComImport(), Guid("0000010b-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IPersistFile
        Inherits IPersist
        Shadows Sub GetClassID(ByRef pClassID As Guid)
        <PreserveSig()>
        Function IsDirty() As Integer

        <PreserveSig()>
        Sub Load(<[In](), MarshalAs(UnmanagedType.LPWStr)> pszFileName As String, dwMode As UInteger)

        <PreserveSig()>
        Sub Save(<[In](), MarshalAs(UnmanagedType.LPWStr)> pszFileName As String, <[In](), MarshalAs(UnmanagedType.Bool)> fRemember As Boolean)

        <PreserveSig()>
        Sub SaveCompleted(<[In](), MarshalAs(UnmanagedType.LPWStr)> pszFileName As String)

        <PreserveSig()>
        Sub GetCurFile(<[In](), MarshalAs(UnmanagedType.LPWStr)> ppszFileName As String)
    End Interface

    Const STGM_READ As UInteger = 0
    Const MAX_PATH As Integer = 260

    ' CLSID_ShellLink from ShlGuid.h 
    <ComImport(), Guid("00021401-0000-0000-C000-000000000046")> Public Class ShellLink
    End Class


    Public Shared Function ResolveShortcut(filename As String) As String
        Dim link As New ShellLink()
        DirectCast(link, IPersistFile).Load(filename, STGM_READ)
        ' TODO: if I can get hold of the hwnd call resolve first. This handles moved and renamed files.  
        ' ((IShellLinkW)link).Resolve(hwnd, 0) 
        Dim sb As New StringBuilder(MAX_PATH)
        Dim data As New WIN32_FIND_DATAW()
        DirectCast(link, IShellLinkW).GetPath(sb, sb.Capacity, data, 0)
        Return sb.ToString()
    End Function
End Class

Imports System.Runtime.InteropServices.Marshal


Public Class Splash_Screen
    Inherits System.Windows.Forms.Form
    
#Region " Windows Form Designer generated code "

    Public Sub New(ByVal heading_input As String, ByVal displayFile As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        heading_string = heading_input
        display_file = displayFile
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Decline_Button As System.Windows.Forms.Button
    Friend WithEvents Accept_Button As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents heading As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Splash_Screen))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.heading = New System.Windows.Forms.Label
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.Decline_Button = New System.Windows.Forms.Button
        Me.Accept_Button = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Panel1.BackgroundImage = CType(resources.GetObject("Panel1.BackgroundImage"), System.Drawing.Image)
        Me.Panel1.Controls.Add(Me.heading)
        Me.Panel1.Controls.Add(Me.RichTextBox1)
        Me.Panel1.Controls.Add(Me.Decline_Button)
        Me.Panel1.Controls.Add(Me.Accept_Button)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Location = New System.Drawing.Point(58, 86)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(524, 308)
        Me.Panel1.TabIndex = 4
        '
        'heading
        '
        Me.heading.BackColor = System.Drawing.Color.Transparent
        Me.heading.Font = New System.Drawing.Font("Arial", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.heading.ForeColor = System.Drawing.Color.Black
        Me.heading.Location = New System.Drawing.Point(8, 16)
        Me.heading.Name = "heading"
        Me.heading.Size = New System.Drawing.Size(384, 40)
        Me.heading.TabIndex = 8
        Me.heading.Text = "Heading"
        Me.heading.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.Color.White
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox1.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.RichTextBox1.ForeColor = System.Drawing.Color.Black
        Me.RichTextBox1.Location = New System.Drawing.Point(32, 70)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(464, 174)
        Me.RichTextBox1.TabIndex = 1
        Me.RichTextBox1.Text = ""
        '
        'Decline_Button
        '
        Me.Decline_Button.BackColor = System.Drawing.Color.Gainsboro
        Me.Decline_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Decline_Button.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Decline_Button.ForeColor = System.Drawing.Color.Black
        Me.Decline_Button.Location = New System.Drawing.Point(432, 276)
        Me.Decline_Button.Name = "Decline_Button"
        Me.Decline_Button.Size = New System.Drawing.Size(64, 22)
        Me.Decline_Button.TabIndex = 3
        Me.Decline_Button.Text = "No"
        '
        'Accept_Button
        '
        Me.Accept_Button.BackColor = System.Drawing.Color.Gainsboro
        Me.Accept_Button.Enabled = False
        Me.Accept_Button.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Accept_Button.ForeColor = System.Drawing.Color.Black
        Me.Accept_Button.Location = New System.Drawing.Point(360, 276)
        Me.Accept_Button.Name = "Accept_Button"
        Me.Accept_Button.Size = New System.Drawing.Size(64, 22)
        Me.Accept_Button.TabIndex = 2
        Me.Accept_Button.Text = "Yes"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(30, 282)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(332, 16)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "If you agree to the above, click 'Yes' to login. Click 'No' to logout."
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Gray
        Me.Label3.Location = New System.Drawing.Point(63, 255)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(408, 16)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Note: You need to scroll to the bottom of the text before the 'Yes' option become" & _
        "s enabled."
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(216, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(256, 48)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Label2"
        Me.Label2.Visible = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(104, 32)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Label4"
        Me.Label4.Visible = False
        '
        'Splash_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.ClientSize = New System.Drawing.Size(640, 480)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(640, 480)
        Me.Name = "Splash_Screen"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login Splash Screen 20060121.1"
        Me.TopMost = True
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Public display_file As String
    Public heading_string As String
    Private maxvscroll As Integer = 0


    ' Scrollbar direction
    '
    Const SBS_HORZ = 0
    Const SBS_VERT = 1


    ' Windows Messages
    '
    Const WM_VSCROLL = &H115
    Const WM_HSCROLL = &H114
    Const SB_THUMBPOSITION = 4



    Private Declare Function GetScrollPos Lib "user32.dll" ( _
               ByVal hWnd As IntPtr, _
               ByVal nBar As Integer) As Integer

    'Example: position = GetScrollPos(textbox1.handle, SBS_HORZ)

    Private Declare Function SetScrollPos Lib "user32.dll" ( _
                ByVal hWnd As IntPtr, _
                ByVal nBar As Integer, _
                ByVal nPos As Integer, _
                ByVal bRedraw As Boolean) As Integer

    'Example: SetScrollPos(hWnd, SBS_HORZ, position, True

    Private Declare Function PostMessageA Lib "user32.dll" ( _
               ByVal hwnd As IntPtr, _
               ByVal wMsg As Integer, _
               ByVal wParam As Integer, _
               ByVal lParam As Integer) As Boolean

    'Example: PostMessageA(hWnd, WM_HSCROLL, SB_THUMBPOSITION _
    '                         + &H10000 * position, Nothing)

    Private Sub Decline_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Decline_Button.Click
        Try
            Me.Hide()
            Dim ProcID As Integer
            ProcID = Shell("shutdown -l -f", AppWinStyle.NormalFocus, True, 30000)
            Application.Exit()
        Catch ex As Exception
            Application.Exit()
        End Try
    End Sub

    Private Sub Accept_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Accept_Button.Click
        Try
            Me.Hide()
            Application.Exit()
        Catch ex As Exception
            Application.Exit()
        End Try
    End Sub

    Private Sub Splash_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            RichTextBox1.Text = "Loading Message File from Server..."
            Try
                heading.Text = heading_string
                Try
                    RichTextBox1.LoadFile(display_file)
                Catch et As Exception
                    RichTextBox1.Text = "Welcome to " & heading_string & "."
                End Try
                If RichTextBox1.TextLength <= 0 Then
                    RichTextBox1.Text = "Welcome to " & heading_string & "."
                End If

                RichTextBox1.Focus()
                RichTextBox1.Select()
                Dim numberoflines As Long = RichTextBox1.Lines.LongLength
                RichTextBox1.Select(RichTextBox1.TextLength() - 2, 1)
                Dim Position = GetScrollPos(RichTextBox1.Handle, SBS_VERT)
                Label2.Text = ""
                Label2.Text = Label2.Text & "Number of chars: " & RichTextBox1.TextLength & vbCrLf
                Label2.Text = Label2.Text & "Number of lines: " & RichTextBox1.Lines.LongLength & vbCrLf
                Label2.Text = Label2.Text & "Scroll Position: " & Position
                maxvscroll = Position
                maxvscroll = maxvscroll - 2
                If maxvscroll < 0 Then
                    maxvscroll = 0
                End If
                Label4.Text = "Maxvscroll: " & maxvscroll
                RichTextBox1.Select(0, 0)
                Accept_Button.Enabled = False
                Position = GetScrollPos(RichTextBox1.Handle, SBS_VERT)
                If Position >= maxvscroll Then
                    Accept_Button.Enabled = True
                    Label3.Visible = False
                End If


            Catch et As Exception
                    RichTextBox1.Text = "Welcome to " & heading_string & "."
                End Try
            Catch ex As Exception
            Application.Exit()
        End Try
    End Sub

    Private Sub handlerichtextbox1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RichTextBox1.KeyDown
        Dim Position = GetScrollPos(RichTextBox1.Handle, SBS_VERT)
        If Position >= maxvscroll Then
            Accept_Button.Enabled = True
        End If
        Label2.Text = "keyDown: " & Position
    End Sub

    Private Sub aa(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove, Panel1.MouseMove, RichTextBox1.MouseMove
        Dim Position = GetScrollPos(RichTextBox1.Handle, SBS_VERT)
        If Position >= maxvscroll Then
            Accept_Button.Enabled = True
        End If
        Label2.Text = "MouseMove: " & Position
    End Sub

    Private Sub aa1(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.MouseEnter, Panel1.MouseEnter, RichTextBox1.MouseEnter
        Dim Position = GetScrollPos(RichTextBox1.Handle, SBS_VERT)
        If Position >= maxvscroll Then
            Accept_Button.Enabled = True
        End If
        Label2.Text = "MouseEnter: " & Position
    End Sub
    Private Sub aa2(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.MouseLeave, Panel1.MouseLeave, RichTextBox1.MouseLeave
        Dim Position = GetScrollPos(RichTextBox1.Handle, SBS_VERT)
        If Position >= maxvscroll Then
            Accept_Button.Enabled = True
        End If
        Label2.Text = "MouseLeave: " & Position
    End Sub

    Private Sub aa3(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown, Panel1.MouseDown, RichTextBox1.MouseDown
        Dim Position = GetScrollPos(RichTextBox1.Handle, SBS_VERT)
        If Position >= maxvscroll Then
            Accept_Button.Enabled = True
        End If
        Label2.Text = "MouseDown: " & Position
    End Sub

    Private Sub aa4(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.MouseHover, Panel1.MouseHover, RichTextBox1.MouseHover
        Dim Position = GetScrollPos(RichTextBox1.Handle, SBS_VERT)
        If Position >= maxvscroll Then
            Accept_Button.Enabled = True
        End If
        Label2.Text = "MouseHover: " & Position
    End Sub
    Private Sub aa5(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp, Panel1.MouseUp, RichTextBox1.MouseUp
        Dim Position = GetScrollPos(RichTextBox1.Handle, SBS_VERT)
        If Position >= maxvscroll Then
            Accept_Button.Enabled = True
        End If
        Label2.Text = "MouseUp: " & Position
    End Sub

    Private Sub handlerichtextbox(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox1.VScroll
        Dim Position = GetScrollPos(RichTextBox1.Handle, SBS_VERT)
        If Position >= maxvscroll Then
            Accept_Button.Enabled = True
        End If
        Label2.Text = "VScroll: " & Position
    End Sub
    Private Sub handleacceptbuttonenabled(ByVal sender As Object, ByVal e As System.EventArgs) Handles Accept_Button.EnabledChanged
        If Accept_Button.Enabled = False Then
            Accept_Button.BackColor = System.Drawing.Color.WhiteSmoke
        End If
        If Accept_Button.Enabled = True Then
            Accept_Button.BackColor = System.Drawing.Color.Gainsboro
            Accept_Button.ForeColor = System.Drawing.Color.Black
        End If
    End Sub
End Class

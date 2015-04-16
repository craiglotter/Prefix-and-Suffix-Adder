Public Class Form1
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker
    Public Delegate Sub WorkerCompletehHandler(ByVal Result As String)
    Public Delegate Sub WorkerProgresshHandler(ByVal Result As String)
    Public Delegate Sub WorkerAlerthHandler(ByVal Result As String)

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerCompleteHandler
        AddHandler Worker1.WorkerProgress, AddressOf WorkerProgressHandler
        AddHandler Worker1.WorkerAlert, AddressOf WorkerAlertHandler
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
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.BackColor = System.Drawing.Color.LimeGreen
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.Cursor = System.Windows.Forms.Cursors.Default
        Me.TextBox1.Location = New System.Drawing.Point(16, 88)
        Me.TextBox1.MaxLength = 3276700
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox1.Size = New System.Drawing.Size(720, 192)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Text = ""
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "txt"
        Me.OpenFileDialog1.Filter = "Text Files|*.txt|All Files|*.*"
        Me.OpenFileDialog1.Title = "Select the text file to be manipulated"
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.BackColor = System.Drawing.Color.LimeGreen
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Location = New System.Drawing.Point(632, 56)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(104, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Process File"
        '
        'Label1
        '
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(600, 40)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Prefix and Suffix Adder is a tool used to add prefixes and suffixes to each line " & _
        "of a given text file. All that is required is a valid text file and an inputted " & _
        "prefix or suffix. To launch the process, simply click on the ""Process File"" butt" & _
        "on."
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.Location = New System.Drawing.Point(672, 2)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "| About |"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.Location = New System.Drawing.Point(722, 2)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 16)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Exit |"
        '
        'Label4
        '
        Me.Label4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.Location = New System.Drawing.Point(16, 288)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(720, 16)
        Me.Label4.TabIndex = 5
        '
        'TextBox2
        '
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox2.Location = New System.Drawing.Point(88, 58)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(222, 20)
        Me.TextBox2.TabIndex = 6
        Me.TextBox2.Text = ""
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 58)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 16)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Prefix to add:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(320, 58)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(74, 16)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "Suffix to add:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBox3
        '
        Me.TextBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox3.Location = New System.Drawing.Point(392, 58)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(224, 20)
        Me.TextBox3.TabIndex = 10
        Me.TextBox3.Text = ""
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.PaleGreen
        Me.ClientSize = New System.Drawing.Size(752, 310)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label3)
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "Prefix and Suffix Adder 1.0"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As String)
        Try
            Dim Display_Message1 As New Display_Message(message)
            Display_Message1.ShowDialog()
            TextBox1.Clear()
            TextBox1.Text = message
            TextBox1.Refresh()
        Catch ex As Exception
            MsgBox("An error occurred in Prefix and Suffix Adder's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Public Sub WorkerCompleteHandler(ByVal Result As String)
        Try
            Button1.Enabled = True
            Label4.Text = Label4.Text.Replace("Processing...", "Complete")
        Catch ex As Exception
            Error_Handler(ex.ToString)
        End Try
    End Sub

    Public Sub WorkerAlertHandler(ByVal Result As String)
        Try
            TextBox1.Text = TextBox1.Text & Result
            TextBox1.Select(TextBox1.Text.Length, 0)
            TextBox1.ScrollToCaret()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Public Sub WorkerProgressHandler(ByVal value As Integer)
        Try
            Label4.Text = "Processing... (" & value & " lines parsed)"
            ' Label4.Refresh()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub run_worker(ByVal filelist As String)
        Try
            Label4.Text = ""
            Label4.Refresh()
            TextBox1.Clear()
            TextBox1.Refresh()
            Worker1.prefix = TextBox2.Text.Trim
            Worker1.suffix = TextBox3.Text.Trim
            Worker1.filelist_file = filelist.Trim
            Worker1.ChooseThreads(1)
            Button1.Enabled = False
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Dim result As DialogResult
            result = OpenFileDialog1.ShowDialog()
            If result = DialogResult.OK Then
                run_worker(OpenFileDialog1.FileName)
            End If
            OpenFileDialog1.Dispose()
        Catch ex As Exception
            Error_Handler(ex.ToString)
        End Try
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        Application.ExitThread()
        Application.Exit()
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        Label4.Text = ""
        Label4.Refresh()
        TextBox1.Clear()
        TextBox1.Text = TextBox1.Text & "------------------------------------------" & vbCrLf
        TextBox1.Text = TextBox1.Text & "Prefix and Suffix Adder 1.0" & vbCrLf
        TextBox1.Text = TextBox1.Text & "------------------------------------------" & vbCrLf
        TextBox1.Text = TextBox1.Text & "Version: 1.0" & vbCrLf
        TextBox1.Text = TextBox1.Text & "Build: 1.0.1" & vbCrLf
        TextBox1.Text = TextBox1.Text & "Release: 09/05/2005" & vbCrLf
        TextBox1.Text = TextBox1.Text & "Developer: Craig G. Lotter" & vbCrLf
        TextBox1.Text = TextBox1.Text & "------------------------------------------" & vbCrLf
        TextBox1.Text = TextBox1.Text & "" & vbCrLf
        TextBox1.Text = TextBox1.Text & "Change log: 09/05/2005" & vbCrLf
        TextBox1.Text = TextBox1.Text & "Special Replace Character Sequence Added: @LINE@ will cause file line number to be inserted in either prefix or suffix." & vbCrLf
        TextBox1.Refresh()
    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Application.ExitThread()
        Application.Exit()
    End Sub
End Class

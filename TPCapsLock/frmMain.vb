'
'
'
''Title: Display/Show Text or Numbers in the System Tray/Notify Icon
'
''By: Jason Hensley
'
''Contact: mailto:vbcodesource@gmail.com
'
''Webpage: http://www.vbcodesource.com|org and http://www.vbforfree.com
'
''Description: A simple example of how to display text and/or numbers in the windows system tray;
'notify icon area. This example creates a bitmap object and draws the string to the bitmap that
'you want to display and then create a icon from the bitmap and set the NotifyIcon's icon to the
'icon that was created from the bitmap. All there is to it :) Thanks to The Code Project for the
'idea.
'
''Copyright: 2007, November 15th
'
'
'
Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        Me.WindowState = FormWindowState.Minimized
        Me.Visible = False
        Label1.Text = "Just to display CapsLock" & vbCrLf & "Modified by nineapple"
        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents notifyText As System.Windows.Forms.NotifyIcon
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents cMnu As System.Windows.Forms.ContextMenu
    Friend WithEvents Timer1 As Timer
    Friend WithEvents Label1 As Label
    Friend WithEvents cMnuExit As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.notifyText = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.cMnu = New System.Windows.Forms.ContextMenu()
        Me.cMnuExit = New System.Windows.Forms.MenuItem()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'notifyText
        '
        Me.notifyText.ContextMenu = Me.cMnu
        Me.notifyText.Text = "Caps"
        Me.notifyText.Visible = True
        '
        'cMnu
        '
        Me.cMnu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.cMnuExit})
        '
        'cMnuExit
        '
        Me.cMnuExit.Index = 0
        Me.cMnuExit.Text = "Exit"
        '
        'btnRemove
        '
        Me.btnRemove.Cursor = System.Windows.Forms.Cursors.PanNW
        Me.btnRemove.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnRemove.Font = New System.Drawing.Font("Papyrus", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRemove.Location = New System.Drawing.Point(14, 10)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(192, 34)
        Me.btnRemove.TabIndex = 1
        Me.btnRemove.Text = "Remove CapsLock Icon"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 200
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(18, 57)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(281, 12)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Just to display CapsLock status at System Tray"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.ClientSize = New System.Drawing.Size(427, 254)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnRemove)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CapsLock By Nineapple"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    '
    '
    '
    'I'm using this in the createTextIcon sub to keep the icons from piling up in memory and
    'releases all of the resources that they would have used.
    Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Int32) As Int32
    '
    'You should fine tune the font you want to use so the user can see the text you want them to
    'see. Certain Fonts will obviously display your text better than other fonts might.
    Dim fontToUse As Font = New Font("Microsoft Sans Serif", 18, FontStyle.Bold,
        GraphicsUnit.Pixel)
    Dim fontToUseN As Font = New Font("Microsoft Sans Serif", 10, FontStyle.Bold,
        GraphicsUnit.Pixel)
    Dim fontToUse2 As Font = New Font("Microsoft Sans Serif", 24, FontStyle.Regular,
        GraphicsUnit.Pixel)
    '
    'A basic brush with a Dark Blue Color. This should show up pretty well in the icon tray if the user
    'uses the default tray color.
    Dim brushToUse As Brush = New SolidBrush(Color.Black)
    '
    'A bitmap used to setup how the icon will display.
    Dim bitmapText As Bitmap = New Bitmap(16, 16)

    '
    'A simply Grahics object to draw the text to a bitmap.
    Dim g As Graphics = Drawing.Graphics.FromImage(bitmapText)
    '

    'Will be used to get the Handle to the bitmap Icon.
    Dim hIcon As IntPtr
    '
    '
    'This code was actually taken from my CPUMonLite Application which displays the processors
    'clockspeed in the system tray. But the clockspeed code is removed :)
    Sub createTextIcon()
        '
        Try
            '
            'I the display icon text idea from i'm thinking the Code Project website.
            '
            'Clear the previous 'stuff' instead of creating a new bitmap.
            g.Clear(Color.White)
            '
            '
            'Setup the text, font, brush, and position for the system tray icon. For the font
            'type and size I used, a good position for the X coordinate is a -1 or -2. And the Y
            'coordinate seems to work well at a 5.
            '
            'You specify the actual text you want to be displayed in the draw string parameter that
            'you want to display in the notify area of the system tray. You will only be able to
            'display a few characters, depending on the font, size of the font, and the coordinates
            'you used.
            If Control.IsKeyLocked(Keys.CapsLock) Then
                If Control.IsKeyLocked(Keys.NumLock) Then
                    g.DrawString("A", fontToUse, brushToUse, -1, -4)
                    g.DrawString("1", fontToUseN, brushToUse, -4, -2)
                    notifyText.Text = "Caps Lock is ON" & vbCrLf & "Num Lock is ON"
                Else
                    g.DrawString("A", fontToUse, brushToUse, -1, -4)
                    notifyText.Text = "Caps Lock is ON" & vbCrLf & "Num Lock is OFF"
                End If

                'g.DrawString("A", fontToUse, brushToUse, 0, -3)

            Else
                If Control.IsKeyLocked(Keys.NumLock) Then
                    g.DrawString("a", fontToUse2, brushToUse, -2, -9)
                    g.DrawString("1", fontToUseN, brushToUse, -4, -2)
                    notifyText.Text = "Caps Lock is OFF" & vbCrLf & "Num Lock is ON"
                Else
                    g.DrawString("a", fontToUse2, brushToUse, -2, -9)
                    notifyText.Text = "Caps Lock is OFF" & vbCrLf & "Num Lock is OFF"
                End If

                'g.DrawString("a", fontToUse2, brushToUse, -1, -8)


            End If

            '
            'Get a handle to the bitmap as a Icon.
            hIcon = (bitmapText.GetHicon)
            '
            'Display that new usage value image in the system tray.
            notifyText.Icon = Drawing.Icon.FromHandle(hIcon)
            '
            'Added this to try and get away from a rare Generic Error from the code above.
            'Using this API Function seems to have stopped that error from happening. The error
            'would only occur when I made the SysMonLite program which called these codes every
            'second. So adding this API fixed that problem. I went ahead and carried it over
            'to this example just in case. :)
            DestroyIcon(hIcon.ToInt32)

        Catch exc As Exception

            MessageBox.Show(exc.InnerException.ToString, "Somethings not right?",
                MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '
        'Call the Sub that contains the code for system tray text.
        createTextIcon()

    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        '
        'This is remove the icon from the system tray.
        notifyText.Icon = Nothing

    End Sub

    Private Sub cMnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cMnuExit.Click
        '
        Me.Close()

    End Sub

    Private Sub frmMain_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        '
        'Clean up
        notifyText.Icon = Nothing
        notifyText.Dispose()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        createTextIcon()
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

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
    Friend WithEvents TextControl1 As TXTextControl.TextControl
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.TextControl1 = New TXTextControl.TextControl()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.SuspendLayout()
        '
        'TextControl1
        '
        Me.TextControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextControl1.Font = New System.Drawing.Font("Arial", 10.0!)
        Me.TextControl1.FormattingPrinter = "Standard"
        Me.TextControl1.Location = New System.Drawing.Point(0, 0)
        Me.TextControl1.Name = "TextControl1"
        Me.TextControl1.Size = New System.Drawing.Size(696, 413)
        Me.TextControl1.TabIndex = 0
        Me.TextControl1.Text = "TextControl1"
        Me.TextControl1.UserNames = Nothing
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Insert Image"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 19)
        Me.ClientSize = New System.Drawing.Size(696, 413)
        Me.Controls.Add(Me.TextControl1)
        Me.Menu = Me.MainMenu1
        Me.Name = "Form1"
        Me.Text = "TX Sample: Insert Images from Memory"
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        Dim myImage As System.Drawing.Image = System.Drawing.Image.FromFile("Blue hills.jpg")
        InsertImageFromMemory(myImage, False, TextControl1)
    End Sub

    Private Sub InsertImageFromMemory(ByVal image As System.Drawing.Image, ByVal AsFixedImage As Boolean, ByVal tx As TXTextControl.TextControl)
        Dim stream As New System.IO.MemoryStream
        Dim buffer() As Byte
        Dim RTFString As String
        Dim ImageString As New System.Text.StringBuilder

        image.Save(stream, System.Drawing.Imaging.ImageFormat.Jpeg)
        buffer = stream.ToArray()

        For Each nbyte As Byte In buffer
            ImageString.Append(nbyte.ToString("x2"))
        Next

        RTFString = "{\rtf1" + vbCr

        If AsFixedImage = True Then RTFString += "{\shp{\*\shpinst\shpfhdr0\shpbxcolumn\shpbypara\shpwr2\shpwrk0\shpfblwtxt0\shplid1025{\sp{\sn shapeType}{\sv 75}}{\sp{\sn dxWrapDistLeft}{\sv 0}}{\sp{\sn dxWrapDistRight}{\sv 0}}{\sp{\sn pib}{\sv "

        RTFString += "{\pict\jpegblip\picscalex100\picscaley100" + vbCr
        RTFString += ImageString.ToString()
        RTFString += "}}"

        If AsFixedImage = True Then RTFString += "\par}"

        tx.Selection.Load(RTFString, TXTextControl.StringStreamType.RichTextFormat)
    End Sub
End Class

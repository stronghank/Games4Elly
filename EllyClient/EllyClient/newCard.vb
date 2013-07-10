Public Class newCard
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents lbName As System.Windows.Forms.Label
    Friend WithEvents lbAmt As System.Windows.Forms.Label
    Friend WithEvents lbContent As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.lbContent = New System.Windows.Forms.Label
        Me.lbAmt = New System.Windows.Forms.Label
        Me.lbName = New System.Windows.Forms.Label
        Me.btnOk = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lbContent)
        Me.GroupBox1.Controls.Add(Me.lbAmt)
        Me.GroupBox1.Controls.Add(Me.lbName)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(200, 144)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Card Details"
        '
        'lbContent
        '
        Me.lbContent.Location = New System.Drawing.Point(8, 64)
        Me.lbContent.Name = "lbContent"
        Me.lbContent.Size = New System.Drawing.Size(184, 72)
        Me.lbContent.TabIndex = 2
        Me.lbContent.Text = "Content:"
        '
        'lbAmt
        '
        Me.lbAmt.Location = New System.Drawing.Point(8, 40)
        Me.lbAmt.Name = "lbAmt"
        Me.lbAmt.Size = New System.Drawing.Size(184, 16)
        Me.lbAmt.TabIndex = 1
        Me.lbAmt.Text = "Amount:"
        '
        'lbName
        '
        Me.lbName.Location = New System.Drawing.Point(8, 16)
        Me.lbName.Name = "lbName"
        Me.lbName.Size = New System.Drawing.Size(184, 16)
        Me.lbName.TabIndex = 0
        Me.lbName.Text = "Name:"
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(72, 160)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.TabIndex = 1
        Me.btnOk.Text = "Accept"
        '
        'newCard
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(216, 190)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "newCard"
        Me.Text = "newCard"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public cdName As String
    Public cdAmt As String
    Public cdContent As String
    Public cdType As String
    Public cn As New ADODB.Connection
    Public ellyClient As New EllyClient
    Private sql As String
    Private rs As New ADODB.Recordset
    Private Sub newCard_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lbName.Text = "Name:" & cdName
        lbAmt.Text = "Amount:" & cdAmt
        lbContent.Text = "Content:" & cdContent
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        If cdType = "N" Then
            sql = "update EllyCards set CardType = 'E' where CardName = '" & cdName & "';"
        ElseIf cdType = "NO" Then
            sql = "update EllyCards set CardType = 'O' where CardName = '" & cdName & "';"
        ElseIf cdType = "S" Then
            sql = "update EllyCards set CardType = 'E' where CardName = '" & cdName & "';"
        ElseIf cdType = "SO" Then
            sql = "update EllyCards set CardType = 'O' where CardName = '" & cdName & "';"
        End If
        rs.Open(sql, cn)
        ellyClient.rfshForm()
        Me.Close()
    End Sub
End Class

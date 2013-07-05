Public Class edDetail
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
    Friend WithEvents btnCdRst As System.Windows.Forms.Button
    Friend WithEvents edCdContent As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents edCdType As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents edCdAmt As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents edCdName As System.Windows.Forms.TextBox
    Friend WithEvents btnCdCfm As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnCdRst = New System.Windows.Forms.Button
        Me.edCdContent = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.edCdType = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.edCdAmt = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.edCdName = New System.Windows.Forms.TextBox
        Me.btnCdCfm = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnCdRst)
        Me.GroupBox1.Controls.Add(Me.edCdContent)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.edCdType)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.edCdAmt)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.edCdName)
        Me.GroupBox1.Controls.Add(Me.btnCdCfm)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(168, 208)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Card Maker"
        '
        'btnCdRst
        '
        Me.btnCdRst.Location = New System.Drawing.Point(88, 176)
        Me.btnCdRst.Name = "btnCdRst"
        Me.btnCdRst.TabIndex = 8
        Me.btnCdRst.Text = "Reset"
        '
        'edCdContent
        '
        Me.edCdContent.Location = New System.Drawing.Point(8, 112)
        Me.edCdContent.Multiline = True
        Me.edCdContent.Name = "edCdContent"
        Me.edCdContent.Size = New System.Drawing.Size(152, 56)
        Me.edCdContent.TabIndex = 7
        Me.edCdContent.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 96)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 24)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Content:"
        '
        'edCdType
        '
        Me.edCdType.Location = New System.Drawing.Point(56, 72)
        Me.edCdType.Name = "edCdType"
        Me.edCdType.Size = New System.Drawing.Size(48, 20)
        Me.edCdType.TabIndex = 5
        Me.edCdType.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 24)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Type:"
        '
        'edCdAmt
        '
        Me.edCdAmt.Location = New System.Drawing.Point(56, 48)
        Me.edCdAmt.Name = "edCdAmt"
        Me.edCdAmt.Size = New System.Drawing.Size(48, 20)
        Me.edCdAmt.TabIndex = 3
        Me.edCdAmt.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 24)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Amount:"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 24)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Name:"
        '
        'edCdName
        '
        Me.edCdName.Location = New System.Drawing.Point(56, 24)
        Me.edCdName.Name = "edCdName"
        Me.edCdName.Size = New System.Drawing.Size(104, 20)
        Me.edCdName.TabIndex = 1
        Me.edCdName.Text = ""
        '
        'btnCdCfm
        '
        Me.btnCdCfm.Location = New System.Drawing.Point(8, 176)
        Me.btnCdCfm.Name = "btnCdCfm"
        Me.btnCdCfm.TabIndex = 1
        Me.btnCdCfm.Text = "Confirm"
        '
        'edDetail
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(184, 222)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "edDetail"
        Me.Text = "edDetail"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Public cardName As String
    Public cardAmt As Integer
    Public cardType As Char
    Public sql As String
    Public cn As New ADODB.Connection
    Public rs As New ADODB.Recordset
    Private Sub edDetail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.initInfo()
    End Sub

    Private Sub btnCdRst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCdRst.Click
        Me.initInfo()
    End Sub

    Private Sub initInfo()
        edCdName.Text = cardName
        edCdAmt.Text = cardAmt
        edCdType.Text = cardType
        sql = "select CardDesc from EllyCards where CardName = '" & cardName & "';"
        rs.Open(sql, cn)
        If rs.RecordCount Then
            edCdContent.Text = rs(0).Value
        End If
    End Sub

    Private Sub btnCdCfm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCdCfm.Click
        rs.Close()
        sql = "update EllyCards set CardName = '" & edCdName.Text & "',CardAmt = " & edCdAmt.Text & ",CardType = '" & edCdType.Text & "',CardDesc = '" & edCdContent.Text & "' Where CardName = '" & cardName & "';"
        rs.Open(sql, cn)
        Me.Close()
    End Sub
End Class

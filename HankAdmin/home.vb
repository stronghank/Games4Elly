
Public Class home
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents edCdName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents edCdAmt As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents edCdType As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnCdGen As System.Windows.Forms.Button
    Friend WithEvents btnCdRst As System.Windows.Forms.Button
    Friend WithEvents edCdContent As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnChgAmt As System.Windows.Forms.Button
    Friend WithEvents btnEdDetail As System.Windows.Forms.Button
    Friend WithEvents btDel As System.Windows.Forms.Button
    Friend WithEvents edEdAmt As System.Windows.Forms.TextBox
    Friend WithEvents lbCdList As System.Windows.Forms.ListBox
    Friend WithEvents btnRfsh As System.Windows.Forms.Button
    Friend WithEvents btnQuit As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.edCdContent = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.edCdType = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.edCdAmt = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.edCdName = New System.Windows.Forms.TextBox
        Me.btnCdGen = New System.Windows.Forms.Button
        Me.btnCdRst = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnChgAmt = New System.Windows.Forms.Button
        Me.btnEdDetail = New System.Windows.Forms.Button
        Me.btDel = New System.Windows.Forms.Button
        Me.edEdAmt = New System.Windows.Forms.TextBox
        Me.lbCdList = New System.Windows.Forms.ListBox
        Me.btnRfsh = New System.Windows.Forms.Button
        Me.btnQuit = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
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
        Me.GroupBox1.Controls.Add(Me.btnCdGen)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(168, 208)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Card Maker"
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
        'btnCdGen
        '
        Me.btnCdGen.Location = New System.Drawing.Point(8, 176)
        Me.btnCdGen.Name = "btnCdGen"
        Me.btnCdGen.TabIndex = 1
        Me.btnCdGen.Text = "Generate"
        '
        'btnCdRst
        '
        Me.btnCdRst.Location = New System.Drawing.Point(88, 176)
        Me.btnCdRst.Name = "btnCdRst"
        Me.btnCdRst.TabIndex = 8
        Me.btnCdRst.Text = "Reset"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lbCdList)
        Me.GroupBox2.Controls.Add(Me.btDel)
        Me.GroupBox2.Controls.Add(Me.btnEdDetail)
        Me.GroupBox2.Controls.Add(Me.btnChgAmt)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.edEdAmt)
        Me.GroupBox2.Location = New System.Drawing.Point(184, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(168, 208)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Card Editer"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 24)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "Card List:"
        '
        'btnChgAmt
        '
        Me.btnChgAmt.Location = New System.Drawing.Point(8, 144)
        Me.btnChgAmt.Name = "btnChgAmt"
        Me.btnChgAmt.TabIndex = 4
        Me.btnChgAmt.Text = "Edit amount"
        '
        'btnEdDetail
        '
        Me.btnEdDetail.Location = New System.Drawing.Point(8, 176)
        Me.btnEdDetail.Name = "btnEdDetail"
        Me.btnEdDetail.TabIndex = 5
        Me.btnEdDetail.Text = "Edit detail"
        '
        'btDel
        '
        Me.btDel.Location = New System.Drawing.Point(88, 176)
        Me.btDel.Name = "btDel"
        Me.btDel.TabIndex = 6
        Me.btDel.Text = "Delete"
        '
        'edEdAmt
        '
        Me.edEdAmt.Location = New System.Drawing.Point(88, 144)
        Me.edEdAmt.Name = "edEdAmt"
        Me.edEdAmt.Size = New System.Drawing.Size(72, 20)
        Me.edEdAmt.TabIndex = 9
        Me.edEdAmt.Text = ""
        '
        'lbCdList
        '
        Me.lbCdList.Location = New System.Drawing.Point(8, 40)
        Me.lbCdList.Name = "lbCdList"
        Me.lbCdList.Size = New System.Drawing.Size(152, 95)
        Me.lbCdList.TabIndex = 10
        '
        'btnRfsh
        '
        Me.btnRfsh.Location = New System.Drawing.Point(16, 224)
        Me.btnRfsh.Name = "btnRfsh"
        Me.btnRfsh.TabIndex = 2
        Me.btnRfsh.Text = "Refresh"
        '
        'btnQuit
        '
        Me.btnQuit.Location = New System.Drawing.Point(96, 224)
        Me.btnQuit.Name = "btnQuit"
        Me.btnQuit.TabIndex = 3
        Me.btnQuit.Text = "Quit"
        '
        'home
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(360, 254)
        Me.Controls.Add(Me.btnQuit)
        Me.Controls.Add(Me.btnRfsh)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "home"
        Me.Text = "Admin"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private cn As New ADODB.Connection
    Private rs As New ADODB.Recordset
    Private sql As String
    Private Sub home_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim conStr As String
        Dim userName As String
        Dim passWord As String
        Dim host As String
        Dim dbName As String
        Dim temp As String
        lbCdList.Items.Clear()
        userName = "sql313005"
        passWord = "iI8!qL4%"
        host = "sql3.freesqldatabase.com"
        dbName = "sql313005"
        conStr = "DRIVER={MySQL ODBC 5.2 Unicode Driver}; SERVER=" & host & "; DATABASE=" & dbName & "; UID=" & userName & ";passWord=" & passWord & "; OPTION=3 stmt=SET NAMES GB2312"
        cn.Open(conStr)
        Me.rfshCdList()
    End Sub

    Private Sub btnCdGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCdGen.Click
        On Error GoTo cdGenErrHandle
        Dim cdName As String
        Dim cdAmt As Integer
        Dim cdType As Char
        Dim cdContent As String
        Dim nth As String
        nth = ""
        cdName = edCdName.Text
        cdAmt = edCdAmt.Text
        cdType = edCdType.Text
        cdContent = edCdContent.Text
        sql = "INSERT INTO EllyCards Values('" & cdName & "','" & cdAmt & "','" & cdContent & "','" & nth & "','" & cdType & "');"
        rs.Open(sql, cn)
        Me.rfshCdList()
        Exit Sub
cdGenErrHandle:
        MsgBox("Error occurs: Maybe your network is bad, or the card name you entered is already existed.")

    End Sub
    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub btnRfsh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRfsh.Click
        Me.rfshCdList()
    End Sub

    Private Sub btnEdDetail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdDetail.Click
        Dim subForm As New edDetail
        Dim temp As String
        Dim tempArr As Array
        If lbCdList.SelectedItem <> "" Then
            temp = lbCdList.SelectedItem
            tempArr = Split(temp, " ")
            subForm.cardName = tempArr(0)
            subForm.cardAmt = tempArr(1)
            subForm.cardType = tempArr(2)
            subForm.cn = cn
            subForm.ShowDialog()
            Me.rfshCdList()
        End If
    End Sub

    Private Sub btnChgAmt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChgAmt.Click
        On Error GoTo errHandle
        Dim newAmt As Integer
        Dim cdName As String
        Dim nameList As Array
        newAmt = edEdAmt.Text
        cdName = lbCdList.SelectedItem
        nameList = Split(cdName, " ")
        sql = "update EllyCards set CardAmt =" & newAmt & " where CardName = '" & nameList(0) & "';"
        rs.Open(sql, cn)
        Me.rfshCdList()
        Exit Sub
errHandle:
        MsgBox("Error occurs: Maybe your network is bad or you have entered an invalid number.")
    End Sub

    Private Sub btnQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuit.Click
        Me.Close()
    End Sub

    Private Sub rfshCdList()
        sql = "select * from EllyCards"
        Dim temp As String
        lbCdList.Items.Clear()
        rs.Open(sql, cn)
        Do While Not rs.EOF
            temp = rs(0).Value & " " & rs(1).Value & " " & rs("CardType").Value
            lbCdList.Items.Add(temp)
            rs.MoveNext()
        Loop
        rs.Close()
    End Sub

    Private Sub btDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btDel.Click
        Dim cdName As String = lbCdList.SelectedItem
        Dim nameList As Array = Split(cdName, " ")
        sql = "delete from EllyCards where CardName = '" & nameList(0) & "';"
        rs.Open(sql, cn)
        Me.rfshCdList()
    End Sub

    Private Sub btnCdRst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCdRst.Click
        edCdName.Clear()
        edCdAmt.Clear()
        edCdContent.Clear()
        edCdType.Clear()
    End Sub
End Class

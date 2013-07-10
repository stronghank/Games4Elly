Public Class EllyClient
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
    Friend WithEvents lbCdList As System.Windows.Forms.ListBox
    Friend WithEvents pTtlAmt As System.Windows.Forms.ProgressBar
    Friend WithEvents lbTtlAmt As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnGetTdCd As System.Windows.Forms.Button
    Friend WithEvents btnUseCd As System.Windows.Forms.Button
    Friend WithEvents btnRfh As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents edForWrite As System.Windows.Forms.TextBox
    Friend WithEvents lbCdName As System.Windows.Forms.Label
    Friend WithEvents lbCdAmt As System.Windows.Forms.Label
    Friend WithEvents lbCdContent As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnReset As System.Windows.Forms.Button
    Friend WithEvents btnSuperise As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnRfh = New System.Windows.Forms.Button
        Me.btnUseCd = New System.Windows.Forms.Button
        Me.lbTtlAmt = New System.Windows.Forms.Label
        Me.pTtlAmt = New System.Windows.Forms.ProgressBar
        Me.lbCdList = New System.Windows.Forms.ListBox
        Me.btnGetTdCd = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnSuperise = New System.Windows.Forms.Button
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.edForWrite = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.lbCdContent = New System.Windows.Forms.Label
        Me.lbCdAmt = New System.Windows.Forms.Label
        Me.lbCdName = New System.Windows.Forms.Label
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnReset = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnRfh)
        Me.GroupBox1.Controls.Add(Me.btnUseCd)
        Me.GroupBox1.Controls.Add(Me.lbTtlAmt)
        Me.GroupBox1.Controls.Add(Me.pTtlAmt)
        Me.GroupBox1.Controls.Add(Me.lbCdList)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(200, 176)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Cards List"
        '
        'btnRfh
        '
        Me.btnRfh.Location = New System.Drawing.Point(104, 144)
        Me.btnRfh.Name = "btnRfh"
        Me.btnRfh.Size = New System.Drawing.Size(88, 23)
        Me.btnRfh.TabIndex = 5
        Me.btnRfh.Text = "Refresh List"
        '
        'btnUseCd
        '
        Me.btnUseCd.Enabled = False
        Me.btnUseCd.Location = New System.Drawing.Point(8, 144)
        Me.btnUseCd.Name = "btnUseCd"
        Me.btnUseCd.Size = New System.Drawing.Size(88, 23)
        Me.btnUseCd.TabIndex = 4
        Me.btnUseCd.Text = "Use this Card"
        '
        'lbTtlAmt
        '
        Me.lbTtlAmt.Location = New System.Drawing.Point(144, 16)
        Me.lbTtlAmt.Name = "lbTtlAmt"
        Me.lbTtlAmt.Size = New System.Drawing.Size(48, 23)
        Me.lbTtlAmt.TabIndex = 3
        '
        'pTtlAmt
        '
        Me.pTtlAmt.Location = New System.Drawing.Point(8, 16)
        Me.pTtlAmt.Name = "pTtlAmt"
        Me.pTtlAmt.Size = New System.Drawing.Size(120, 16)
        Me.pTtlAmt.TabIndex = 2
        '
        'lbCdList
        '
        Me.lbCdList.Location = New System.Drawing.Point(8, 40)
        Me.lbCdList.Name = "lbCdList"
        Me.lbCdList.Size = New System.Drawing.Size(184, 95)
        Me.lbCdList.TabIndex = 1
        '
        'btnGetTdCd
        '
        Me.btnGetTdCd.Location = New System.Drawing.Point(8, 16)
        Me.btnGetTdCd.Name = "btnGetTdCd"
        Me.btnGetTdCd.Size = New System.Drawing.Size(184, 23)
        Me.btnGetTdCd.TabIndex = 1
        Me.btnGetTdCd.Text = "Get Today's new cards"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnSuperise)
        Me.GroupBox2.Controls.Add(Me.btnGetTdCd)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 192)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(200, 80)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Getting Cards"
        '
        'btnSuperise
        '
        Me.btnSuperise.Location = New System.Drawing.Point(8, 48)
        Me.btnSuperise.Name = "btnSuperise"
        Me.btnSuperise.Size = New System.Drawing.Size(184, 23)
        Me.btnSuperise.TabIndex = 2
        Me.btnSuperise.Text = "Here maybe a suprise"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.edForWrite)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Controls.Add(Me.lbCdContent)
        Me.GroupBox3.Controls.Add(Me.lbCdAmt)
        Me.GroupBox3.Controls.Add(Me.lbCdName)
        Me.GroupBox3.Controls.Add(Me.btnSave)
        Me.GroupBox3.Controls.Add(Me.btnReset)
        Me.GroupBox3.Location = New System.Drawing.Point(216, 8)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(200, 264)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Card detail"
        '
        'edForWrite
        '
        Me.edForWrite.Location = New System.Drawing.Point(8, 136)
        Me.edForWrite.Multiline = True
        Me.edForWrite.Name = "edForWrite"
        Me.edForWrite.Size = New System.Drawing.Size(184, 88)
        Me.edForWrite.TabIndex = 4
        Me.edForWrite.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 120)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(176, 16)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "For you to write: (English only)"
        '
        'lbCdContent
        '
        Me.lbCdContent.Location = New System.Drawing.Point(8, 72)
        Me.lbCdContent.Name = "lbCdContent"
        Me.lbCdContent.Size = New System.Drawing.Size(184, 40)
        Me.lbCdContent.TabIndex = 2
        Me.lbCdContent.Text = "Content:"
        '
        'lbCdAmt
        '
        Me.lbCdAmt.Location = New System.Drawing.Point(8, 48)
        Me.lbCdAmt.Name = "lbCdAmt"
        Me.lbCdAmt.Size = New System.Drawing.Size(184, 16)
        Me.lbCdAmt.TabIndex = 1
        Me.lbCdAmt.Text = "Amount:"
        '
        'lbCdName
        '
        Me.lbCdName.Location = New System.Drawing.Point(8, 24)
        Me.lbCdName.Name = "lbCdName"
        Me.lbCdName.Size = New System.Drawing.Size(184, 16)
        Me.lbCdName.TabIndex = 0
        Me.lbCdName.Text = "Name:"
        '
        'btnSave
        '
        Me.btnSave.Enabled = False
        Me.btnSave.Location = New System.Drawing.Point(8, 232)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(88, 23)
        Me.btnSave.TabIndex = 6
        Me.btnSave.Text = "Save"
        '
        'btnReset
        '
        Me.btnReset.Location = New System.Drawing.Point(104, 232)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(88, 23)
        Me.btnReset.TabIndex = 6
        Me.btnReset.Text = "Reset"
        '
        'EllyClient
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(426, 278)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "EllyClient"
        Me.Text = "Elly's cards Box"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private cn As New ADODB.Connection
    Private rs As New ADODB.Recordset
    Private rs2 As New ADODB.Recordset
    Private sql As String
    Private Sub EllyClient_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim conStr As String
        Dim userName As String
        Dim passWord As String
        Dim host As String
        Dim dbName As String
        userName = "sql313005"
        passWord = "iI8!qL4%"
        host = "sql3.freesqldatabase.com"
        dbName = "sql313005"
        conStr = "DRIVER={MySQL ODBC 5.2 Unicode Driver}; SERVER=" & host & "; DATABASE=" & dbName & "; UID=" & userName & ";passWord=" & passWord & "; OPTION=3 stmt=SET NAMES GB2312"
        cn.Open(conStr)
        Me.rfshForm()
    End Sub

    Public Sub rfshForm()
        btnUseCd.Enabled = False
        lbCdList.Items.Clear()
        lbCdName.Text = "Name:"
        lbCdAmt.Text = "Amount:"
        lbCdContent.Text = "Content:"
        edForWrite.Text = ""
        Dim item As String
        Dim ttlAmt As String
        Dim currAmt As Double = 0
        sql = "select * from EllyCards where CardType = 'E' or CardType = 'O';"
        rs2.Open(sql, cn)
        Do While Not rs2.EOF
            item = rs2(0).Value & " * " & rs2(1).Value
            lbCdList.Items.Add(item)
            currAmt = currAmt + rs2(1).Value
            rs2.MoveNext()
        Loop
        rs2.Close()
        sql = "select CarAmtLmt from CardsCtrl;"
        rs2.Open(sql, cn)
        ttlAmt = currAmt & "/" & rs2("CarAmtLmt").Value
        lbTtlAmt.Text = ttlAmt
        currAmt = currAmt / rs2("CarAmtLmt").Value
        pTtlAmt.Value = currAmt * 100
        rs2.Close()
    End Sub

    Private Sub lbCdList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbCdList.SelectedIndexChanged

        Dim arrItem As Array
        If lbCdList.SelectedIndex >= 0 Then
            btnUseCd.Enabled = True
            arrItem = Split(lbCdList.SelectedItem, " ")
            sql = "select * from EllyCards where CardName = '" & arrItem(0) & "' and ( CardType = 'E' or CardType = 'O');"
            rs.Open(sql, cn)
            lbCdName.Text = "Name:" & rs(0).Value
            lbCdAmt.Text = "Amount:" & rs(1).Value
            lbCdContent.Text = "Content:" & rs(2).Value
            edForWrite.Text = rs(3).Value
            rs.Close()
        Else
            btnUseCd.Enabled = False
        End If
        Exit Sub
errHandler:
        MsgBox("Bad Network like 1 tuo shit")
    End Sub

    Private Sub btnUseCd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUseCd.Click
        Dim arrItem As Array
        arrItem = Split(lbCdList.SelectedItem, " ")
        If CInt(arrItem(2)) = 0 Then
            MsgBox("Give me a fuck if you really want to use this none existing card.")
            Exit Sub
        End If
        sql = "select CardType, CardName from EllyCards where CardName = '" & arrItem(0) & "';"
        rs.Open(sql, cn)
        If rs(0).Value = "O" Then
            Select Case rs(1).Value
                Case "Add"
                    rs.Close()
                    sql = "update CardsCtrl set CarAmtLmt = CarAmtLmt +1;"
                    rs.Open(sql, cn)
                    sql = "update EllyCards set CardAmt =" & (arrItem(2) - 1) & " where CardName ='" & arrItem(0) & "';"
                    rs.Open(sql, cn)
                    Me.rfshForm()
                    MsgBox("Your Maximum card amount is added")
            End Select
        Else
            rs.Close()
            sql = "update EllyCards set CardAmt =" & (arrItem(2) - 1) & " where CardName ='" & arrItem(0) & "';"
            rs.Open(sql, cn)
            Me.rfshForm()
            MsgBox("Hank will be notified, or you can tell him immediately.")
        End If
        Exit Sub
errHandler:
        MsgBox("Network issue, retry later")
    End Sub

    Private Sub btnRfh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRfh.Click
        rfshForm()
    End Sub

    Private Sub edForWrite_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edForWrite.TextChanged
        If edForWrite.Text <> Nothing Then
            btnSave.Enabled = True
        Else
            btnSave.Enabled = False
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        On Error GoTo errHandler
        Dim arrItem As Array
        arrItem = Split(lbCdList.SelectedItem, " ")
        sql = "update EllyCards set CardMoreDesc ='" & edForWrite.Text & "' where CardName ='" & arrItem(0) & "';"
        rs.Open(sql, cn)
        Exit Sub
errHandler:
        MsgBox("Invalid operation")
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        On Error GoTo errHandler
        Dim arrItem As Array
        arrItem = Split(lbCdList.SelectedItem, " ")
        sql = "select CardMoreDesc from EllyCards where CardName = '" & arrItem(0) & "';"
        rs.Open(sql, cn)
        edForWrite.Text = rs(0).Value
        rs.Close()
        Exit Sub
errHandler:
        MsgBox("Invalid operation")
    End Sub

    Private Sub btnGetTdCd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetTdCd.Click
        sql = "select * from EllyCards where CardType = 'N' or CardType = 'NO';"
        Dim cardAdded As Boolean = False
        rs.Open(sql, cn)
        sql = "select nextdate, now() from CardsCtrl"
        rs2.Open(sql, cn)
        If rs2(0).Value > rs2(1).Value Then
            MsgBox("If you perform well tonight, I will give you more cards tommorrow.")
            rs2.Close()
            rs.Close()
            Exit Sub
        Else
            sql = "update CardsCtrl set nextDate = date_add(nextdate,interval 1 day);"
            rs2.Close()
            rs2.Open(sql, cn)
        End If

        Do While Not rs.EOF
            If rs(0).Value = "0" Then
                MsgBox("Your card: " & rs(2).Value & " is added by " & rs(1).Value)
                sql = "update EllyCards set CardAmt = CardAmt + " & rs(1).Value & " where CardName = '" & rs(2).Value & "';"
                rs2.Open(sql, cn)
                rs.MoveNext()
                cardAdded = True
                Me.rfshForm()
            Else
                Dim card As New newCard
                card.cdName = rs(0).Value
                card.cdAmt = rs(1).Value
                card.cdContent = rs(2).Value
                card.cdType = rs(4).Value
                card.cn = cn
                card.ellyClient = Me
                card.Show()
                rs.MoveNext()
            End If
        Loop
        rs.Close()
        If cardAdded Then
            sql = "delete from EllyCards where CardName = '0' ;"
            rs.Open(sql, cn)
        End If
    End Sub


    Private Sub btnSuperise_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSuperise.Click
        sql = "select * from EllyCards where CardType = 'S' or CardType = 'SO';"
        Dim cardAdded As Boolean = False
        rs.Open(sql, cn)
        If rs.EOF Then
            MsgBox("Superise comes when Hank is in a good feeling.")
        End If
        Do While Not rs.EOF
            If rs(0).Value = "0" Then
                MsgBox("Your card: " & rs(2).Value & " is added by " & rs(1).Value)
                sql = "update EllyCards set CardAmt = CardAmt + " & rs(1).Value & " where CardName = '" & rs(2).Value & "';"
                rs2.Open(sql, cn)
                rs.MoveNext()
                cardAdded = True
                Me.rfshForm()
            Else
                Dim card As New newCard
                card.cdName = rs(0).Value
                card.cdAmt = rs(1).Value
                card.cdContent = rs(2).Value
                card.cdType = rs(4).Value
                card.cn = cn
                card.ellyClient = Me
                card.Show()
                rs.MoveNext()
            End If
        Loop
        rs.Close()
        If cardAdded Then
            sql = "delete from EllyCards where CardName = '0' ;"
            rs.Open(sql, cn)
        End If
    End Sub
End Class

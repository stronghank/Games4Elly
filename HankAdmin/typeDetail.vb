Public Class typeDetail
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
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents lbTpDetail As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnOk = New System.Windows.Forms.Button
        Me.lbTpDetail = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(104, 136)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.TabIndex = 1
        Me.btnOk.Text = "OK"
        '
        'lbTpDetail
        '
        Me.lbTpDetail.Location = New System.Drawing.Point(8, 8)
        Me.lbTpDetail.Name = "lbTpDetail"
        Me.lbTpDetail.Size = New System.Drawing.Size(280, 120)
        Me.lbTpDetail.TabIndex = 2
        '
        'typeDetail
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 166)
        Me.Controls.Add(Me.lbTpDetail)
        Me.Controls.Add(Me.btnOk)
        Me.Name = "typeDetail"
        Me.Text = "typeDetail"
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub typeDetail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim detail As String = "E: existing card, can be used" & Chr(13) & Chr(10) & _
                               "N: new cards for today, can be checked" & Chr(13) & Chr(10) & _
                               "S: cards for superise" & Chr(13) & Chr(10) & _
                               "ES: existing special card" & Chr(13) & Chr(10) & _
                               "NS: new special cards for today, can be checked" & Chr(13) & Chr(10) & _
                               "SS: special cards for superise" & Chr(13) & Chr(10) & _
                               "How to add the amount of existing card: set the card name to 0 and the card content to the existing card name, the card amount is the amount you want to add"
        lbTpDetail.Text = detail
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Me.Close()
    End Sub
End Class

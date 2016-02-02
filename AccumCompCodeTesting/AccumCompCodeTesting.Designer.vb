<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AccumCompCodeTesting
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbEnv = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblSpreadsheetLoc = New System.Windows.Forms.Label()
        Me.lblSpreadsheetName = New System.Windows.Forms.Label()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.lblRxCounter = New System.Windows.Forms.Label()
        Me.lblFillDate = New System.Windows.Forms.Label()
        Me.DTP_FillDate = New System.Windows.Forms.DateTimePicker()
        Me.txtRxNum = New System.Windows.Forms.TextBox()
        Me.lblProductID = New System.Windows.Forms.Label()
        Me.txtProdID = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblQty = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblDaySupply = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblCost = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Environment"
        '
        'cmbEnv
        '
        Me.cmbEnv.FormattingEnabled = True
        Me.cmbEnv.Items.AddRange(New Object() {"DEV01", "DEV02", "PROD01", "PROD03"})
        Me.cmbEnv.Location = New System.Drawing.Point(106, 24)
        Me.cmbEnv.Name = "cmbEnv"
        Me.cmbEnv.Size = New System.Drawing.Size(121, 21)
        Me.cmbEnv.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(14, 241)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(114, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Spreadsheet Location:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(15, 268)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(101, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Spreadsheet Name:"
        '
        'lblSpreadsheetLoc
        '
        Me.lblSpreadsheetLoc.AutoSize = True
        Me.lblSpreadsheetLoc.Location = New System.Drawing.Point(135, 241)
        Me.lblSpreadsheetLoc.Name = "lblSpreadsheetLoc"
        Me.lblSpreadsheetLoc.Size = New System.Drawing.Size(22, 13)
        Me.lblSpreadsheetLoc.TabIndex = 4
        Me.lblSpreadsheetLoc.Text = "NA"
        '
        'lblSpreadsheetName
        '
        Me.lblSpreadsheetName.AutoSize = True
        Me.lblSpreadsheetName.Location = New System.Drawing.Point(138, 267)
        Me.lblSpreadsheetName.Name = "lblSpreadsheetName"
        Me.lblSpreadsheetName.Size = New System.Drawing.Size(22, 13)
        Me.lblSpreadsheetName.TabIndex = 5
        Me.lblSpreadsheetName.Text = "NA"
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(106, 183)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(75, 23)
        Me.btnRun.TabIndex = 6
        Me.btnRun.Text = "Run"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'lblRxCounter
        '
        Me.lblRxCounter.AutoSize = True
        Me.lblRxCounter.Location = New System.Drawing.Point(12, 64)
        Me.lblRxCounter.Name = "lblRxCounter"
        Me.lblRxCounter.Size = New System.Drawing.Size(57, 13)
        Me.lblRxCounter.TabIndex = 7
        Me.lblRxCounter.Text = "RxNumber"
        '
        'lblFillDate
        '
        Me.lblFillDate.AutoSize = True
        Me.lblFillDate.Location = New System.Drawing.Point(13, 96)
        Me.lblFillDate.Name = "lblFillDate"
        Me.lblFillDate.Size = New System.Drawing.Size(45, 13)
        Me.lblFillDate.TabIndex = 8
        Me.lblFillDate.Text = "Fill Date"
        '
        'DTP_FillDate
        '
        Me.DTP_FillDate.Location = New System.Drawing.Point(106, 96)
        Me.DTP_FillDate.Name = "DTP_FillDate"
        Me.DTP_FillDate.Size = New System.Drawing.Size(200, 20)
        Me.DTP_FillDate.TabIndex = 9
        Me.DTP_FillDate.Value = New Date(2015, 1, 1, 0, 0, 0, 0)
        '
        'txtRxNum
        '
        Me.txtRxNum.Location = New System.Drawing.Point(106, 64)
        Me.txtRxNum.Name = "txtRxNum"
        Me.txtRxNum.Size = New System.Drawing.Size(121, 20)
        Me.txtRxNum.TabIndex = 10
        Me.txtRxNum.Text = "100"
        '
        'lblProductID
        '
        Me.lblProductID.AutoSize = True
        Me.lblProductID.Location = New System.Drawing.Point(13, 139)
        Me.lblProductID.Name = "lblProductID"
        Me.lblProductID.Size = New System.Drawing.Size(87, 13)
        Me.lblProductID.TabIndex = 11
        Me.lblProductID.Text = "ProductID (NDC)"
        '
        'txtProdID
        '
        Me.txtProdID.Location = New System.Drawing.Point(106, 136)
        Me.txtProdID.Name = "txtProdID"
        Me.txtProdID.Size = New System.Drawing.Size(121, 20)
        Me.txtProdID.TabIndex = 12
        Me.txtProdID.Text = "61958070101"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(16, 26)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(26, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Qty:"
        '
        'lblQty
        '
        Me.lblQty.AutoSize = True
        Me.lblQty.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQty.Location = New System.Drawing.Point(88, 26)
        Me.lblQty.Name = "lblQty"
        Me.lblQty.Size = New System.Drawing.Size(24, 13)
        Me.lblQty.TabIndex = 14
        Me.lblQty.Text = "NA"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(16, 52)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 13)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Day Supply:"
        '
        'lblDaySupply
        '
        Me.lblDaySupply.AutoSize = True
        Me.lblDaySupply.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDaySupply.Location = New System.Drawing.Point(88, 52)
        Me.lblDaySupply.Name = "lblDaySupply"
        Me.lblDaySupply.Size = New System.Drawing.Size(24, 13)
        Me.lblDaySupply.TabIndex = 16
        Me.lblDaySupply.Text = "NA"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(16, 81)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(31, 13)
        Me.Label8.TabIndex = 17
        Me.Label8.Text = "Cost:"
        '
        'lblCost
        '
        Me.lblCost.AutoSize = True
        Me.lblCost.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCost.Location = New System.Drawing.Point(88, 81)
        Me.lblCost.Name = "lblCost"
        Me.lblCost.Size = New System.Drawing.Size(24, 13)
        Me.lblCost.TabIndex = 18
        Me.lblCost.Text = "NA"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblCost)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.lblQty)
        Me.GroupBox1.Controls.Add(Me.lblDaySupply)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Location = New System.Drawing.Point(323, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(264, 104)
        Me.GroupBox1.TabIndex = 19
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Hard-Coded Values"
        '
        'AccumCompCodeTesting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(599, 302)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtProdID)
        Me.Controls.Add(Me.lblProductID)
        Me.Controls.Add(Me.txtRxNum)
        Me.Controls.Add(Me.DTP_FillDate)
        Me.Controls.Add(Me.lblFillDate)
        Me.Controls.Add(Me.lblRxCounter)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.lblSpreadsheetName)
        Me.Controls.Add(Me.lblSpreadsheetLoc)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmbEnv)
        Me.Controls.Add(Me.Label1)
        Me.Name = "AccumCompCodeTesting"
        Me.Text = "Accumulator Component Code"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbEnv As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblSpreadsheetLoc As System.Windows.Forms.Label
    Friend WithEvents lblSpreadsheetName As System.Windows.Forms.Label
    Friend WithEvents btnRun As System.Windows.Forms.Button
    Friend WithEvents lblRxCounter As System.Windows.Forms.Label
    Friend WithEvents lblFillDate As System.Windows.Forms.Label
    Friend WithEvents DTP_FillDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtRxNum As System.Windows.Forms.TextBox
    Friend WithEvents lblProductID As System.Windows.Forms.Label
    Friend WithEvents txtProdID As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblQty As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblDaySupply As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents lblCost As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox

End Class

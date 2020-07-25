<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Find
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Find))
        Me.txtFind = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtReplace = New System.Windows.Forms.TextBox()
        Me.chkMC = New System.Windows.Forms.CheckBox()
        Me.btnFind = New System.Windows.Forms.Button()
        Me.btnFindN = New System.Windows.Forms.Button()
        Me.btnReplace = New System.Windows.Forms.Button()
        Me.btnReplaceA = New System.Windows.Forms.Button()
        Me.chkWW = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'txtFind
        '
        Me.txtFind.Location = New System.Drawing.Point(29, 25)
        Me.txtFind.Name = "txtFind"
        Me.txtFind.Size = New System.Drawing.Size(220, 20)
        Me.txtFind.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(116, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Find :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(106, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Replace :"
        '
        'txtReplace
        '
        Me.txtReplace.Location = New System.Drawing.Point(29, 64)
        Me.txtReplace.Name = "txtReplace"
        Me.txtReplace.Size = New System.Drawing.Size(220, 20)
        Me.txtReplace.TabIndex = 2
        '
        'chkMC
        '
        Me.chkMC.AutoSize = True
        Me.chkMC.Location = New System.Drawing.Point(29, 103)
        Me.chkMC.Name = "chkMC"
        Me.chkMC.Size = New System.Drawing.Size(83, 17)
        Me.chkMC.TabIndex = 4
        Me.chkMC.Text = "Match Case"
        Me.chkMC.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Enabled = False
        Me.btnFind.Location = New System.Drawing.Point(274, 22)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(75, 23)
        Me.btnFind.TabIndex = 5
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnFindN
        '
        Me.btnFindN.Enabled = False
        Me.btnFindN.Location = New System.Drawing.Point(365, 22)
        Me.btnFindN.Name = "btnFindN"
        Me.btnFindN.Size = New System.Drawing.Size(75, 23)
        Me.btnFindN.TabIndex = 6
        Me.btnFindN.Text = "Find &Next"
        Me.btnFindN.UseVisualStyleBackColor = True
        '
        'btnReplace
        '
        Me.btnReplace.Enabled = False
        Me.btnReplace.Location = New System.Drawing.Point(274, 64)
        Me.btnReplace.Name = "btnReplace"
        Me.btnReplace.Size = New System.Drawing.Size(75, 23)
        Me.btnReplace.TabIndex = 7
        Me.btnReplace.Text = "&Replace"
        Me.btnReplace.UseVisualStyleBackColor = True
        '
        'btnReplaceA
        '
        Me.btnReplaceA.Enabled = False
        Me.btnReplaceA.Location = New System.Drawing.Point(365, 64)
        Me.btnReplaceA.Name = "btnReplaceA"
        Me.btnReplaceA.Size = New System.Drawing.Size(75, 23)
        Me.btnReplaceA.TabIndex = 8
        Me.btnReplaceA.Text = "Replace &All"
        Me.btnReplaceA.UseVisualStyleBackColor = True
        '
        'chkWW
        '
        Me.chkWW.AutoSize = True
        Me.chkWW.Location = New System.Drawing.Point(166, 103)
        Me.chkWW.Name = "chkWW"
        Me.chkWW.Size = New System.Drawing.Size(86, 17)
        Me.chkWW.TabIndex = 9
        Me.chkWW.Text = "Whole Word"
        Me.chkWW.UseVisualStyleBackColor = True
        '
        'Find
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(456, 132)
        Me.Controls.Add(Me.chkWW)
        Me.Controls.Add(Me.btnReplaceA)
        Me.Controls.Add(Me.btnReplace)
        Me.Controls.Add(Me.btnFindN)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.chkMC)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtReplace)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtFind)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Find"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Find"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtFind As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtReplace As System.Windows.Forms.TextBox
    Friend WithEvents chkMC As System.Windows.Forms.CheckBox
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnFindN As System.Windows.Forms.Button
    Friend WithEvents btnReplace As System.Windows.Forms.Button
    Friend WithEvents btnReplaceA As System.Windows.Forms.Button
    Friend WithEvents chkWW As System.Windows.Forms.CheckBox
End Class

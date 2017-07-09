<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ExceptionDialog
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
    Friend WithEvents btn1 As System.Windows.Forms.Button
    Friend WithEvents btn2 As System.Windows.Forms.Button
    Friend WithEvents btn3 As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents lblErrorHeading As System.Windows.Forms.Label
    Friend WithEvents lblScopeHeading As System.Windows.Forms.Label
    Friend WithEvents lblActionHeading As System.Windows.Forms.Label
    Friend WithEvents lblMoreHeading As System.Windows.Forms.Label
    Friend WithEvents txtMore As System.Windows.Forms.TextBox
    Friend WithEvents btnMore As System.Windows.Forms.Button
    Friend WithEvents ErrorBox As System.Windows.Forms.RichTextBox
    Friend WithEvents ScopeBox As System.Windows.Forms.RichTextBox
    Friend WithEvents ActionBox As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ExceptionDialog))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.lblErrorHeading = New System.Windows.Forms.Label
        Me.ErrorBox = New System.Windows.Forms.RichTextBox
        Me.lblScopeHeading = New System.Windows.Forms.Label
        Me.ScopeBox = New System.Windows.Forms.RichTextBox
        Me.lblActionHeading = New System.Windows.Forms.Label
        Me.ActionBox = New System.Windows.Forms.RichTextBox
        Me.lblMoreHeading = New System.Windows.Forms.Label
        Me.btn1 = New System.Windows.Forms.Button
        Me.btn2 = New System.Windows.Forms.Button
        Me.btn3 = New System.Windows.Forms.Button
        Me.txtMore = New System.Windows.Forms.TextBox
        Me.btnMore = New System.Windows.Forms.Button
        Me.chkEmail = New System.Windows.Forms.CheckBox
        Me.btn4 = New System.Windows.Forms.Button
        Me.picDropDown = New System.Windows.Forms.PictureBox
        Me.picCloseUp = New System.Windows.Forms.PictureBox
        Me.picUp2 = New System.Windows.Forms.PictureBox
        Me.picDrop2 = New System.Windows.Forms.PictureBox
        Me.CopyToClipboardLink = New System.Windows.Forms.LinkLabel
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picDropDown, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCloseUp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picUp2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picDrop2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(8, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(32, 32)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'lblErrorHeading
        '
        Me.lblErrorHeading.AutoSize = True
        Me.lblErrorHeading.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblErrorHeading.Location = New System.Drawing.Point(48, 4)
        Me.lblErrorHeading.Name = "lblErrorHeading"
        Me.lblErrorHeading.Size = New System.Drawing.Size(123, 16)
        Me.lblErrorHeading.TabIndex = 0
        Me.lblErrorHeading.Text = "What happened..."
        '
        'ErrorBox
        '
        Me.ErrorBox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ErrorBox.BackColor = System.Drawing.SystemColors.Control
        Me.ErrorBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.ErrorBox.CausesValidation = False
        Me.ErrorBox.Location = New System.Drawing.Point(69, 24)
        Me.ErrorBox.Name = "ErrorBox"
        Me.ErrorBox.ReadOnly = True
        Me.ErrorBox.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.ErrorBox.Size = New System.Drawing.Size(395, 64)
        Me.ErrorBox.TabIndex = 1
        Me.ErrorBox.Text = "(error message)"
        '
        'lblScopeHeading
        '
        Me.lblScopeHeading.AutoSize = True
        Me.lblScopeHeading.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblScopeHeading.Location = New System.Drawing.Point(8, 92)
        Me.lblScopeHeading.Name = "lblScopeHeading"
        Me.lblScopeHeading.Size = New System.Drawing.Size(171, 16)
        Me.lblScopeHeading.TabIndex = 2
        Me.lblScopeHeading.Text = "How this will affect you..."
        '
        'ScopeBox
        '
        Me.ScopeBox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ScopeBox.BackColor = System.Drawing.SystemColors.Control
        Me.ScopeBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.ScopeBox.CausesValidation = False
        Me.ScopeBox.Location = New System.Drawing.Point(24, 112)
        Me.ScopeBox.Name = "ScopeBox"
        Me.ScopeBox.ReadOnly = True
        Me.ScopeBox.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.ScopeBox.Size = New System.Drawing.Size(440, 64)
        Me.ScopeBox.TabIndex = 3
        Me.ScopeBox.Text = "(scope)"
        '
        'lblActionHeading
        '
        Me.lblActionHeading.AutoSize = True
        Me.lblActionHeading.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblActionHeading.Location = New System.Drawing.Point(8, 180)
        Me.lblActionHeading.Name = "lblActionHeading"
        Me.lblActionHeading.Size = New System.Drawing.Size(185, 16)
        Me.lblActionHeading.TabIndex = 4
        Me.lblActionHeading.Text = "What you can do about it..."
        '
        'ActionBox
        '
        Me.ActionBox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ActionBox.BackColor = System.Drawing.SystemColors.Control
        Me.ActionBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.ActionBox.CausesValidation = False
        Me.ActionBox.Location = New System.Drawing.Point(24, 200)
        Me.ActionBox.Name = "ActionBox"
        Me.ActionBox.ReadOnly = True
        Me.ActionBox.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.ActionBox.Size = New System.Drawing.Size(440, 92)
        Me.ActionBox.TabIndex = 5
        Me.ActionBox.Text = "(action)"
        '
        'lblMoreHeading
        '
        Me.lblMoreHeading.AutoSize = True
        Me.lblMoreHeading.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMoreHeading.Location = New System.Drawing.Point(26, 301)
        Me.lblMoreHeading.Name = "lblMoreHeading"
        Me.lblMoreHeading.Size = New System.Drawing.Size(119, 16)
        Me.lblMoreHeading.TabIndex = 6
        Me.lblMoreHeading.Text = "More information"
        '
        'btn1
        '
        Me.btn1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btn1.Location = New System.Drawing.Point(142, 544)
        Me.btn1.Name = "btn1"
        Me.btn1.Size = New System.Drawing.Size(75, 23)
        Me.btn1.TabIndex = 9
        Me.btn1.Text = "Button1"
        '
        'btn2
        '
        Me.btn2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btn2.Location = New System.Drawing.Point(223, 544)
        Me.btn2.Name = "btn2"
        Me.btn2.Size = New System.Drawing.Size(75, 23)
        Me.btn2.TabIndex = 10
        Me.btn2.Text = "Button2"
        '
        'btn3
        '
        Me.btn3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn3.Location = New System.Drawing.Point(304, 544)
        Me.btn3.Name = "btn3"
        Me.btn3.Size = New System.Drawing.Size(75, 23)
        Me.btn3.TabIndex = 11
        Me.btn3.Text = "Button3"
        '
        'txtMore
        '
        Me.txtMore.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMore.BackColor = System.Drawing.SystemColors.Control
        Me.txtMore.CausesValidation = False
        Me.txtMore.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMore.Location = New System.Drawing.Point(24, 324)
        Me.txtMore.Multiline = True
        Me.txtMore.Name = "txtMore"
        Me.txtMore.ReadOnly = True
        Me.txtMore.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtMore.Size = New System.Drawing.Size(436, 212)
        Me.txtMore.TabIndex = 8
        Me.txtMore.Text = "(detailed information, such as exception details)"
        Me.txtMore.WordWrap = False
        '
        'btnMore
        '
        Me.btnMore.FlatAppearance.BorderColor = System.Drawing.SystemColors.ControlDark
        Me.btnMore.FlatAppearance.BorderSize = 0
        Me.btnMore.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnMore.Location = New System.Drawing.Point(11, 301)
        Me.btnMore.Margin = New System.Windows.Forms.Padding(0)
        Me.btnMore.Name = "btnMore"
        Me.btnMore.Size = New System.Drawing.Size(16, 17)
        Me.btnMore.TabIndex = 7
        Me.btnMore.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnMore.UseCompatibleTextRendering = True
        '
        'chkEmail
        '
        Me.chkEmail.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkEmail.AutoSize = True
        Me.chkEmail.Location = New System.Drawing.Point(8, 548)
        Me.chkEmail.Name = "chkEmail"
        Me.chkEmail.Size = New System.Drawing.Size(121, 17)
        Me.chkEmail.TabIndex = 12
        Me.chkEmail.Text = "Submit Error Report"
        Me.chkEmail.UseVisualStyleBackColor = True
        Me.chkEmail.Visible = False
        '
        'btn4
        '
        Me.btn4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn4.Location = New System.Drawing.Point(385, 544)
        Me.btn4.Name = "btn4"
        Me.btn4.Size = New System.Drawing.Size(75, 23)
        Me.btn4.TabIndex = 13
        Me.btn4.Text = "Button4"
        '
        'picDropDown
        '
        Me.picDropDown.Image = CType(resources.GetObject("picDropDown.Image"), System.Drawing.Image)
        Me.picDropDown.Location = New System.Drawing.Point(269, 294)
        Me.picDropDown.Name = "picDropDown"
        Me.picDropDown.Size = New System.Drawing.Size(18, 22)
        Me.picDropDown.TabIndex = 14
        Me.picDropDown.TabStop = False
        Me.picDropDown.Visible = False
        '
        'picCloseUp
        '
        Me.picCloseUp.Image = CType(resources.GetObject("picCloseUp.Image"), System.Drawing.Image)
        Me.picCloseUp.Location = New System.Drawing.Point(293, 294)
        Me.picCloseUp.Name = "picCloseUp"
        Me.picCloseUp.Size = New System.Drawing.Size(18, 22)
        Me.picCloseUp.TabIndex = 15
        Me.picCloseUp.TabStop = False
        Me.picCloseUp.Visible = False
        '
        'picUp2
        '
        Me.picUp2.Image = CType(resources.GetObject("picUp2.Image"), System.Drawing.Image)
        Me.picUp2.Location = New System.Drawing.Point(245, 294)
        Me.picUp2.Name = "picUp2"
        Me.picUp2.Size = New System.Drawing.Size(18, 22)
        Me.picUp2.TabIndex = 16
        Me.picUp2.TabStop = False
        Me.picUp2.Visible = False
        '
        'picDrop2
        '
        Me.picDrop2.Image = CType(resources.GetObject("picDrop2.Image"), System.Drawing.Image)
        Me.picDrop2.Location = New System.Drawing.Point(223, 294)
        Me.picDrop2.Name = "picDrop2"
        Me.picDrop2.Size = New System.Drawing.Size(18, 22)
        Me.picDrop2.TabIndex = 17
        Me.picDrop2.TabStop = False
        Me.picDrop2.Visible = False
        '
        'CopyToClipboardLink
        '
        Me.CopyToClipboardLink.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CopyToClipboardLink.AutoSize = True
        Me.CopyToClipboardLink.Location = New System.Drawing.Point(8, 549)
        Me.CopyToClipboardLink.Name = "CopyToClipboardLink"
        Me.CopyToClipboardLink.Size = New System.Drawing.Size(95, 13)
        Me.CopyToClipboardLink.TabIndex = 18
        Me.CopyToClipboardLink.TabStop = True
        Me.CopyToClipboardLink.Text = "Copy To Clipboard"
        '
        'swiExceptionWin
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(472, 573)
        Me.ControlBox = False
        Me.Controls.Add(Me.CopyToClipboardLink)
        Me.Controls.Add(Me.picDrop2)
        Me.Controls.Add(Me.picUp2)
        Me.Controls.Add(Me.picCloseUp)
        Me.Controls.Add(Me.picDropDown)
        Me.Controls.Add(Me.btn4)
        Me.Controls.Add(Me.chkEmail)
        Me.Controls.Add(Me.btnMore)
        Me.Controls.Add(Me.txtMore)
        Me.Controls.Add(Me.btn3)
        Me.Controls.Add(Me.btn2)
        Me.Controls.Add(Me.btn1)
        Me.Controls.Add(Me.lblMoreHeading)
        Me.Controls.Add(Me.lblActionHeading)
        Me.Controls.Add(Me.lblScopeHeading)
        Me.Controls.Add(Me.lblErrorHeading)
        Me.Controls.Add(Me.ActionBox)
        Me.Controls.Add(Me.ScopeBox)
        Me.Controls.Add(Me.ErrorBox)
        Me.Controls.Add(Me.PictureBox1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "swiExceptionWin"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "(app) has encountered a problem"
        Me.TopMost = True
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picDropDown, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCloseUp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picUp2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picDrop2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkEmail As System.Windows.Forms.CheckBox
    Friend WithEvents btn4 As System.Windows.Forms.Button
    Friend WithEvents picDropDown As System.Windows.Forms.PictureBox
    Friend WithEvents picCloseUp As System.Windows.Forms.PictureBox
    Friend WithEvents picUp2 As System.Windows.Forms.PictureBox
    Friend WithEvents picDrop2 As System.Windows.Forms.PictureBox
    Friend WithEvents CopyToClipboardLink As System.Windows.Forms.LinkLabel

End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMain
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents chkEquipes As System.Windows.Forms.CheckBox
	Public WithEvents chkCompteurs As System.Windows.Forms.CheckBox
    Public WithEvents chkHoraire As System.Windows.Forms.CheckBox
    Public WithEvents chkClassement As System.Windows.Forms.CheckBox
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents lblAction As System.Windows.Forms.Label
    '	Public WithEvents optEquipes As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.optToutesLes�quipes = New System.Windows.Forms.RadioButton()
        Me.opt2�quipe = New System.Windows.Forms.RadioButton()
        Me.chkEquipes = New System.Windows.Forms.CheckBox()
        Me.chkCompteurs = New System.Windows.Forms.CheckBox()
        Me.chkHoraire = New System.Windows.Forms.CheckBox()
        Me.chkClassement = New System.Windows.Forms.CheckBox()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.lblAction = New System.Windows.Forms.Label()
        Me.txtBaseDedonn�es = New System.Windows.Forms.TextBox()
        Me.lblBaseDeDonn�es = New System.Windows.Forms.Label()
        Me.txtR�pertoireSiteWeb = New System.Windows.Forms.TextBox()
        Me.lblR�pertoireSiteWeb = New System.Windows.Forms.Label()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.optToutesLes�quipes)
        Me.Frame1.Controls.Add(Me.opt2�quipe)
        Me.Frame1.Controls.Add(Me.chkEquipes)
        Me.Frame1.Controls.Add(Me.chkCompteurs)
        Me.Frame1.Controls.Add(Me.chkHoraire)
        Me.Frame1.Controls.Add(Me.chkClassement)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(5, 6)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(325, 187)
        Me.Frame1.TabIndex = 3
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Options"
        '
        'optToutesLes�quipes
        '
        Me.optToutesLes�quipes.AutoSize = True
        Me.optToutesLes�quipes.Location = New System.Drawing.Point(40, 151)
        Me.optToutesLes�quipes.Name = "optToutesLes�quipes"
        Me.optToutesLes�quipes.Size = New System.Drawing.Size(154, 17)
        Me.optToutesLes�quipes.TabIndex = 11
        Me.optToutesLes�quipes.TabStop = True
        Me.optToutesLes�quipes.Text = "Toutes les pages d'�quipes"
        Me.optToutesLes�quipes.UseVisualStyleBackColor = True
        '
        'opt2�quipe
        '
        Me.opt2�quipe.AutoSize = True
        Me.opt2�quipe.Location = New System.Drawing.Point(40, 125)
        Me.opt2�quipe.Name = "opt2�quipe"
        Me.opt2�quipe.Size = New System.Drawing.Size(244, 17)
        Me.opt2�quipe.TabIndex = 10
        Me.opt2�quipe.TabStop = True
        Me.opt2�quipe.Text = "Seulement les 2 derni�res �quipes qui ont jou�"
        Me.opt2�quipe.UseVisualStyleBackColor = True
        '
        'chkEquipes
        '
        Me.chkEquipes.BackColor = System.Drawing.SystemColors.Control
        Me.chkEquipes.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEquipes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkEquipes.Location = New System.Drawing.Point(12, 95)
        Me.chkEquipes.Name = "chkEquipes"
        Me.chkEquipes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEquipes.Size = New System.Drawing.Size(216, 24)
        Me.chkEquipes.TabIndex = 9
        Me.chkEquipes.Text = "Mettre � jour les pages d'�quipe"
        Me.chkEquipes.UseVisualStyleBackColor = False
        '
        'chkCompteurs
        '
        Me.chkCompteurs.BackColor = System.Drawing.SystemColors.Control
        Me.chkCompteurs.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCompteurs.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCompteurs.Location = New System.Drawing.Point(12, 71)
        Me.chkCompteurs.Name = "chkCompteurs"
        Me.chkCompteurs.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCompteurs.Size = New System.Drawing.Size(258, 24)
        Me.chkCompteurs.TabIndex = 8
        Me.chkCompteurs.Text = "Mettre � jour les compteurs"
        Me.chkCompteurs.UseVisualStyleBackColor = False
        '
        'chkHoraire
        '
        Me.chkHoraire.BackColor = System.Drawing.SystemColors.Control
        Me.chkHoraire.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkHoraire.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkHoraire.Location = New System.Drawing.Point(12, 41)
        Me.chkHoraire.Name = "chkHoraire"
        Me.chkHoraire.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkHoraire.Size = New System.Drawing.Size(211, 30)
        Me.chkHoraire.TabIndex = 5
        Me.chkHoraire.Text = "Mettre � jour le calendrier"
        Me.chkHoraire.UseVisualStyleBackColor = False
        '
        'chkClassement
        '
        Me.chkClassement.BackColor = System.Drawing.SystemColors.Control
        Me.chkClassement.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkClassement.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkClassement.Location = New System.Drawing.Point(12, 18)
        Me.chkClassement.Name = "chkClassement"
        Me.chkClassement.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkClassement.Size = New System.Drawing.Size(196, 23)
        Me.chkClassement.TabIndex = 4
        Me.chkClassement.Text = "Mettre � jour le classement"
        Me.chkClassement.UseVisualStyleBackColor = False
        '
        'Command1
        '
        Me.Command1.BackColor = System.Drawing.SystemColors.Control
        Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command1.Location = New System.Drawing.Point(89, 306)
        Me.Command1.Name = "Command1"
        Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command1.Size = New System.Drawing.Size(159, 48)
        Me.Command1.TabIndex = 0
        Me.Command1.Text = "G�n�rer Pages Web"
        Me.Command1.UseVisualStyleBackColor = False
        '
        'lblAction
        '
        Me.lblAction.BackColor = System.Drawing.SystemColors.Control
        Me.lblAction.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAction.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAction.Location = New System.Drawing.Point(70, 177)
        Me.lblAction.Name = "lblAction"
        Me.lblAction.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAction.Size = New System.Drawing.Size(283, 16)
        Me.lblAction.TabIndex = 2
        '
        'txtBaseDedonn�es
        '
        Me.txtBaseDedonn�es.Location = New System.Drawing.Point(2, 230)
        Me.txtBaseDedonn�es.Name = "txtBaseDedonn�es"
        Me.txtBaseDedonn�es.Size = New System.Drawing.Size(328, 20)
        Me.txtBaseDedonn�es.TabIndex = 4
        '
        'lblBaseDeDonn�es
        '
        Me.lblBaseDeDonn�es.AutoSize = True
        Me.lblBaseDeDonn�es.Location = New System.Drawing.Point(2, 208)
        Me.lblBaseDeDonn�es.Name = "lblBaseDeDonn�es"
        Me.lblBaseDeDonn�es.Size = New System.Drawing.Size(90, 13)
        Me.lblBaseDeDonn�es.TabIndex = 5
        Me.lblBaseDeDonn�es.Text = "Base de donn�es"
        '
        'txtR�pertoireSiteWeb
        '
        Me.txtR�pertoireSiteWeb.Location = New System.Drawing.Point(2, 280)
        Me.txtR�pertoireSiteWeb.Name = "txtR�pertoireSiteWeb"
        Me.txtR�pertoireSiteWeb.Size = New System.Drawing.Size(328, 20)
        Me.txtR�pertoireSiteWeb.TabIndex = 6
        '
        'lblR�pertoireSiteWeb
        '
        Me.lblR�pertoireSiteWeb.AutoSize = True
        Me.lblR�pertoireSiteWeb.Location = New System.Drawing.Point(2, 261)
        Me.lblR�pertoireSiteWeb.Name = "lblR�pertoireSiteWeb"
        Me.lblR�pertoireSiteWeb.Size = New System.Drawing.Size(98, 13)
        Me.lblR�pertoireSiteWeb.TabIndex = 7
        Me.lblR�pertoireSiteWeb.Text = "R�pertoire site web"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(336, 365)
        Me.Controls.Add(Me.lblR�pertoireSiteWeb)
        Me.Controls.Add(Me.txtR�pertoireSiteWeb)
        Me.Controls.Add(Me.lblBaseDeDonn�es)
        Me.Controls.Add(Me.txtBaseDedonn�es)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Command1)
        Me.Controls.Add(Me.lblAction)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "frmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "GEH Statcan - Cr�ation de pages web "
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents opt2�quipe As RadioButton
    Friend WithEvents optToutesLes�quipes As RadioButton
    Friend WithEvents txtBaseDedonn�es As TextBox
    Friend WithEvents lblBaseDeDonn�es As Label
    Friend WithEvents txtR�pertoireSiteWeb As TextBox
    Friend WithEvents lblR�pertoireSiteWeb As Label
#End Region
End Class
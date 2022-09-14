<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FManageFP
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
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

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
		Me.btnLoadFiles = New System.Windows.Forms.Button()
		Me.tbAdressDirForLoad = New System.Windows.Forms.TextBox()
		Me.lName = New System.Windows.Forms.Label()
		Me.lPsw = New System.Windows.Forms.Label()
		Me.lPath = New System.Windows.Forms.Label()
		Me.tbUserName = New System.Windows.Forms.TextBox()
		Me.tbPSW = New System.Windows.Forms.TextBox()
		Me.lCNString = New System.Windows.Forms.Label()
		Me.mtbCnString = New System.Windows.Forms.MaskedTextBox()
		Me.SuspendLayout()
		'
		'btnLoadFiles
		'
		Me.btnLoadFiles.BackColor = System.Drawing.SystemColors.ButtonShadow
		Me.btnLoadFiles.FlatStyle = System.Windows.Forms.FlatStyle.System
		Me.btnLoadFiles.Location = New System.Drawing.Point(102, 150)
		Me.btnLoadFiles.Name = "btnLoadFiles"
		Me.btnLoadFiles.Size = New System.Drawing.Size(111, 23)
		Me.btnLoadFiles.TabIndex = 0
		Me.btnLoadFiles.Text = "LoadData"
		Me.btnLoadFiles.UseVisualStyleBackColor = False
		'
		'tbAdressDirForLoad
		'
		Me.tbAdressDirForLoad.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystemDirectories
		Me.tbAdressDirForLoad.Location = New System.Drawing.Point(135, 103)
		Me.tbAdressDirForLoad.Name = "tbAdressDirForLoad"
		Me.tbAdressDirForLoad.Size = New System.Drawing.Size(285, 20)
		Me.tbAdressDirForLoad.TabIndex = 1
		Me.tbAdressDirForLoad.Text = "h:\Работа\ИнтернетТрейдинг\Выписка\дебаг\"
		'
		'lName
		'
		Me.lName.AutoSize = True
		Me.lName.Location = New System.Drawing.Point(39, 46)
		Me.lName.Name = "lName"
		Me.lName.Size = New System.Drawing.Size(57, 13)
		Me.lName.TabIndex = 2
		Me.lName.Text = "UserName"
		'
		'lPsw
		'
		Me.lPsw.AutoSize = True
		Me.lPsw.Location = New System.Drawing.Point(39, 78)
		Me.lPsw.Name = "lPsw"
		Me.lPsw.Size = New System.Drawing.Size(53, 13)
		Me.lPsw.TabIndex = 3
		Me.lPsw.Text = "Password"
		'
		'lPath
		'
		Me.lPath.AutoSize = True
		Me.lPath.Location = New System.Drawing.Point(39, 106)
		Me.lPath.Name = "lPath"
		Me.lPath.Size = New System.Drawing.Size(81, 13)
		Me.lPath.TabIndex = 4
		Me.lPath.Text = "PathForLoadDir"
		'
		'tbUserName
		'
		Me.tbUserName.Location = New System.Drawing.Point(135, 46)
		Me.tbUserName.MaxLength = 50
		Me.tbUserName.Name = "tbUserName"
		Me.tbUserName.Size = New System.Drawing.Size(100, 20)
		Me.tbUserName.TabIndex = 5
		Me.tbUserName.Text = "Rinat"
		'
		'tbPSW
		'
		Me.tbPSW.Location = New System.Drawing.Point(135, 75)
		Me.tbPSW.MaxLength = 24
		Me.tbPSW.Name = "tbPSW"
		Me.tbPSW.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
		Me.tbPSW.Size = New System.Drawing.Size(100, 20)
		Me.tbPSW.TabIndex = 6
		Me.tbPSW.UseSystemPasswordChar = True
		'
		'lCNString
		'
		Me.lCNString.AutoSize = True
		Me.lCNString.Location = New System.Drawing.Point(42, 13)
		Me.lCNString.Name = "lCNString"
		Me.lCNString.Size = New System.Drawing.Size(47, 13)
		Me.lCNString.TabIndex = 7
		Me.lCNString.Text = "CnString"
		'
		'mtbCnString
		'
		Me.mtbCnString.Location = New System.Drawing.Point(135, 10)
		Me.mtbCnString.Name = "mtbCnString"
		Me.mtbCnString.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
		Me.mtbCnString.Size = New System.Drawing.Size(774, 20)
		Me.mtbCnString.TabIndex = 8
		Me.mtbCnString.Text = "Data Source=RINATS\SQLEXPRESS; Initial Catalog=MyBag; Integrated Security=False; Application Name=ManagerFP"
		Me.mtbCnString.UseSystemPasswordChar = True
		'
		'FManageFP
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(1491, 634)
		Me.Controls.Add(Me.mtbCnString)
		Me.Controls.Add(Me.lCNString)
		Me.Controls.Add(Me.tbPSW)
		Me.Controls.Add(Me.tbUserName)
		Me.Controls.Add(Me.lPath)
		Me.Controls.Add(Me.lPsw)
		Me.Controls.Add(Me.lName)
		Me.Controls.Add(Me.tbAdressDirForLoad)
		Me.Controls.Add(Me.btnLoadFiles)
		Me.Name = "FManageFP"
		Me.Text = "Form1"
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents btnLoadFiles As Button
	Friend WithEvents tbAdressDirForLoad As TextBox
	Friend WithEvents lName As Label
	Friend WithEvents lPsw As Label
	Friend WithEvents lPath As Label
	Friend WithEvents tbUserName As TextBox
	Friend WithEvents tbPSW As TextBox
	Friend WithEvents lCNString As Label
	Friend WithEvents mtbCnString As MaskedTextBox
End Class

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
		Me.SuspendLayout()
		'
		'btnLoadFiles
		'
		Me.btnLoadFiles.BackColor = System.Drawing.SystemColors.ButtonShadow
		Me.btnLoadFiles.FlatStyle = System.Windows.Forms.FlatStyle.System
		Me.btnLoadFiles.Location = New System.Drawing.Point(96, 96)
		Me.btnLoadFiles.Name = "btnLoadFiles"
		Me.btnLoadFiles.Size = New System.Drawing.Size(111, 23)
		Me.btnLoadFiles.TabIndex = 0
		Me.btnLoadFiles.Text = "LoadData"
		Me.btnLoadFiles.UseVisualStyleBackColor = False
		'
		'tbAdressDirForLoad
		'
		Me.tbAdressDirForLoad.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystemDirectories
		Me.tbAdressDirForLoad.Location = New System.Drawing.Point(42, 54)
		Me.tbAdressDirForLoad.Name = "tbAdressDirForLoad"
		Me.tbAdressDirForLoad.Size = New System.Drawing.Size(285, 20)
		Me.tbAdressDirForLoad.TabIndex = 1
		Me.tbAdressDirForLoad.Text = "h:\Работа\ИнтернетТрейдинг\Выписка\дебаг\"
		'
		'FManageFP
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(1491, 634)
		Me.Controls.Add(Me.tbAdressDirForLoad)
		Me.Controls.Add(Me.btnLoadFiles)
		Me.Name = "FManageFP"
		Me.Text = "Form1"
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents btnLoadFiles As Button
    Friend WithEvents tbAdressDirForLoad As TextBox
End Class

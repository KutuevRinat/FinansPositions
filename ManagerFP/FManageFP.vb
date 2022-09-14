Public Class FManageFP
  Private Sub btnLoadFiles_Click(sender As Object, e As EventArgs) Handles btnLoadFiles.Click
		Dim Loader As ParentXlLoader
		Loader = New ParentXlLoader(mtbCnString.Text + "; User Id = " + tbUserName.Text + "; Password = " + tbPSW.Text, tbAdressDirForLoad.Text)
		Loader.Load()
		Loader = Nothing
		'Loader.OrdererFilesReport(tbAdressDirForLoad.Text)
		'Stop
		'FManageFP.ActiveForm.Show()
		'FManageFP.ActiveForm.Close()
	End Sub

  Private Sub tbAdressDirForLoad_TextChanged(sender As Object, e As EventArgs) Handles tbAdressDirForLoad.TextChanged

    End Sub

	Private Sub FManageFP_Load(sender As Object, e As EventArgs) Handles MyBase.Load

	End Sub


End Class

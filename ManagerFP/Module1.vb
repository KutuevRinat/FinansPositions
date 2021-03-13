Imports System.IO
Imports Microsoft.Office.Interop
Module Loader
	Public Structure FileAtributes
		Dim FileName As String
		Dim FileDate As Date
		Dim Broker As String
		Dim Accnt As String
		Dim Market As String
		'Dim Period As String
	End Structure

	'Sub Main()
	'	Dim P As ParentBrokerReportLoader
	'	P = New ParentBrokerReportLoader("p", "f", "a", "m", "b")
	'	Debug.Print(P.FileName)
	'	'Dim oXL As Excel.Application = Nothing
	'	'oXL = New excel.Application
	'	'OrdererFilesReport("h:\Работа\ИнтернетТрейдинг\Выписка\дебаг\")
	'End Sub
	Sub OrdererFilesReport(pathDirReports As String) 'упорядочивает файлы-отчеты в директории по датам и запускает загрузчики 

    Dim DirReports As IO.DirectoryInfo = New IO.DirectoryInfo(pathDirReports) 'объект директории 
    Dim FileReport As IO.FileInfo() = DirReports.GetFiles()
    Dim q As Integer 'счетчик файлов
    Dim qq As Integer 'вспомогательный счетчик файлов для упорядочивания
    Dim Up_list_files(150) As String 'массив имен файлов в директории
    'Dim Data_list_files(150) As Date 'массив дат последних изменений в директориях
    Dim Files(0 To 150) As FileAtributes
    Const conKitFinanceBroker = "КИТ Финанс(АО)"
    Const conSberbankBroker = "ПАО Сбербанк"
    Const conBCS = "ООО Компания БКС"
		Dim Loader As ParentBrokerReportLoader

		For Each f As IO.FileInfo In FileReport
			If Right(f.Name, Len(f.Name) - InStrRev(f.Name, ".")) = "xlsx" _
			Or Right(f.Name, Len(f.Name) - InStrRev(f.Name, ".")) = "xls" _
			Or Right(f.Name, Len(f.Name) - InStrRev(f.Name, ".")) = "html" Then
				If Mid(f.Name, 1, 6) = "forts1" Then
					Files(q).FileName = f.Name
					Files(q).FileDate = CDate(Mid(f.Name, 14, 10))
					Files(q).Broker = conKitFinanceBroker
					Files(q).Accnt = Mid(f.Name, 6, 7)
					Files(q).Market = "Срочный рынок"
					q = q + 1
					'Files(q).Period = "d"
				ElseIf Mid(f.Name, 1, 6) = "Forts_" Then
					Files(q).FileName = f.Name
					Files(q).FileDate = CDate(Mid(f.Name, 27, 10))
					Files(q).Broker = conKitFinanceBroker
					Files(q).Accnt = Mid(f.Name, 19, 7)
					Files(q).Market = "Срочный рынок"
					q = q + 1
					'Files(q).Period = "d"
				ElseIf Mid(f.Name, 13, 5) = "17449" Then
					Files(q).FileName = f.Name
					Files(q).FileDate = CDate(Mid(f.Name, 19, 2) + "." + Mid(f.Name, 21, 2) + "." + Mid(f.Name, 23, 2))
					Files(q).Broker = conKitFinanceBroker
					Files(q).Accnt = Mid(f.Name, 13, 5)
					Files(q).Market = "Фондовый рынок"
					q = q + 1
					'Files(q).Period = "d"
				ElseIf Mid(f.Name, 15, 5) = "17449" Then
					Files(q).FileName = f.Name
					Files(q).FileDate = CDate(Mid(f.Name, 27, 10))
					Files(q).Broker = conKitFinanceBroker
					Files(q).Accnt = Mid(f.Name, 15, 5)
					Files(q).Market = "Фондовый рынок"
					q = q + 1
				ElseIf Mid(f.Name, 1, 5) = "4HS4Y" Then
					Files(q).FileName = f.Name
					Files(q).FileDate = CDate(Mid(f.Name, 14, 2) + "." + Mid(f.Name, 16, 2) + "." + Mid(f.Name, 18, 2))
					Files(q).Broker = conSberbankBroker
					Files(q).Accnt = Mid(f.Name, 1, 5)
					Files(q).Market = "Фондовый рынок"
					q = q + 1
				ElseIf Mid(f.Name, 1, 3) = "B_k" And Len(f.Name) = 24 Then
					Files(q).FileName = f.Name
					Files(q).FileDate = CDate("20" + Mid(f.Name, 16, 2) + "." + Mid(f.Name, 19, 2) + "." + "01")
					Files(q).Broker = conBCS
					Files(q).Accnt = Mid(f.Name, 5, 6)
					Files(q).Market = "Фондовый рынок"
					q = q + 1
				End If


			End If
		Next
		'упорядочиваем по дате создания
		Dim Fqq As FileAtributes
    For qq = 0 To q - 2
      If Files(qq).FileDate > Files(qq + 1).FileDate Then
        Fqq = Files(qq)
        Files(qq) = Files(qq + 1)
        Files(qq + 1) = Fqq
        If qq > 0 Then qq = qq - 2
      End If
    Next qq
		Loader = New ParentBrokerReportLoader(pathDirReports, "f", "2020-01-01", "a", "m", "b")
		For qq = 0 To q - 1
			Loader.FileName = Files(qq).FileName
			Loader.FileDate = Files(qq).FileDate
			Loader.Broker = Files(qq).Broker
			Loader.Account = Files(qq).Accnt
			Loader.Market = Files(qq).Market
			Loader.Load
		Next qq
		Loader = Nothing


	End Sub

End Module

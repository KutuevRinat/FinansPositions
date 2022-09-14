Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Public Class ParentBrokerReportLoader

	'Protected Fl As FileAtributes

	'характеристики файла-отчета
	Protected DirName As String
	Protected FName As String
	Protected FDate As DateTime
	Protected Accnt As String 'номер счета
	Protected Mrkt As String 'секция рынка
	Protected Brk As String 'брокер

	Friend Rsp As String 'строка сообщений

	Protected oXL As Excel.Application = Nothing

	Protected strCn
	Protected strCnUser As String

	Protected cnInsert As SqlConnection 'соединение с БД для вставки строк
	Protected cmdInsert As SqlCommand
	Protected cnSpProc As SqlConnection 'соединение с БД для вызова встроенных процедур
	Protected cmdSpProc As SqlCommand

	Protected Structure ImpFunct
		Dim NImpFunct As String
		Dim vImpFunct As Boolean
	End Structure

	Protected aImpFunct(0 To 10) As ImpFunct

	Protected Structure FieldsNClmn
		Dim Fld As String
		Dim NClmn As Integer
	End Structure

	Protected Structure FieldsData
		Dim Fld As String
		Dim FldData As String
	End Structure

	Public Property ConnectionString() As String
		Protected Get
			Return ManagerFP.My.Settings.CnStrSQLClient
		End Get
		Set(ConStringValue As String)
			strCnUser = ConStringValue
			If strCnUser IsNot Nothing And strCnUser <> "" Then
				strCn = strCnUser
			Else
				strCnUser = ManagerFP.My.Settings.CnStrSQLClient
				strCn = ManagerFP.My.Settings.CnStrSQLClient + "; User ID=Rinat; Password = 'RInnat'"
			End If
		End Set
	End Property

	Public Property PathToFile() As String
		Get
			Return DirName
		End Get
		Set(PathStr As String)
			DirName = PathStr
		End Set
	End Property

	Public Property FileName() As String
		Get
			Return FName
		End Get
		Set(FilName As String)
			FName = FilName
		End Set
	End Property

	Public Property FileDate() As Date
		Get
			Return FDate
		End Get
		Set(value As Date)
			FDate = value
		End Set
	End Property

	Public Property Account() As String
		Get
			Return Accnt
		End Get
		Set(value As String)
			Accnt = value
		End Set
	End Property

	Public Property Market() As String
		Get
			Return Mrkt
		End Get
		Set(value As String)
			Mrkt = value
		End Set
	End Property

	Public Property Broker() As String
		Get
			Return Brk
		End Get
		Set(value As String)
			Brk = value
		End Set
	End Property

	Public Sub New(ByVal PathToFile As String, ByVal FileName As String, ByVal FileDate As Date, ByVal Account As String, ByVal Market As String, ByVal Broker As String)
		Me.New("", PathToFile, FileName, FileDate, Account, Market, Broker)
		'Me.New("DRIVER=MSOLEDBSQL; SERVER=RINATS\SQLEXPRESS; Description = MyBag; UID = Ринат; WSID = RINATS; APP=ManagerFP", PathToFile, FileName, FileDate, Account, Market, Broker)
	End Sub

	Public Sub New(ByVal ConnectionString As String, ByVal PathToFile As String, ByVal FileName As String, ByVal FileDate As Date, ByVal Account As String, ByVal Market As String, ByVal Broker As String)

		oXL = New Excel.Application

		DirName = PathToFile
		FName = FileName
		FDate = FileDate
		Accnt = Account
		Mrkt = Market
		Brk = Broker
		strCnUser = ConnectionString
		If strCnUser <> "" Then
			strCn = strCnUser
		Else
			strCnUser = ManagerFP.My.Settings.CnStrSQLClient
			strCn = ManagerFP.My.Settings.CnStrSQLClient + "; User ID=Rinat; Password = 'RInnat'"
		End If

		'контролер для запуска процедур базы данных
		aImpFunct(0).NImpFunct = "spImpPrices"
		aImpFunct(0).vImpFunct = True
		aImpFunct(1).NImpFunct = "spImpOtherStockAct"
		aImpFunct(1).vImpFunct = True
		aImpFunct(2).NImpFunct = "spImpOtherDerivAct"
		aImpFunct(2).vImpFunct = True
		aImpFunct(3).NImpFunct = "spImpCurAct"
		aImpFunct(3).vImpFunct = True
		aImpFunct(4).NImpFunct = "spImpStockAct"
		aImpFunct(4).vImpFunct = True
		aImpFunct(5).NImpFunct = "spImpRedemption"
		aImpFunct(5).vImpFunct = True
		aImpFunct(6).NImpFunct = "spImpDvdCpnPartP"
		aImpFunct(6).vImpFunct = True
		aImpFunct(7).NImpFunct = "spImpFutureAct"
		aImpFunct(7).vImpFunct = True
		aImpFunct(8).NImpFunct = "spImpOptionAct"
		aImpFunct(8).vImpFunct = False
		aImpFunct(9).NImpFunct = "spImpFutureRecalc"
		aImpFunct(9).vImpFunct = True
		aImpFunct(10).NImpFunct = "spImpOtherSecurAct"
		aImpFunct(10).vImpFunct = True

	End Sub

	Protected Overrides Sub Finalize()
		If oXL IsNot Nothing Then
			oXL.Application.ScreenUpdating = True
			oXL.Application.DecimalSeparator = ","
			oXL.Application.UseSystemSeparators = True
			oXL.Quit()
			oXL = Nothing
		End If

	End Sub

	Public Sub Load()
		'Dim Response As Long
		Me.OpenFile()
		Me.RecipientTblNamesStructure()
		Me.CloseFile()
		'cnImp2 = Nothing

		Me.ProcessImpTables()

		'If RTrim(Rsp) <> "Успешная загрузка" Then
		'	Response = MsgBox(Rsp, vbOKOnly)
		'	Stop
		'Else
		Me.MoveFile()
		'End If
	End Sub

	Public Sub OpenFile()
		'oXL = New Excel.Application
		With oXL
			If Right(FName, Len(FName) - InStrRev(FName, ".")) = "html" Then
				.Application.DecimalSeparator = "."
				.Application.UseSystemSeparators = False
			End If
			.Workbooks.Open(DirName & FName)
			.Application.ReferenceStyle = Excel.XlReferenceStyle.xlR1C1
		End With
	End Sub

	Public Sub CloseFile()
		With oXL
			.DecimalSeparator = ","
			.Workbooks(FName).Close(True)
			'.ScreenUpdating = True
			'.DecimalSeparator = ","
			'.UseSystemSeparators = True
		End With
		'oXL.Quit()
		'oXL = Nothing
	End Sub

	Protected Sub RecipientTblNamesStructure()
		'запрашивает из базы данных таблицы-получатели, таблицы-источники и их свойства(обязательна ли таблица, 
		'разделители строк, начало и конец таблицы и данных в ней.
		'Далее вызывает функцию, определяющую соотвествие между загружаемыми полями и данными источников
		Dim ImpTbl As String
		Dim ExpTbl As String
		Dim Alws As Boolean
		Dim Brdr As String
		Dim DtInClmn As Boolean
		Dim TblBeg As Integer
		Dim DtBeg As Integer
		Dim DtEnd As Integer
		Dim SeparBeg As String
		Dim SeparEnd As String
		Dim Skip As Boolean

		'подготовка к выборке наименований таблиц назначений и таблиц источников
		'готовим подключение
		Dim cnReadTblNames As New SqlConnection(strCn) 'соединение с БД чтение 
		'готовим объект команды
		Dim cmdReadTblNames As New SqlCommand()
		cmdReadTblNames.Connection = cnReadTblNames
		cmdReadTblNames.CommandType = CommandType.StoredProcedure
		cmdReadTblNames.CommandText = "spImpTblNames"

		'готовим параметры команды
		Dim prmtReadTblNames As New SqlParameter()
		prmtReadTblNames.ParameterName = "@Broker"
		prmtReadTblNames.SqlDbType = SqlDbType.NChar
		prmtReadTblNames.Size = 24
		prmtReadTblNames.Direction = ParameterDirection.Input
		prmtReadTblNames.Value = Brk
		cmdReadTblNames.Parameters.Add(prmtReadTblNames)

		prmtReadTblNames = New SqlParameter()
		prmtReadTblNames.ParameterName = "@Market"
		prmtReadTblNames.SqlDbType = SqlDbType.NChar
		prmtReadTblNames.Size = 24
		prmtReadTblNames.Direction = ParameterDirection.Input
		prmtReadTblNames.Value = Mrkt
		cmdReadTblNames.Parameters.Add(prmtReadTblNames)

		prmtReadTblNames = New SqlParameter()
		prmtReadTblNames.ParameterName = "@DFrom"
		prmtReadTblNames.SqlDbType = SqlDbType.DateTime
		prmtReadTblNames.Direction = ParameterDirection.Input
		prmtReadTblNames.Value = FDate
		cmdReadTblNames.Parameters.Add(prmtReadTblNames)

		cnReadTblNames.Open()
		Dim rdTblNames As SqlDataReader = cmdReadTblNames.ExecuteReader()
		'чтение наименований таблиц назначений и таблиц источников
		While (rdTblNames.Read())
			ImpTbl = rdTblNames("ImpTable").ToString.Trim
			ExpTbl = rdTblNames("ExpTable").ToString.Trim
			Alws = rdTblNames("Always")
			Brdr = rdTblNames("Brdr").ToString.Trim
			DtInClmn = rdTblNames("DateInColumn")
			TblBeg = rdTblNames("TblBeg")
			DtBeg = rdTblNames("DataBeg")
			If IsDBNull(rdTblNames("DataEnd")) Then DtEnd = 0 Else DtEnd = rdTblNames("DataEnd")
			If IsDBNull(rdTblNames("SeparBeg")) Then SeparBeg = "" Else SeparBeg = rdTblNames("SeparBeg").ToString.Trim
			If IsDBNull(rdTblNames("SeparEnd")) Then SeparEnd = "" Else SeparEnd = rdTblNames("SeparBeg").ToString.Trim
			Skip = rdTblNames("Skip")


			Me.ComplianserFieldClmn(ImpTbl, ExpTbl, TblBeg, Alws, Brdr, DtInClmn, SeparBeg, SeparEnd, DtBeg, DtEnd, Skip)
		End While
		rdTblNames.Close()
		cnReadTblNames.Close()
	End Sub

	Protected Sub ComplianserFieldClmn(ByRef ImpTbl As String, ByRef ExpTbl As String, ByVal TblBeg As Integer, ByRef Always As Boolean, ByRef Brdr As String, ByRef DtInClmn As Boolean, ByRef SeparBeg As String, ByRef SeparEnd As String, ByVal DtBeg As Integer, ByVal DtEnd As Integer, ByRef Skip As Boolean)
		'запрашивает из базы данных соотвествие полей таблицы-адресата и таблицы источника, заполняет массив соответствия полей и номеров столбцов
		'Далее вызывает процедуру-вставки данных в следующей последовательности:
		'Поименованные строки данных, строки данных с границами, обычные строки данных.
		'Из границ строк данных считываем содержащиеся в них поля в массив

		Dim rngNameTbl As Excel.Range 'ячейка с именем таблицы 
		Dim rngTbl As Excel.Range 'таблица
		Dim rngLastCell As Excel.Range
		Dim aFieldsNClmns() As FieldsNClmn 'массив с названиями полей и соответствующим им номерам столбцов/строк в таблице источнике
		Dim Response As Long
		Dim NBeg As Integer
		Dim NEnd As Integer
		Dim i As Integer 'счетчик полей начиная с 0
		Dim k As Integer 'счетчик полей в разделителе
		Dim r As Excel.Range
		Dim rngRow As Excel.Range 'строка данных

		'подготовка к выборке структуры таблиц назначений и соответствия с полями таблиц источников
		'готовим подключение
		Dim cnReadTblStrct As New SqlConnection(strCn) 'соединение с БД чтение 
		'готовим объект команды
		Dim cmdReadTblStrct As New SqlCommand
		cmdReadTblStrct.Connection = cnReadTblStrct
		cmdReadTblStrct.CommandType = CommandType.StoredProcedure
		cmdReadTblStrct.CommandText = "spImpTblStrct"
		'готовим параметры команды
		Dim prmtReadTblStrct As New SqlParameter()
		prmtReadTblStrct.ParameterName = "@Broker"
		prmtReadTblStrct.SqlDbType = SqlDbType.NChar
		prmtReadTblStrct.Size = 24
		prmtReadTblStrct.Direction = ParameterDirection.Input
		prmtReadTblStrct.Value = Brk
		cmdReadTblStrct.Parameters.Add(prmtReadTblStrct)

		prmtReadTblStrct = New SqlParameter()
		prmtReadTblStrct.ParameterName = "@DFrom"
		prmtReadTblStrct.SqlDbType = SqlDbType.DateTime
		prmtReadTblStrct.Direction = ParameterDirection.Input
		prmtReadTblStrct.Value = FDate
		cmdReadTblStrct.Parameters.Add(prmtReadTblStrct)

		prmtReadTblStrct = New SqlParameter()
		prmtReadTblStrct.ParameterName = "@ImpTable"
		prmtReadTblStrct.SqlDbType = SqlDbType.NChar
		prmtReadTblStrct.Size = 24
		prmtReadTblStrct.Direction = ParameterDirection.Input
		prmtReadTblStrct.Value = ImpTbl
		cmdReadTblStrct.Parameters.Add(prmtReadTblStrct)

		prmtReadTblStrct = New SqlParameter()
		prmtReadTblStrct.ParameterName = "@ExpTable"
		prmtReadTblStrct.SqlDbType = SqlDbType.NChar
		prmtReadTblStrct.Size = 55
		prmtReadTblStrct.Direction = ParameterDirection.Input
		prmtReadTblStrct.Value = ExpTbl
		cmdReadTblStrct.Parameters.Add(prmtReadTblStrct)

		'заполняет массив с названиями полей и номерами столбцов/строк в таблице источнике
		With oXL.Workbooks(FName).Worksheets(1)
			' находим ячейку с названием таблицы источника
			' проверка наличия таблицы источника и ее обязательности
			rngNameTbl = .Cells.Find(ExpTbl, , , LookAt:=Excel.XlLookAt.xlWhole)
			If Always And rngNameTbl Is Nothing Then
				Response = MsgBox("Изменилось название таблицы " & ExpTbl & " в " & FName & ", дополните таблицу ImpTblNames новым названием таблицы, потом догрузите файл", vbOK)
				Stop
			End If

			If Not rngNameTbl Is Nothing Then
				rngTbl = .Cells(rngNameTbl.Row + TblBeg, rngNameTbl.Column).CurrentRegion 'определили область таблицы источника как окружение вокруг ячейки на 2 деления ниже названия

				i = 0
				'читаем таблицу соотвествий. Записи должны быть упорядоченны следующим образом:
				'сначала перечень полей, потом поименованных строк, потом границ 
				cnReadTblStrct.Open()
				Dim rdTblStrct As SqlDataReader = cmdReadTblStrct.ExecuteReader()
				While (rdTblStrct.Read)
					'Заполняем массив с полями из структуры таблиц и номерами строк для транспонированных таблиц/номерами столбцов для обычных в таблицах-источниках
					If rdTblStrct("Content").ToString.Trim = "Field" Then
						ReDim Preserve aFieldsNClmns(0 To i)
						aFieldsNClmns(i).Fld = rdTblStrct("Impfield").ToString.Trim
						r = rngTbl.Find(What:=rdTblStrct("ExpField").ToString.Trim, After:=rngTbl(1, 1), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlPart, SearchDirection:=Excel.XlSearchDirection.xlNext)
						If r Is Nothing Then r = rngTbl.Find(What:=rdTblStrct("ExpField").ToString.Trim, After:=rngTbl(rngTbl.Rows.Count, rngTbl.Columns.Count), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlPart, SearchDirection:=Excel.XlSearchDirection.xlPrevious)
						If rdTblStrct("Number") > 1 Then
							For iNum = 2 To rdTblStrct("Number")
								r = rngTbl.FindNext(r)
							Next iNum
						End If
						'В зависимости от того, явялется таблица обычной(данные в строках) или транспонированной(данные в столбцах) заполняем номерами строк или столбцов
						If DtInClmn Then aFieldsNClmns(i).NClmn = r.Row Else aFieldsNClmns(i).NClmn = r.Column
						i = i + 1
					ElseIf rdTblStrct("Content").ToString.Trim = "Name" Then
						'Вставляем записи с поименованными данными
						'While (rdTblStrct.Read)
						rngRow = rngTbl.Find(rdTblStrct("ExpField").ToString.Trim, After:=rngTbl(1, 1), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
						If rngRow Is Nothing Then
							If rdTblStrct("Always") Then
								Response = MsgBox("Изменилось название поля " & rdTblStrct("ExpField").ToString.Trim & " таблицы " & ExpTbl & " в " & FName & ", дополните таблицу ImpTblNames новым названием поля, потом догрузите файл", vbOK)
								Stop
							End If
						Else
							Do
								Do
									If DtInClmn Then NBeg = rngRow.Column Else NBeg = rngRow.Row
									NEnd = NBeg
									Call InsertArroyInTbl(ImpTbl, ExpTbl, DtInClmn, rdTblStrct("ExpField").ToString.Trim, NBeg, NEnd, aFieldsNClmns)
									rngRow = rngTbl.Find(rdTblStrct("ExpField").ToString.Trim, rngRow, , LookAt:=Excel.XlLookAt.xlWhole)
								Loop While (rdTblStrct("Repeat") And NBeg < IIf(DtInClmn, rngRow.Column, rngRow.Row))
								If Skip = True Then
									rngLastCell = rngTbl(rngTbl.Rows.Count, 1)
									rngTbl = .Cells(rngLastCell.Row + 2, rngLastCell.Column).CurrentRegion
									rngRow = rngTbl.Find(rdTblStrct("ExpField").ToString.Trim, After:=rngTbl(1, 1), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
								End If
							Loop While Not rngRow Is Nothing
							rngTbl = .Cells(rngNameTbl.Row + TblBeg, rngNameTbl.Column).CurrentRegion
						End If
					End If
				End While
				rdTblStrct.Close()
				cnReadTblStrct.Close()

				If Brdr = "Separator" Then
					'ElseIf rdTblStrct("Content").ToString.Trim In ("Separator", "SeparatorBeg") Then
					'SeparBeg = rdTblStrct("ExpField").ToString.Trim
					'обрабатываем разрыв в таблице
					If Skip = True Then
						rngLastCell = rngTbl(rngTbl.Rows.Count, 1)
						'Dim ProvercaI As Integer
						rngTbl = .Cells(rngLastCell.Row + 2, rngLastCell.Column).CurrentRegion
						rngTbl = .Range(rngTbl(0, 1), rngTbl(rngTbl.Rows.Count, 1))
					End If
					rngRow = rngTbl.Find(SeparBeg, After:=rngTbl(1, 1), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlPart)
					If rngRow Is Nothing Then
						If rdTblStrct("Always") Then
							Response = MsgBox("Изменился разделитель строк данных " & SeparBeg & " таблицы " & ExpTbl & " в " & FName & ", дополните таблицу ImpTblNames новым разделителем, потом догрузите файл", vbOK)
							Stop
						End If
					Else

						'открываем подключение для выборки значений полей из разделителя
						'готовим подключение
						Dim cnReadTblStrctSepar As New SqlConnection(strCn) 'соединение с БД чтение 
						'готовим объект команды
						Dim cmdReadTblStrctSepar As New SqlCommand
						cmdReadTblStrctSepar.Connection = cnReadTblStrctSepar
						cmdReadTblStrctSepar.CommandType = CommandType.StoredProcedure
						cmdReadTblStrctSepar.CommandText = "spImpTblStrctSepar"
						'готовим параметры команды
						Dim prmtReadTblStrctSepar As New SqlParameter()
						prmtReadTblStrctSepar.ParameterName = "@Broker"
						prmtReadTblStrctSepar.SqlDbType = SqlDbType.NChar
						prmtReadTblStrctSepar.Size = 24
						prmtReadTblStrctSepar.Direction = ParameterDirection.Input
						prmtReadTblStrctSepar.Value = Brk
						cmdReadTblStrctSepar.Parameters.Add(prmtReadTblStrctSepar)

						prmtReadTblStrctSepar = New SqlParameter()
						prmtReadTblStrctSepar.ParameterName = "@DFrom"
						prmtReadTblStrctSepar.SqlDbType = SqlDbType.DateTime
						prmtReadTblStrctSepar.Direction = ParameterDirection.Input
						prmtReadTblStrctSepar.Value = FDate
						cmdReadTblStrctSepar.Parameters.Add(prmtReadTblStrctSepar)

						prmtReadTblStrctSepar = New SqlParameter()
						prmtReadTblStrctSepar.ParameterName = "@ImpTable"
						prmtReadTblStrctSepar.SqlDbType = SqlDbType.NChar
						prmtReadTblStrctSepar.Size = 24
						prmtReadTblStrctSepar.Direction = ParameterDirection.Input
						prmtReadTblStrctSepar.Value = ImpTbl
						cmdReadTblStrctSepar.Parameters.Add(prmtReadTblStrctSepar)

						prmtReadTblStrctSepar = New SqlParameter()
						prmtReadTblStrctSepar.ParameterName = "@ExpTable"
						prmtReadTblStrctSepar.SqlDbType = SqlDbType.NChar
						prmtReadTblStrctSepar.Size = 50
						prmtReadTblStrctSepar.Direction = ParameterDirection.Input
						prmtReadTblStrctSepar.Value = ExpTbl
						cmdReadTblStrctSepar.Parameters.Add(prmtReadTblStrctSepar)

						prmtReadTblStrctSepar = New SqlParameter()
						prmtReadTblStrctSepar.ParameterName = "@ExpField"
						prmtReadTblStrctSepar.SqlDbType = SqlDbType.NChar
						prmtReadTblStrctSepar.Size = 50
						prmtReadTblStrctSepar.Direction = ParameterDirection.Input
						prmtReadTblStrctSepar.Value = SeparBeg
						cmdReadTblStrctSepar.Parameters.Add(prmtReadTblStrctSepar)

						cnReadTblStrctSepar.Open()

						Dim rdTblStrctSepar As SqlDataReader = cmdReadTblStrctSepar.ExecuteReader()
						Dim aSeparFieldsInClmn() As FieldsNClmn
						'читаем данные, содержащиеся в разделителе, заполняем массив c номерами столбцов в разделителе
						k = 0
						While (rdTblStrctSepar.Read)
							ReDim Preserve aSeparFieldsInClmn(0 To k)
							aSeparFieldsInClmn(k).Fld = rdTblStrctSepar("ImpField").ToString.Trim
							aSeparFieldsInClmn(k).NClmn = rdTblStrctSepar("Number")
							k = k + 1
						End While
						rdTblStrctSepar.Close()
						cnReadTblStrctSepar.Close()

						Dim aSeparFieldsData(0 To k - 1) As FieldsData
						Do
							'заполняем массив с данными из разделителя
							For k = 0 To UBound(aSeparFieldsInClmn)
								aSeparFieldsData(k).Fld = aSeparFieldsInClmn(k).Fld
								aSeparFieldsData(k).FldData = CStr(.cells(rngRow.Row, aSeparFieldsInClmn(k).NClmn).value)
							Next
							NBeg = rngRow.Row + 1
							If SeparEnd <> "" Then
								rngRow = rngTbl.Find(SeparEnd, After:= .cells(NBeg, rngRow.Column), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlPart)
							Else
								rngRow = rngTbl.FindNext(rngRow)
							End If

							If rngRow Is Nothing Or rngRow.Row < NBeg Then
								rngLastCell = rngTbl(rngTbl.Rows.Count, rngTbl.Columns.Count)
								NEnd = rngLastCell.Row + DtEnd
								Call InsertArroyInTbl(ImpTbl, ExpTbl, DtInClmn, SeparBeg, NBeg, NEnd, aFieldsNClmns, aSeparFieldsData)
								Exit Do
							Else
								NEnd = rngRow.Row - 1 + DtEnd
								Call InsertArroyInTbl(ImpTbl, ExpTbl, DtInClmn, SeparBeg, NBeg, NEnd, aFieldsNClmns, aSeparFieldsData)
								If SeparEnd <> "" Then
									'обрабатываем разрыв в таблице
									If Skip = True Then
										rngLastCell = rngTbl(rngTbl.Rows.Count, 1)
										rngTbl = .Cells(rngLastCell.Row + 2, rngLastCell.Column).CurrentRegion
										rngTbl = .Range(rngTbl(0, 1), rngTbl(rngTbl.Rows.Count, 1))
									End If
									rngRow = rngTbl.Find(SeparBeg, After:=rngTbl(1, 1), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlPart)
								End If
							End If
						Loop Until rngRow Is Nothing
					End If
				ElseIf Brdr = "Not" Then
					'вставляем записи с непоименованными строками данных
					If DtInClmn Then
						NBeg = rngNameTbl.Column + DtBeg
						NEnd = rngTbl.Columns.Count - 1 + DtEnd
					Else
						NBeg = rngNameTbl.Row + DtBeg 'определили номер строки с началом данных в таблице источнике
						NEnd = NBeg + rngTbl.Rows.Count - 1 + DtEnd
					End If
					Call InsertArroyInTbl(ImpTbl, ExpTbl, DtInClmn, "", NBeg, NEnd, aFieldsNClmns)
				End If
			End If
		End With
	End Sub

	Protected Sub InsertArroyInTbl(ImpTbl As String, ExpTbl As String,
		ByVal DtInClmn As Boolean, ExpFld As String, ByVal NDataBeg As Integer,
																 ByVal NDataEnd As Integer, ByRef aFieldsNClmns() As FieldsNClmn, Optional ByVal aSeparFieldsData() As FieldsData = Nothing)
		'вставляет данные в таблицу базы данных
		Dim strInsert As String
		Dim strValues As String
		Dim SubAccnt As String
		Dim i As Integer
		Dim j As Integer
		Dim NData As Integer
		'Const ImpMoney As String = "ImpMoney"
		Const conImpCurAct As String = "ImpCurAct"
		Const conCurMarket As String = "Валютный рынок"
		Const conSignCurMarket As String = "Торговый счет, Валютный рынок"
		Const conImpBcsDrvtvMove As String = "ImpBcsDerivativeMove"
		Const conDrvtMrkt As String = "Срочный рынок"

		'готовим подключение
		Dim cnInsrt As New ADODB.Connection 'соединение с БД чтение 
		cnInsrt.ConnectionString = ManagerFP.My.Settings.CnStrMSOLEDBSQL + "; User ID=Rinat; Password = 'RInnat'"
		cnInsrt.Open() 'готовим объект команды
		Dim cmdInsrt As New ADODB.Command
		cmdInsrt.ActiveConnection = cnInsrt
		cmdInsrt.CommandType = CommandType.Text

		'определяем субсчет (в случае валютных операций "валютный рынок"
		'If ImpTbl <> conImpCurAct _
		'	And rsImpTblStruct.EOF Then
		'	SubAccnt = Mrkt
		'ElseIf ImpTbl <> conImpCurAct And
		'    And RTrim(rsImpTblStruct.Fields("ExpField").Value) <> conSignCurMarket _
		'	And RTrim(rsImpTblStruct.Fields("ExpField").Value) <> conCurMarket Then
		'	SubAccnt = MrktImpTbl
		'Else SubAccnt = conCurMarket
		'End If
		If (ImpTbl = conImpCurAct Or ExpFld = conSignCurMarket Or ExpFld = conCurMarket) Then
			SubAccnt = conCurMarket
		ElseIf ImpTbl = conImpBcsDrvtvMove Then
			SubAccnt = conDrvtMrkt
		Else
			SubAccnt = Mrkt
		End If
		'формируем команду на вставку

		strInsert = "INSERT INTO " & ImpTbl & " (Broker, Accnt, SubAccnt, DFrom"
		strValues = "VALUES (?, ?, ?, ?"
		If aSeparFieldsData Is Nothing Then
			j = 0
		Else
			For j = LBound(aSeparFieldsData) To UBound(aSeparFieldsData)
				strInsert = strInsert & ", " & aSeparFieldsData(j).Fld
				strValues = strValues & ", ?"
			Next
		End If
		For i = LBound(aFieldsNClmns) To UBound(aFieldsNClmns)
			strInsert = strInsert & ", " & aFieldsNClmns(i).Fld
			strValues = strValues & ", ?"
		Next i

		cmdInsrt.CommandText = strInsert & ") " & strValues & ")"
		cmdInsrt.Parameters.Refresh()
		cmdInsrt.Prepared = True
		cmdInsrt.Parameters.Item(0).Value = Brk
		cmdInsrt.Parameters.Item(1).Value = Accnt
		cmdInsrt.Parameters.Item(2).Value = SubAccnt
		cmdInsrt.Parameters.Item(3).Value = FDate

		If j <> 0 Then
			For j = LBound(aSeparFieldsData) To UBound(aSeparFieldsData)
				cmdInsrt.Parameters.Item(4 + j).Value = aSeparFieldsData(j).FldData
			Next
		End If

		With oXL.Workbooks(FName).Worksheets(1)
			For NData = NDataBeg To NDataEnd
				For i = 0 To UBound(aFieldsNClmns)
					If DtInClmn Then
						If Len(RTrim(CStr(.Cells(aFieldsNClmns(i).NClmn, NData).Value))) = 0 Then cmdInsrt.Parameters.Item(4 + j + i).Value = DBNull.Value Else cmdInsrt.Parameters.Item(4 + j + i).Value = .Cells(aFieldsNClmns(i).NClmn, NData)
					Else
						If Len(RTrim(CStr(.Cells(NData, aFieldsNClmns(i).NClmn).Value))) = 0 Then cmdInsrt.Parameters.Item(4 + j + i).Value = DBNull.Value Else cmdInsrt.Parameters.Item(4 + j + i).Value = .Cells(NData, aFieldsNClmns(i).NClmn)
					End If
				Next i
				Try
					cmdInsrt.Execute()
				Catch ex As Exception
					MessageBox.Show(ex.Message)
					CloseFile()
					Stop
				End Try
			Next NData
		End With
		cnInsrt.Close()
	End Sub


	Protected Sub ProcessImpTables()
		'обработка вставленных таблиц процедурами базы данных
		Dim j As Integer
		Dim CnSP As New SqlConnection(strCn)
		CnSP.FireInfoMessageEventOnUserErrors = True
		Dim CmdSP As SqlCommand = CnSP.CreateCommand
		'Dim Trnsct As SqlTransaction
		'Trnsct = CnSP.BeginTransaction(IsolationLevel.ReadUncommitted)
		AddHandler CnSP.InfoMessage, New SqlInfoMessageEventHandler(AddressOf OnInfoMessage)
		CmdSP.Connection = CnSP
		CmdSP.CommandType = CommandType.StoredProcedure

		CnSP.Open()
		For j = 0 To 10
			If aImpFunct(j).vImpFunct = True Then
				CmdSP.CommandText = aImpFunct(j).NImpFunct
				CmdSP.Transaction = CnSP.BeginTransaction(IsolationLevel.ReadUncommitted)
				Try
					CmdSP.ExecuteNonQuery()
					CmdSP.Transaction.Commit()
				Catch ex As Exception
					MessageBox.Show(ex.Message)
					Stop
					CmdSP.Transaction.Rollback()
				End Try
			End If
		Next j
		CmdSP.CommandText = "spBlnsBldr"
		Dim prmtSP As New SqlParameter()
		prmtSP.ParameterName = "@RETURN_VALUE"
		prmtSP.SqlDbType = SqlDbType.Int
		prmtSP.Direction = ParameterDirection.ReturnValue
		CmdSP.Parameters.Add(prmtSP)

		prmtSP = New SqlParameter()
		prmtSP.ParameterName = "@Rsp"
		prmtSP.SqlDbType = SqlDbType.NChar
		prmtSP.Size = 50
		prmtSP.Direction = ParameterDirection.Output
		CmdSP.Parameters.Add(prmtSP)

		CmdSP.Transaction = CnSP.BeginTransaction(IsolationLevel.Serializable)
		Try
			CmdSP.ExecuteNonQuery()
		Catch ex As Exception
			MessageBox.Show(ex.Message)
			Stop
			CmdSP.Transaction.Rollback()
		End Try
		If RTrim(CmdSP.Parameters("@Rsp").Value) = "Успешная загрузка" Then
			CmdSP.Transaction.Commit()
		Else
			MessageBox.Show(CmdSP.Parameters("@Rsp").Value)
			CmdSP.Transaction.Rollback()
			Stop
		End If
		CnSP.Close()
	End Sub

	Protected Sub MoveFile()
		'при отсутствии в родительском каталоге папки год и субпапки месяц, их создает и перемещает туда файл
		Dim DirYear As String
		Dim SubDirMonth As String

		FileIO.FileSystem.CurrentDirectory = FileIO.FileSystem.GetParentPath(DirName)
		DirYear = FileIO.FileSystem.CurrentDirectory + "\" + CStr(Year(FDate)) + "\"
		SubDirMonth = DirYear + MonthName(Month(FDate)) + "\"
		If Not FileIO.FileSystem.DirectoryExists(DirYear) Then FileIO.FileSystem.CreateDirectory(DirYear)
		If Not FileIO.FileSystem.DirectoryExists(SubDirMonth) Then FileIO.FileSystem.CreateDirectory(SubDirMonth)
		FileIO.FileSystem.MoveFile(DirName & FName, SubDirMonth & FName)
	End Sub

	Private Shared Sub OnInfoMessage(sender As Object,
	args As SqlInfoMessageEventArgs)
		Dim err As SqlError
		For Each err In args.Errors
			Console.WriteLine("The {0} has received a severity {1}, _  
       state {2} error number {3}\n" &
				"on line {4} of procedure {5} on server {6}:\n{7}",
				err.Source, err.Class, err.State, err.Number, err.LineNumber,
			err.Procedure, err.Server, err.Message)
		Next
		Stop
	End Sub
End Class

'Public Class ParentXlLoader
'	Protected DirName As String
'	Protected strCn As String
'	Protected strCnUser As String
'	Protected Structure ImpFunct
'		Dim NImpFunct As String
'		Dim vImpFunct As Boolean
'	End Structure

'	Protected aImpFunct(0 To 10) As ImpFunct

'	Public Property ConnectionString() As String
'		Protected Get
'			Return ManagerFP.My.Settings.CnStrSQLClient
'		End Get
'		Set(ConStringValue As String)
'			strCnUser = ConStringValue
'			If strCnUser IsNot Nothing And strCnUser <> "" Then
'				strCn = strCnUser
'			Else
'				strCnUser = ManagerFP.My.Settings.CnStrSQLClient
'				strCn = ManagerFP.My.Settings.CnStrSQLClient + "; User ID=Rinat; Password = 'RInnat'"
'			End If
'		End Set
'	End Property

'	Friend Rsp As String 'строка сообщений
'	Protected oXL As Excel.Application = Nothing
'	Protected Structure FieldsData
'		Dim Fld As String
'		Dim FldData As String
'	End Structure

'	Public Sub New(ByVal ConnectionString As String, ByVal PathToFile As String)

'		oXL = New Excel.Application

'		DirName = PathToFile
'		strCnUser = ConnectionString
'		If strCnUser <> "" Then
'			strCn = strCnUser
'		Else
'			strCnUser = ManagerFP.My.Settings.CnStrSQLClient
'			strCn = ManagerFP.My.Settings.CnStrSQLClient + "; User ID=Rinat; Password = 'RInnat'"
'		End If

'		'контролер для запуска процедур базы данных
'		'aImpFunct(0).NImpFunct = "spImpPrices"
'		'aImpFunct(0).vImpFunct = True
'		'aImpFunct(1).NImpFunct = "spImpOtherStockAct"
'		'aImpFunct(1).vImpFunct = True
'		'aImpFunct(2).NImpFunct = "spImpOtherDerivAct"
'		'aImpFunct(2).vImpFunct = True
'		'aImpFunct(3).NImpFunct = "spImpCurAct"
'		'aImpFunct(3).vImpFunct = True
'		'aImpFunct(4).NImpFunct = "spImpStockAct"
'		'aImpFunct(4).vImpFunct = True
'		'aImpFunct(5).NImpFunct = "spImpRedemption"
'		'aImpFunct(5).vImpFunct = True
'		'aImpFunct(6).NImpFunct = "spImpDvdCpnPartP"
'		'aImpFunct(6).vImpFunct = True
'		'aImpFunct(7).NImpFunct = "spImpFutureAct"
'		'aImpFunct(7).vImpFunct = True
'		'aImpFunct(8).NImpFunct = "spImpOptionAct"
'		'aImpFunct(8).vImpFunct = False
'		'aImpFunct(9).NImpFunct = "spImpFutureRecalc"
'		'aImpFunct(9).vImpFunct = True
'		'aImpFunct(10).NImpFunct = "spImpOtherSecurAct"
'		'aImpFunct(10).vImpFunct = True

'	End Sub

'	Public Sub Load()

'	End Sub

'	Protected Sub OrdererFilesReport()

'	End Sub

'End Class

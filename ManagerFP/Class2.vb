Imports Microsoft.Office.Interop
Imports System.Data.Common
Imports System.Data.SqlClient
'Imports System.IO
''' <summary>
''' Класс предназначен для загрузки данных из файлов exel в базы данных Microsoft Server. 
''' 
''' <para> База данных для загрузки данных определяется строкой соединения в формате SQLConnection,
''' которая является параметром для конструктора класса
''' </para>
''' 
''' <para> exelфайлы для загрузки должны быть расположены в одной директории 
''' Путь к этой директории является параметром конструктора класса
''' Директория может содержать так же и файлы, не предназначенные для загрузки
''' В этом случае база данных должна содержать таблицу ImpFiles
''' <list type="Table">
''' <term> Id </term>
''' <description> Идентификатор строки, целое </description>
''' <term> FilesNamePattern </term>
''' <description> шаблоны имен файлов в формате шаблонов VB </description>
''' <term> field </term>
''' <description> имя поля, связанное с файлом, для внесения во все таблицы. Если Null, то таких данных нет </description>
''' <term> fielddata </term>
''' <description> данные для загрузки в имя поля, связанное с файлом. Если Null, то данные содержатся в имени файла </description>
''' <term> CharNumber </term> 
''' <description> номер символа в названии файла, откуда начинаются данные для загрузки </description>
''' <term> NumberOfChars </term>
''' <description> число символов для загрузки в названии файла </description>
''' </list>
''' </para>
''' 
''' <para> Таблицы для загрузки должны быть размещены на первом листе Exel
''' Одна таблица exelфайла загружается в одну таблицу базы данных
''' Одна таблица exelфайла не может грузиться в несколько таблиц базы данных
''' несколько таблиц exelфала могут грузиться в одну таблицу базы данных
''' Таблица базы данных может содержать, а может и не содержать первичные ключи
''' </para>
''' 
''' <para> В базе данных должна быть таблица ImpTblNames, связывающая названия таблиц источников в exel файлах и таблиц получателей в базах данных и 
'''		содержащая характеристики таблицы источника
'''		<List type = "Table">
'''			<term> Id </term>
'''			<description> Идентификатор строки, целое </description>
'''			<term> ExpTable </term>
'''			<description> Название таблицы источника в exelфайле. Используется для поиска местоположения таблицы 
'''				Если у таблицы нет названия, необходимо указать имя поля, уникальное для этой таблицы, не встречающееся в других местах exelфайла </description>
'''			<term> FilesNamePattern </term>
'''			<description> Шаблон имени файла, в котором содержится таблица </description>
'''			<term> ImpTable </term>
'''			<description> Название таблицы, получателя данных </description>
'''			<term> Always </term>
'''			<description> Бинарное, всегда ли должна быть таблица в файле. 
'''			В случае необнаружения таблицы в файле с always = True, выдается сообщение об отсутствии таблицы </description>
'''			<term> Field </term>
'''			<description> Имя поля для загрузки в ImpTable, связанное с таблицей источником </description>
'''			<term> FieldData </term>
'''			<description> значение для заполнения поля, связанного с таблицей источником </description>
'''			<term> TblBeg </term>
'''			<description> Отступ от клетки с названием таблицы до начала таблицы </description>
'''			<term> DateInColumn</term>
'''			<description> данные таблицы в столбцах(транспонированная таблица)</description>
'''			<term> DataBeg </term>
'''			<description> Отступ до начала блока данных в таблице от ее начала (в строках, если таблица обычная и в столбцах, если таблица транспонированная) </description>
'''			<term> DataEnd </term>
'''			<description> Окончание блока данных. Отступ от последней строки для обычной/столбца для транспонированной таблицы </description>
'''			<term> Brdr </term>
'''			<description> Состоит ли таблица из блоков и что разграничивает блоки таблицы 
'''			Значение Non означает цельную таблицу, Name  означает имя поля, которое содержит разные значения для разных блоков, 
'''			Separator означает, что блоки разделены разделителем </description>
'''			<term> SeparBeg </term>
'''			<description> Значение разделителя, обозначающего начало блока таблицы (может быть частью значения клетки таблицы) </description>
'''			<term> SeparEnd </term>
'''			<description> Значение разделителя, обозначающего конец блока таблицы (может быть частью значения клетки таблицы) </description>
'''			<term> Skip </term>
'''			<description> Имеется ли разрыв между блоками таблицы </description>
'''			<term> DFrom </term>
'''			<description> техническое поле, означает начало периода, для которого верны характеристики таблицы. DFrom входит в период </description>
'''			<term> DTo </term>
'''			<description> техническое поле, означает конец периода, для которого верны характеристики таблицы. DTo не входит в период </description>
'''		</List>
''' </para>
''' <para> В базе данных должна быть таблица ImpTblStrct, связывающая поля таблицы источника с полями таблицы для загрузки
'''		<List type = "Table">
'''			<term> ExpTable </term>
'''			<description> Имя таблицы источника </description> 
'''			<term> ExpField </term>
'''			<description> Имя поля таблицы источника </description>
'''			<term> Always </term>
'''			<description> Бинарное, всегда ли должно быть поле в таблице источнике. 
'''			В случае необнаружения поля с always = True, выдается сообщение об отсутствии поля в таблице </description>
'''			<term> Number </term>
'''			<description> Номер столбца с таким именем, в случае, если таблица содержит несколько столбцов с одинаковыми именами </description>
'''			<term> Content </term>
'''			<description> содержание поля таблицы источника 
'''			значение Field означает обычное поле, name означает значение, отличающее блоки таблицы </description>
'''			<term> FieldData </term>
'''			<description> для поля имени содержит данные для внесения в принимающую таблицу </description>
'''			<term> ImpTable </term>
'''			<description> Имя таблицы получателя данных </description>
'''			<term> ImpField </term>
'''			<description> имя поля получателя </description>
'''			<term> Repeat </term> 
'''			<description> следует ли повторить поиск поля с таким же именем </description>
'''			<term> DFrom </term>
'''			<description> техническое поле, означает начало периода, для которого верны характеристики связи полей. DFrom входит в период </description>
'''			<term> DTo </term>
'''			<description> техническое поле, означает конец периода, для которого верны характеристики связи полей. DTo не входит в период </description>
''' 	</List>
''' </para>
''' <para> Класс имеет один открытый метод Load
'''		Краткое описание алгоритма метода: 
'''		Сначала из директории файлов-источников выделяются файлы, имена которых соответствуют шаблонам, содержащимся в ImpFile.
'''		В таблицу связей полей таблиц источников и получателей заносятся поля и их значения, связанные с именами файлов.
'''		Файлы-источники упорядочиваются по датам, содержащихся в именах файлов.
'''		Файлы последовательно обрабатываются:
'''			Файл открывается, далее из таблицы ImpTableName определяются таблицы, которые в нем содержаться.
'''			В разрезе каждой таблицы производятся следующие действия
'''				Поиск таблицы по ее названию или ключевому слову
'''				В таблицу связей полей таблиц источников и получателей заносятся поля и их значения, связанные с именами таблиц.
'''				Находится тело таблицы 
'''				Делается запрос к таблице ImpTblStrct для получения связей имен полей таблиц источников и таблиц получателей
'''				Каждая запись, содержащая имя поля обрабатывается следующим образом:
'''					Находится ее местоположение в таблице источнике, номер столца/строки заносится в таблицу связей полей таблиц источников и таблиц получателей
'''				Если таблица разделена на блоки разделителями, запукается цикл поиска разделителей
'''					находится разделитель начала блока.
'''					В таблицу связей полей и таблиц источников и таблиц получателей заносятся поля и данные, содержащиеся в строке-разделителе.
'''					Находится разделитель конца блока.
'''					Вносятся данные из строк между разделителями
'''					Из таблицы связей полей и таблиц источников и таблиц получателей удаляются строки, связанные с разделителем.
'''					повторение цикла
'''				Если таблица разделена на блоки с помощью поименованных строк, запускается цикл обработки имен поименованных строк 
'''					Находится поименованная строка с таким именем 
'''					В таблицу связей полей и таблиц источников и таблиц получателей заносятся поля и данные, связанные с именем строки.
'''					Вносятся данные из поименованной строки
'''					повторение цикла
'''					Из таблицы связей полей и таблиц источников и таблиц получателей удаляются строки, связанные с именем строки.
'''					переход к следующему имени
'''				Если таблица не разделена на блоки, находится последняя строка в данных
'''					Вносятся данные с первой по последнюю строку в таблицу получатель в оперативной памяти
'''				Далее из таблицы связей полей и таблиц источников и таблиц получателей удаляются строки, связанные с именем таблицы.
'''				Данные из таблицы получателя в оперативной памяти передаются в таблицу базы данных
'''			Переход к следующей таблице
'''	    Запускаются встроенные процедуры для обработки поступивших данных. 	
'''		Переход к следующему файлу.
'''		</para>
''' 
'''
''' </summary>
Public Class ParentXlLoader


	Protected PthDir As String 'имя директории с загружаемыми файлами
	'Protected FName As String 'имя загружаемого файла
	'Protected ShName As String  'имя листа из загружаемого файла
	'Protected RcInClmn As Boolean 'записи данных в столбцах

	Protected oXL As Excel.Application = Nothing
	Protected rngTblDt As Excel.Range

	Protected strCn As String 'строка подключения
	Protected strCnUser As String
	Protected cnForLoad As SqlConnection
	Protected dsTblsPrmsForLoad As DataSet
	Protected adpImpFiles As SqlDataAdapter
	Protected adpImpTblNames As SqlDataAdapter
	Protected adpImpTblStrct As SqlDataAdapter
	Protected adpImpTbl


	Protected Structure ImpFunct
		Dim NImpFunct As String
		Dim vImpFunct As Boolean
	End Structure

	Protected aImpFunct(0 To 12) As ImpFunct 'список функций для вызова
	'Protected dsTblsPrmsForLoad As DataSet
	'Protected adpImpFiles As SqlDataAdapter
	'Protected adpImpTblNames As SqlDataAdapter
	'Protected adpImpTblStrct As SqlDataAdapter

	Protected Structure OrdererFile
		'Dim NumberFile As Integer
		Dim FName As String
		Dim FieldForOrder As DateTime
		Dim FilesNamePtrn As String
	End Structure

	Protected aOrdererFiles() As OrdererFile

	Protected tblFieldsData As DataTable
	Protected clmnFD As DataColumn
	Protected rowFD As DataRow

	Protected Structure FieldsData
		Dim FName As String
		Dim Fld As String
		Dim DataType As String
		Dim FldData As String
		Dim FldNumber As Integer
	End Structure

	Protected aFieldsData() As FieldsData

	Friend Rsp As String 'строка сообщений

	Public Property PathToDir() As String
		Get
			Return PthDir
		End Get
		Set(PathToDir As String)
			If Right(PathToDir, 1) <> "\" Then PthDir = PathToDir + "\" Else PthDir = PathToDir
		End Set
	End Property

	'Public Property FileName() As String
	'	Get
	'		Return FName
	'	End Get
	'	Set(FileName As String)
	'		FName = FileName
	'	End Set
	'End Property

	'Public Property SheetName() As String
	'	Get
	'		Return ShName
	'	End Get
	'	Set(SheetName As String)
	'		ShName = SheetName
	'	End Set
	'End Property

	'Public Property rngTablData() As Excel.Range
	'	Get
	'		Return rngTblDt
	'	End Get
	'	Set(rngTablData As Excel.Range)
	'		rngTblDt = rngTablData
	'	End Set
	'End Property

	'Public Property RecordInColumn() As Boolean
	'	Get
	'		Return RcInClmn
	'	End Get
	'	Set(RecordInColumn As Boolean)
	'		RcInClmn = RecordInColumn
	'	End Set
	'End Property

	Public Property ConnectionString() As String
		Protected Get
			Return ManagerFP.My.Settings.CnStrSQLClient
		End Get
		Set(ConStringValue As String)
			strCnUser = ConStringValue
			If strCnUser IsNot Nothing And strCnUser <> "" Then
				strCn = strCnUser
			Else
				strCn = ManagerFP.My.Settings.CnStrSQLClient
			End If
		End Set
	End Property

	'''<param name="ConnectionString"> строка соединения в формате SQLClient с базой , в которую загружается информация. 
	'''</param>
	'''<param name="PathToDir"> путь к директории, в которой содержатся exelфайлы-источники  
	'''</param>'''
	Public Sub New(ByVal ConnectionString As String, ByVal PathToDir As String) ', FileName As String, SheetName As String, rngTablData As Excel.Range,
		'ByVal RecordInColumn As Boolean, dsTblsPrmsForLoad As DataSet)

		oXL = New Excel.Application

		If Right(PathToDir, 1) <> "\" Then PthDir = PathToDir + "\" Else PthDir = PathToDir
		'FName = FileName
		'If Len(RTrim(SheetName)) = 0 Then ShName = "Лист1" Else ShName = SheetName

		'rngTblDt = rngTablData

		strCnUser = ConnectionString
		If strCnUser <> "" Then
			strCn = strCnUser
		Else
			strCnUser = ManagerFP.My.Settings.CnStrSQLClient
			strCn = ManagerFP.My.Settings.CnStrSQLClient + "; User ID=Rinat; Password = 'RInnat!1987'"
		End If

		'Открываем подключение и заполняем датасет с параметрами загрузки
		cnForLoad = New SqlConnection(strCn) 'соединение с БД чтение 
		'cnForLoad.Open()

		dsTblsPrmsForLoad = New DataSet("PrmsForLoads")
		'таблица с шаблонами загружаемых файлов и полями данных, содержащихся в имени файла
		adpImpFiles = New SqlDataAdapter("SELECT * From dbo.ImpFiles", cnForLoad)
		'adpImpFiles.MissingMappingAction = MissingMappingAction.Error
		adpImpFiles.Fill(dsTblsPrmsForLoad, "ImpFiles")
		'таблица с параметрами таблиц источников для загрузки
		adpImpTblNames = New SqlDataAdapter("SELECT * FROM ImpTblNames", cnForLoad)
		adpImpTblNames.Fill(dsTblsPrmsForLoad, "ImpTblNames")
		'таблица соотвествий полей таблиц источников и таблиц назначений
		adpImpTblStrct = New SqlDataAdapter("SELECT * FROM ImpTblStrct", cnForLoad)
		adpImpTblStrct.Fill(dsTblsPrmsForLoad, "ImpTblStrct")

		'создаем таблицу для записи значений полей или соответствия полей 
		tblFieldsData = New DataTable("tblFieldsData")

		clmnFD = New DataColumn With {
			.DataType = System.Type.GetType("System.String"),
			.ColumnName = "FName",
			.ReadOnly = False,
			.Unique = False}
		tblFieldsData.Columns.Add(clmnFD)

		clmnFD = New DataColumn With {
			.DataType = System.Type.GetType("System.String"),
			.ColumnName = "TblName",
			.ReadOnly = False,
			.Unique = False}
		tblFieldsData.Columns.Add(clmnFD)

		clmnFD = New DataColumn With {
			.DataType = System.Type.GetType("System.String"),
			.ColumnName = "WhereFrom",
			.ReadOnly = False,
			.Unique = False}
		tblFieldsData.Columns.Add(clmnFD)

		clmnFD = New DataColumn With {
			.DataType = System.Type.GetType("System.String"),
			.ColumnName = "Fld",
			.ReadOnly = False,
			.Unique = False}
		tblFieldsData.Columns.Add(clmnFD)

		clmnFD = New DataColumn With {
			.DataType = System.Type.GetType("System.String"),
			.ColumnName = "FldData",
			.ReadOnly = False,
			.Unique = False}
		tblFieldsData.Columns.Add(clmnFD)

		clmnFD = New DataColumn With {
			.DataType = System.Type.GetType("System.Int32"),
			.ColumnName = "FldNumber",
			.ReadOnly = False,
			.Unique = False}
		tblFieldsData.Columns.Add(clmnFD)

		dsTblsPrmsForLoad.Tables.Add(tblFieldsData)


		'контролер для запуска процедур базы данных
		aImpFunct(0).NImpFunct = "spImpPrices"
		aImpFunct(0).vImpFunct = True
		aImpFunct(1).NImpFunct = "spImpOtherStockAct"
		aImpFunct(1).vImpFunct = True
		aImpFunct(2).NImpFunct = "spImpOtherDerivAct"
		aImpFunct(2).vImpFunct = True
		aImpFunct(3).NImpFunct = "spImpCurAct"
		aImpFunct(3).vImpFunct = True
		aImpFunct(4).NImpFunct = "spImpBcsMoneyMove"
		aImpFunct(4).vImpFunct = True
		aImpFunct(5).NImpFunct = "spImpBcsDerivativeMove"
		aImpFunct(5).vImpFunct = True
		aImpFunct(6).NImpFunct = "spImpStockAct"
		aImpFunct(6).vImpFunct = True
		aImpFunct(7).NImpFunct = "spImpRedemption"
		aImpFunct(7).vImpFunct = True
		aImpFunct(8).NImpFunct = "spImpDvdCpnPartP"
		aImpFunct(8).vImpFunct = True
		aImpFunct(9).NImpFunct = "spImpFutureAct"
		aImpFunct(9).vImpFunct = True
		aImpFunct(10).NImpFunct = "spImpOptionAct"
		aImpFunct(10).vImpFunct = False
		aImpFunct(11).NImpFunct = "spImpFutureRecalc"
		aImpFunct(11).vImpFunct = True
		aImpFunct(12).NImpFunct = "spImpOtherSecurAct"
		aImpFunct(12).vImpFunct = True

	End Sub

	Protected Overrides Sub Finalize()
		If oXL IsNot Nothing Then
			oXL.Application.ScreenUpdating = True
			oXL.Application.DecimalSeparator = ","
			oXL.Application.UseSystemSeparators = True
			oXL.Quit()
			oXL = Nothing
			cnForLoad.Close()
			cnForLoad = Nothing
		End If

	End Sub

	Public Sub Load()
		Me.OrdererFilesReport()
		Dim cntFiles As Integer
		For cntFiles = LBound(aOrdererFiles) To UBound(aOrdererFiles)
			Me.OpenFile(aOrdererFiles(cntFiles).FName)
			Me.RecipientTblNamesStructure(aOrdererFiles(cntFiles))
			Me.CloseFile(aOrdererFiles(cntFiles).FName)
			Me.ProcessImpTables()
			Me.MoveFile(PthDir, aOrdererFiles(cntFiles).FName, aOrdererFiles(cntFiles).FieldForOrder)

		Next cntFiles


	End Sub
	''' <summary>
	''' <para> Выбирает из указанной директории все файлы, имена которых соответствуют одному из шаблонов имен файлов для загрузки,
	'''	содержащихся в таблице ImpFiles.
	'''	Далее упорядочивает по датам, содержащимся в имени файла, при этом если имя файла не содержит даты, такие файлы обрабатываются в 
	'''	первую очередь. Также в таблицу с именами полей и данных для загрузки заносятся атрибуты, содержащиеся в имени файла или ассоциирующиеся с шаблоном.  
	''' </para>
	''' </summary>
	Protected Sub OrdererFilesReport()
		Dim RecOrderer As Boolean 'принимает значение True если имя файла уже записано в ордерер 
		Dim dtr As DataRow
		Dim DirReports As IO.DirectoryInfo = New IO.DirectoryInfo(PthDir) 'объект директории 
		Dim FileReport As IO.FileInfo() = DirReports.GetFiles()
		Dim q As Integer = 0 'счетчик файлов, имена которых соотвествуют шаблонам загружаемых файлов  
		Dim qq As Integer 'вспомогательный счетчик файлов для упорядочивания

		For Each f As IO.FileInfo In FileReport
			'Проверяем файлы из дирректории для загрузки на соответствие одному из шаблонов имен файлов для загрузки
			'Если соответсвующий шаблон обнаружен, то заносим в массив для упорядоченья и в таблицу с именами полей и данных для загрузки, соотвествующих шаблону имени файла
			RecOrderer = False
			For Each dtr In dsTblsPrmsForLoad.Tables("ImpFiles").Rows
				If f.Name Like RTrim(dtr("FilesNamePattern")) Then
					If RecOrderer = 0 Then
						ReDim Preserve aOrdererFiles(0 To q)
						aOrdererFiles(q).FName = f.Name
						aOrdererFiles(q).FilesNamePtrn = dtr("FilesNamePattern")
						aOrdererFiles(q).FieldForOrder = #1900-01-01# 'приоритет загрузки файлов без дат в имени
						RecOrderer = True
						q = q + 1
					End If
					If Not IsDBNull(dtr("Field")) Then
						' Заполняем таблицу с именами полей и данных для загрузки  
						rowFD = tblFieldsData.NewRow
						rowFD("FName") = f.Name
						rowFD("TblName") = "Any"
						rowFD("WhereFrom") = "File"
						rowFD("Fld") = dtr("Field").ToString.Trim
						If Not IsDBNull(dtr("FieldData")) Then
							rowFD("FldData") = dtr("FieldData").ToString.Trim
						Else
							rowFD("FldData") = Mid(f.Name, dtr("CharNumber"), dtr("NumberOfChars"))
						End If
						'Если имя поля - дата, то дополняем массив для упорядочения значением даты
						If RTrim(dtr("Field")) = "DFrom" Then
							If Not IsDBNull(dtr("FieldData")) Then
								aOrdererFiles(q - 1).FieldForOrder = CDate(dtr("FieldDate"))
							Else
								aOrdererFiles(q - 1).FieldForOrder = CDate(Mid(f.Name, dtr("CharNumber"), dtr("NumberOfChars")))
							End If
							' Если имя поля - год и месяц, то в массиве для упорядочения изменяем значение поля на дату последнего дня указанного месяца 

						ElseIf RTrim(dtr("Field")) = "YearMonth" Then
							If Not IsDBNull(dtr("FieldData")) Then
								aOrdererFiles(q - 1).FieldForOrder = DateAdd("d", -1, DateAdd("m", 1, CDate("20" + dtr("FieldDate") + "-01")))
							Else
								aOrdererFiles(q - 1).FieldForOrder = DateAdd("d", -1, DateAdd("m", 1, CDate("20" + Mid(f.Name, dtr("CharNumber"), dtr("NumberOfChars")) + "-01")))
							End If
							' в таблице с именами полей и данных имя поля изменяем на DFrom, значение последним днем месяца
							rowFD("fld") = "DFrom"
							rowFD("fldData") = aOrdererFiles(q - 1).FieldForOrder
						End If
						'добавляем строку в таблицу с именами полей и данных для загрузки
						tblFieldsData.Rows.Add(rowFD)
					End If
				End If
			Next
		Next
		'упорядочиваем файлы по дате создания
		Dim Fqq As OrdererFile
		For qq = 0 To q - 2
			If aOrdererFiles(qq).FieldForOrder > aOrdererFiles(qq + 1).FieldForOrder Then
				Fqq = aOrdererFiles(qq)
				aOrdererFiles(qq) = aOrdererFiles(qq + 1)
				aOrdererFiles(qq + 1) = Fqq
				If qq > 0 Then qq = qq - 2
			End If
		Next qq

		'For qq = 0 To q - 1
		'	Me.
		'Next qq






		''подготовка к выборке шаблонов файлов для загрузки
		''готовим подключение
		'Dim aFlPttrns() As String
		'Dim cnReadFlPttrns As New SqlConnection(strCn) 'соединение с БД чтение 
		'готовим объект команды
		'Dim cmdReadFlPttrns As New SqlCommand()
		'cmdReadFlPttrns.Connection = cnReadFlPttrns
		'cmdReadFlPttrns.CommandType = CommandType.StoredProcedure
		'cmdReadFlPttrns.CommandText = "spImpFlPttrns"
		'готовим параметры команды
		'Dim prmtReadFlPttrns As New SqlParameter()
		'prmtReadFlPttrns.ParameterName = "@Date"
		'prmtReadFlPttrns.SqlDbType = SqlDbType.DateTime
		'prmtReadFlPttrns.Direction = ParameterDirection.Input
		'prmtReadFlPttrns.Value = Date.Today
		'cmdReadFlPttrns.Parameters.Add(prmtReadFlPttrns)
		'читаем шаблоны
		'cnReadFlPttrns.Open()
		'Dim rdFlPttrns As SqlDataReader = cmdReadFlPttrns.ExecuteReader()
		'While (rdFlPttrns.Read())
		'	ReDim Preserve aFlPttrns(0 To i)
		'	aFlPttrns(i) = rdFlPttrns("FilesNamePattern").ToString.Trim
		'	i = i + 1
		'End While
		'Dim dtrdFlPttrns As DataTableReader = dsTblsPrmsForLoad.Tables("ImpFiles").CreateDataReader



	End Sub
	Public Sub OpenFile(FName As String)
		With oXL
			If Right(FName, Len(FName) - InStrRev(FName, ".")) = "html" Then
				.Application.DecimalSeparator = "."
				.Application.UseSystemSeparators = False
			End If
			.Workbooks.Open(PthDir & FName)
			.Application.ReferenceStyle = Excel.XlReferenceStyle.xlR1C1
		End With
	End Sub

	Public Sub CloseFile(FName As String)
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

	''' <summary>
	''' <para> По шаблону имени файла выбирает из таблицы ImpTblNames имена таблиц, содержащихся в таких файлах, а также метаданные таких таблиц.
	''' К метаданным относятся имя таблицы для загрузки, начало и конец данных, а также данные, непосредственно связанные с таблицей
	''' Далее для каждой таблицы:
	'''		Данные,непосредственно связанные с таблицей загружаются в таблицу с данными для загрузки
	'''		Вызывается метод ComplianserFieldClm и передает ему метаданные для заполнения массива соответствия полей таблицы получателя с номерами 
	'''		столбцов таблицы источника
	''' </para>
	''' </summary>
	''' <param name="ExmpOrdererFile"></param>
	Protected Sub RecipientTblNamesStructure(ByRef ExmpOrdererFile As OrdererFile)
		'выборка наименований таблиц назначений и таблиц источников и метаданных о таблицах источниках
		Dim dtr As DataRow
		Dim strSQL As String
		Dim ImpTbl As String
		Dim ExpTbl As String
		Dim Alws As Boolean 'таблица обязательно должна быть в файле или нет
		Dim Brdr As String 'таблица разбита на подтаблицы или нет. И если разбита - поименованными записями или разделителями 
		Dim DtInClmn As Boolean 'данные в строках ("no"), данные в столбцах ("yes")
		Dim TblBeg As Integer 'начало таблицы
		Dim DtBeg As Integer 'начало блока данных таблицы
		Dim DtEnd As Integer 'конец блока данных таблицы
		Dim SeparBeg As String 'разделитель подтаблиц - начало подтаблицы
		Dim SeparEnd As String 'разделитель подтаблиц - окончание подтаблицы
		Dim Skip As Boolean 'есть ли разрыв между подтаблицами

		strSQL = "[FilesNamePattern] ='" + ExmpOrdererFile.FilesNamePtrn + "' And [DFrom] <= CONVERT('" + ExmpOrdererFile.FieldForOrder + "', 'System.DateTime') and CONVERT('" + ExmpOrdererFile.FieldForOrder + "', 'System.DateTime') < [DTo]"
		'strSQL = "[DFrom] <= CONVERT('" + CStr(ExmpOrdererFile.FieldForOrder) + "', 'System.DateTime')"    'N'" + ExmpOrdererFile.FieldForOrder + "'"
		For Each dtr In dsTblsPrmsForLoad.Tables("ImpTblNames").Select(strSQL)
			ImpTbl = dtr("ImpTable").ToString.Trim
			ExpTbl = dtr("ExpTable").ToString.Trim
			'Заполняем таблицу с данными для заггрузки данные, связанные с именем таблицы - источника
			If Not IsDBNull(dtr("Field")) Then
				rowFD = tblFieldsData.NewRow
				rowFD("FName") = ExmpOrdererFile.FName
				rowFD("TblName") = ExpTbl
				rowFD("WhereFrom") = "Table"
				rowFD("Fld") = dtr("Field").ToString.Trim
				rowFD("FldData") = dtr("FieldData").ToString.Trim
				tblFieldsData.Rows.Add(rowFD)
			End If
			Alws = dtr("Always")
			Brdr = dtr("Brdr").ToString.Trim
			DtInClmn = dtr("DateInColumn")
			TblBeg = dtr("TblBeg")
			DtBeg = dtr("DataBeg")
			If IsDBNull(dtr("DataEnd")) Then DtEnd = 0 Else DtEnd = dtr("DataEnd")
			If IsDBNull(dtr("SeparBeg")) Then SeparBeg = "" Else SeparBeg = dtr("SeparBeg").ToString.Trim
			If IsDBNull(dtr("SeparEnd")) Then SeparEnd = "" Else SeparEnd = dtr("SeparBeg").ToString.Trim
			Skip = dtr("Skip")


			'заполняет массив с названиями полей и номерами столбцов/строк в таблице источнике
			Me.ComplianserFieldClmn(ExmpOrdererFile, ImpTbl, ExpTbl, TblBeg, Alws, Brdr, DtInClmn, SeparBeg, SeparEnd, DtBeg, DtEnd, Skip)

		Next
	End Sub

	''' <summary>
	''' <para> Запрашивает из таблицы ImpTblStrct базы данных соотвествие полей таблицы-адресата и таблицы источника, 
	''' заполняет в таблице данных наименования полей и номеров столбцов из таблиц источников
	''' Далее вызывает процедуру-вставки данных в следующей последовательности:
	''' Строки данных с границами, поименованные строки данных,  обычные строки данных.
	''' Из границ строк данных считываем содержащиеся в них поля в массив соответствия 
	''' </para>
	''' </summary>
	''' <param name="ExmpOrdererFile"> структура, содержащая имя файла, поле для упорядочивания и шаблон имени файла </param>
	''' <param name="ExpTbl"> имя таблицы источника данных </param>
	''' <param name="TblBeg"> Отступ от клетки с названием таблицы до начала таблицы </param>
	''' <param name="Always"> Бинарное, всегда ли должна быть таблица в файле </param>
	''' <param name="Brdr"> Состоит ли таблица из блоков и что разграничивает блоки таблицы 
	'''			Значение Non означает цельную таблицу, Name  означает имя поля, которое содержит разные значения для разных блоков, 
	'''			Separator означает, что блоки разделены разделителем </param>
	''' <param name="DtInClmn"> Бинарное, записи в столбиках или строках </param>
	''' <param name="SeparBeg"> Строка для поиска разделителя таблицы - начала блока таблицы </param>
	''' <param name="SeparEnd"> Строка для поиска разделителя таблицы - окончания блока </param>
	''' <param name="DtBeg"> Отступ от начала таблицы до первой записи данных </param>
	''' <param name="DtEnd"> Отступ от конца таблицы до последней записи </param>
	''' <param name="Skip"> Имеется ли разрыв между блоками таблицы </param>
	Protected Sub ComplianserFieldClmn(ByRef ExmpOrdererFile As OrdererFile, ByRef ImpTbl As String, ByRef ExpTbl As String, ByVal TblBeg As Integer, ByRef Always As Boolean _
																		 , ByRef Brdr As String, ByRef DtInClmn As Boolean, ByRef SeparBeg As String, ByRef SeparEnd As String, ByVal DtBeg As Integer _
																		 , ByVal DtEnd As Integer, ByRef Skip As Boolean)

		Dim strSQL As String
		Dim Dtr As DataRow 'запись

		Dim rngNameTbl As Excel.Range 'ячейка с именем таблицы 
		Dim rngTbl As Excel.Range 'таблица
		Dim rngLastCell As Excel.Range 'последняя ячейка таблицы
		'Dim aFieldsNClmns() As FieldsNClmn 'массив с названиями полей и соответствующим им номерам столбцов/строк в таблице источнике
		Dim Response As Long
		Dim NBeg As Integer
		Dim NEnd As Integer
		'Dim i As Integer 'счетчик полей 
		'Dim k As Integer 'счетчик полей в разделителе
		Dim r As Excel.Range
		Dim rngRow As Excel.Range 'строка данных
		'заполняет таблицу с названиями полей и номерами столбцов/строк в таблице источнике
		With oXL.Workbooks(ExmpOrdererFile.FName).Worksheets(1)
			' находим ячейку с названием таблицы источника
			' проверка наличия таблицы источника и ее обязательности
			rngNameTbl = .Cells.Find(ExpTbl, , , LookAt:=Excel.XlLookAt.xlWhole)
			If Always And rngNameTbl Is Nothing Then
				Response = MsgBox("Изменилось название таблицы " & ExpTbl & " в " & ExmpOrdererFile.FName & ", дополните таблицу ImpTblNames новым названием таблицы, потом догрузите файл", vbOK)
				Stop
			End If

			If Not rngNameTbl Is Nothing Then
				'определили область таблицы источника как окружение вокруг ячейки на заданное параметром TblBeg деления ниже названия
				rngTbl = .Cells(rngNameTbl.Row + TblBeg, rngNameTbl.Column).CurrentRegion

				'i = UBound(aFieldsData) + 1
				'читаем таблицу соотвествий. 
				'С начала перечень полей 
				strSQL = "ImpTable = '" + ImpTbl + "' and ExpTable = '" + ExpTbl + "' and Content = 'Field' and DFrom <= CONVERT('" + ExmpOrdererFile.FieldForOrder + "', 'System.DateTime') and CONVERT('" + ExmpOrdererFile.FieldForOrder + "', 'System.DateTime') < [DTo]"
				For Each Dtr In dsTblsPrmsForLoad.Tables("ImpTblStrct").Select(strSQL)
					'Заполняем таблицу с полями из структуры таблиц и номерами строк для транспонированных таблиц/номерами столбцов для обычных в таблицах-источниках
					rowFD = tblFieldsData.NewRow
					rowFD("FName") = ExmpOrdererFile.FName
					rowFD("TblName") = ExpTbl
					rowFD("WhereFrom") = "Table"
					rowFD("Fld") = Dtr("Impfield").ToString.Trim
					If Not IsDBNull(Dtr("FieldData")) Then
						rowFD("FldData") = Dtr("FieldData").ToString.Trim
					Else
						'ReDim Preserve aFieldsData(0 To i)
						'aFieldsData(i).FName = ExmpOrdererFile.FName
						'aFieldsData(i).Fld = Dtr("Impfield").ToString.Trim
						'находим ячейку с именем поля таблицы источника, если таких несколько, то соответствующее по номеру 
						r = rngTbl.Find(What:=Dtr("ExpField").ToString.Trim, After:=rngTbl(1, 1), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlPart, SearchDirection:=Excel.XlSearchDirection.xlNext)
						If r Is Nothing Then r = rngTbl.Find(What:=Dtr("ExpField").ToString.Trim, After:=rngTbl(rngTbl.Rows.Count, rngTbl.Columns.Count), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlPart, SearchDirection:=Excel.XlSearchDirection.xlPrevious)
						If Dtr("Number") > 1 Then
							For iNum = 2 To Dtr("Number")
								r = rngTbl.FindNext(r)
							Next iNum
						End If
						'В зависимости от того, явялется таблица обычной(данные в строках) или транспонированной(данные в столбцах) заполняем номерами строк или столбцов
						If DtInClmn Then rowFD("FldNumber") = r.Row Else rowFD("FldNumber") = r.Column 'Then aFieldsData(i).FldNumber = r.Row Else aFieldsData(i).FldNumber = r.Column
					End If
					tblFieldsData.Rows.Add(rowFD)
				Next

				'Обработка таблиц, содержащих подтаблицы, выделенные разделителями
				If Brdr = "Separator" Then
					'обрабатываем разрыв в таблице
					If Skip = True Then
						rngLastCell = rngTbl(rngTbl.Rows.Count, 1)
						'Dim ProvercaI As Integer
						rngTbl = .Cells(rngLastCell.Row + 2, rngLastCell.Column).CurrentRegion
						rngTbl = .Range(rngTbl(0, 1), rngTbl(rngTbl.Rows.Count, 1))
					End If
					rngRow = rngTbl.Find(SeparBeg, After:=rngTbl(1, 1), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlPart)
					If rngRow Is Nothing Then
						'If rdTblStrct("Always") Then
						Response = MsgBox("Изменился разделитель строк данных " & SeparBeg & " таблицы " & ExpTbl & " в " & ExmpOrdererFile.FName & ", дополните таблицу ImpTblNames новым разделителем, потом догрузите файл", vbOK)
						Stop
						'End If
					Else
						'осуществляем выборку полей из разделителя
						'i = 0
						'Dim aSeparFieldsInClmn() As FieldsData
						'заполняем таблицу с полями для загрузки данными из разделителя
						strSQL = "ImpTable = '" + ImpTbl + "' and ExpTable = '" + ExpTbl + "' and Content = 'FieldInSeparator' and [DFrom] <= CONVERT('" + ExmpOrdererFile.FieldForOrder + "', 'System.DateTime') and CONVERT('" + ExmpOrdererFile.FieldForOrder + "', 'System.DateTime') < [DTo]"
						Do
							For Each Dtr In dsTblsPrmsForLoad.Tables("ImpTblStrct").Select(strSQL)
								rowFD = tblFieldsData.NewRow
								rowFD("FName") = ExmpOrdererFile.FName
								rowFD("TblName") = ExpTbl
								rowFD("WhereFrom") = "Separ"
								rowFD("Fld") = Dtr("ImpField").ToString.Trim
								rowFD("FldData") = CStr(.Cells(rngRow.Row, Dtr("Number")).Value)
								tblFieldsData.Rows.Add(rowFD)
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
								Call InsertArroyInTbl(ExmpOrdererFile, ImpTbl, ExpTbl, DtInClmn, NBeg, NEnd)
								RowsDeleter(tblFieldsData.Select("WhereFrom = 'Separ'"))
								Exit Do
							Else
								NEnd = rngRow.Row - 1 + DtEnd
								Call InsertArroyInTbl(ExmpOrdererFile, ImpTbl, ExpTbl, DtInClmn, NBeg, NEnd)

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
							RowsDeleter(tblFieldsData.Select("WhereFrom = 'Separ'"))
						Loop Until rngRow Is Nothing
					End If

				ElseIf Brdr = "Name" Then
					'Вставляем записи с поименованными данными
					strSQL = "ImpTable = '" + ImpTbl + "' and ExpTable = '" + ExpTbl + "' and Content = 'Name' and  [DFrom] <= CONVERT('" + ExmpOrdererFile.FieldForOrder + "', 'System.DateTime') and CONVERT('" + ExmpOrdererFile.FieldForOrder + "', 'System.DateTime') < [DTo]"
					For Each Dtr In dsTblsPrmsForLoad.Tables("ImpTblStrct").Select(strSQL)
						rngRow = rngTbl.Find(Dtr("ExpField").ToString.Trim, After:=rngTbl(1, 1), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
						If rngRow Is Nothing Then
							If Dtr("Always") Then
								Response = MsgBox("Изменилось название поля " & Dtr("ExpField").ToString.Trim & " таблицы " & ExpTbl & " в " & ExmpOrdererFile.FName & ", дополните таблицу ImpTblNames новым названием поля, потом догрузите файл", vbOK)
								Stop
							End If
						Else
							'Если поименованные данные содержат поле для загрузки, Вносим его и данные в таблицу с именами полей и данных  
							If Not IsDBNull(Dtr("ImpField")) Then
								rowFD = tblFieldsData.NewRow
								rowFD("FName") = ExmpOrdererFile.FName
								rowFD("TblName") = ExpTbl
								rowFD("WhereFrom") = "NamedField"
								rowFD("Fld") = Dtr("ImpField").ToString.Trim
								rowFD("FldData") = Dtr("FieldData")
								tblFieldsData.Rows.Add(rowFD)
							End If
							Do
								Do
									If DtInClmn Then NBeg = rngRow.Column Else NBeg = rngRow.Row
									NEnd = NBeg
									Call InsertArroyInTbl(ExmpOrdererFile, ImpTbl, ExpTbl, DtInClmn, NBeg, NEnd)
									rngRow = rngTbl.Find(Dtr("ExpField").ToString.Trim, rngRow, , LookAt:=Excel.XlLookAt.xlWhole)
								Loop While (Dtr("Repeat") And NBeg < IIf(DtInClmn, rngRow.Column, rngRow.Row))
								RowsDeleter(tblFieldsData.Select("WhereFrom = 'NamedField'"))
								If Skip = True Then
									rngLastCell = rngTbl(rngTbl.Rows.Count, 1)
									rngTbl = .Cells(rngLastCell.Row + 2, rngLastCell.Column).CurrentRegion
									rngRow = rngTbl.Find(Dtr("ExpField").ToString.Trim, After:=rngTbl(1, 1), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
								End If
							Loop While Not rngRow Is Nothing
							rngTbl = .Cells(rngNameTbl.Row + TblBeg, rngNameTbl.Column).CurrentRegion
						End If
					Next
				ElseIf Brdr = "Not" Then
					'вставляем записи с непоименованными строками данных
					If DtInClmn Then
						NBeg = rngNameTbl.Column + DtBeg
						NEnd = rngTbl.Columns.Count - 1 + DtEnd
					Else
						NBeg = rngNameTbl.Row + DtBeg 'определили номер строки с началом данных в таблице источнике
						NEnd = NBeg + rngTbl.Rows.Count - 1 + DtEnd
					End If
					Call InsertArroyInTbl(ExmpOrdererFile, ImpTbl, ExpTbl, DtInClmn, NBeg, NEnd)
				End If
			End If
		End With
	End Sub

	''' <summary>
	''' <para> 'вставляет данные в таблицу базы данных из таблицы данных и номеров столбцов 
	''' сначала формируем датасет, датаадаптер и временную таблицу для загрузки
	''' заносим данные во временную таблицу
	''' 
	''' </para>
	''' </summary>
	''' <param name="ExmpOrdererFile"> структура, содержащая имя файла, поле для упорядочивания и шаблон имени файла </param>
	''' <param name="ImpTbl"> имя таблицы получателя данных </param>
	''' <param name="ExpTbl"> имя таблицы источника данных </param>
	''' <param name="DtInClmn"> Бинарное, записи в столбиках или строках </param>
	''' <param name="NDataBeg"> первая запись данных </param> 
	''' <param name="NDataEnd"> последняя запись данных </param>

	Protected Sub InsertArroyInTbl(ExmpOrdererFile As OrdererFile, ImpTbl As String, ExpTbl As String, ByVal DtInClmn As Boolean, ByVal NDataBeg As Integer, ByVal NDataEnd As Integer)

		'Dim strInsert As String
		'Dim strValues As String
		'Dim SubAccnt As String
		Dim i As Integer
		Dim j As Integer
		Dim NData As Integer
		Dim RowFD As DataRow 'строка из таблицы с именами полей и данными
		Dim RowForInsrt As DataRow 'строка записи для вставки
		Dim ClmnForInsrt As DataColumn
		Dim strSQL As String
		Dim dsTblForLoad As DataSet 'временная таблица для вставки данных
		Dim adpImpTbl As SqlDataAdapter
		dsTblForLoad = New DataSet("dsTblForLoad")

		Dim cmdInsrt As New SqlCommand
		'формируем датасет, датаадаптер и временную таблицу для загрузки
		adpImpTbl = New SqlDataAdapter("SELECT * From " + ImpTbl, cnForLoad)
		adpImpTbl.FillSchema(dsTblForLoad, SchemaType.Mapped, ImpTbl)
		'Dim Builder As New SqlCommandBuilder(adpImpTbl)
		'Builder.GetUpdateCommand()
		'adpImpTbl.InsertCommand = CmdBldr(dsTblForLoad, ImpTbl, "Insert")
		'adpImpTbl.DeleteCommand = CmdBldr(dsTblForLoad, ImpTbl, "Delete")
		'adpImpTbl.UpdateCommand = CmdBldr(dsTblForLoad, ImpTbl, "Update")
		'вносим записи во временную таблицу для загрузки
		strSQL = "FName = '" + ExmpOrdererFile.FName + "' and (TblName = 'Any' or TblName = '" + ExpTbl + "')"
		For NData = NDataBeg To NDataEnd
			RowForInsrt = dsTblForLoad.Tables(ImpTbl).NewRow
			For Each RowFD In tblFieldsData.Select(strSQL)
				If IsDBNull(RowFD("FldData")) Then
					With oXL.Workbooks(ExmpOrdererFile.FName).Worksheets(1)
						If DtInClmn Then
							If Len(RTrim(CStr(.Cells(RowFD("FldNumber"), NData).Value))) = 0 Then
								If dsTblForLoad.Tables(ImpTbl).Columns(RowFD("Fld")).DataType = System.Type.GetType("System.String") Then RowForInsrt(RowFD("Fld")) = "" Else RowForInsrt(RowFD("Fld")) = DBNull.Value
							Else
								RowForInsrt(RowFD("Fld")) = .Cells(RowFD("FldNumber"), NData).Value
							End If
						Else
							If Len(RTrim(CStr(.Cells(NData, RowFD("FldNumber")).Value))) = 0 Then
								If dsTblForLoad.Tables(ImpTbl).Columns(RowFD("Fld")).DataType = System.Type.GetType("System.String") Then RowForInsrt(RowFD("Fld")) = "" Else RowForInsrt(RowFD("Fld")) = DBNull.Value
							Else
								RowForInsrt(RowFD("Fld")) = .Cells(NData, RowFD("FldNumber")).Value
							End If
						End If
					End With
				Else
					RowForInsrt(RowFD("Fld")) = RowFD("FldData")
				End If
			Next
			dsTblForLoad.Tables(ImpTbl).Rows.Add(RowForInsrt)
		Next
		'adpImpTbl.Update(dsTblForLoad)

		'формируем команду на вставку в таблицу в базе данных
		cmdInsrt = CmdBldr(dsTblForLoad, ImpTbl, "Insert")
		cmdInsrt.Connection = cnForLoad
		cmdInsrt.CommandType = CommandType.Text
		cnForLoad.Open()
		For Each RowForInsrt In dsTblForLoad.Tables(ImpTbl).Rows
			For Each ClmnForInsrt In dsTblForLoad.Tables(ImpTbl).Columns
				cmdInsrt.Parameters.Item("@" + ClmnForInsrt.ColumnName).Value = RowForInsrt(ClmnForInsrt)
			Next
			cmdInsrt.ExecuteNonQuery()
		Next
		cmdInsrt = Nothing
		cnForLoad.Close()
	End Sub

	Function CmdBldr(ByRef dsTbl As DataSet, ByRef Tbl As String, ByVal TypeCmd As String) As SqlCommand
		'Dim cmdBlder As SqlCommand 'объект команда
		Dim PrmType As Type
		Dim dbType As SqlDbType
		Dim length As Integer
		Dim strInsert As String
		Dim strUpdate As String
		'Dim strInsertClmn As String
		If TypeCmd = "Insert" Then
			Dim strClmnName As String
			Dim strInsertPrmt As String = ""
			Dim i As Integer
			Dim ClmnCount As Integer
			ClmnCount = dsTbl.Tables(Tbl).Columns.Count 'определили количество столбцов в таблице
			strInsert = "INSERT INTO " + Tbl + " ("
			For i = 0 To ClmnCount - 1
				strClmnName = dsTbl.Tables(Tbl).Columns(i).ColumnName
				strInsert = strInsert + strClmnName
				strInsertPrmt = strInsertPrmt + "@" + strClmnName
				If i < ClmnCount - 1 Then
					strInsert = strInsert + ", "
					strInsertPrmt = strInsertPrmt + ", "
				End If
			Next
			strInsert = strInsert + ") VALUES (" + strInsertPrmt + ")"
			CmdBldr = New SqlCommand(strInsert)
			For i = 0 To ClmnCount - 1
				strClmnName = dsTbl.Tables(Tbl).Columns(i).ColumnName
				PrmType = dsTbl.Tables(Tbl).Columns(i).DataType
				dbType = GetDbType(PrmType)
				length = dsTbl.Tables(Tbl).Columns(i).MaxLength
				CmdBldr.Parameters.Add("@" + strClmnName, dbType, length, strClmnName)
			Next
		ElseIf TypeCmd = "Delete" Then
			CmdBldr = New SqlCommand("Delete From " + Tbl)
		ElseIf TypeCmd = "Update" Then
			strUpdate = "Update " + Tbl + " Set ;"
			CmdBldr = New SqlCommand(strUpdate)
		Else
			CmdBldr = New SqlCommand("")
		End If
		Return CmdBldr
	End Function

	Public Function GetDbType(ByVal myType As System.Type) As SqlDbType
		Dim mySqlDbType As SqlDbType = SqlDbType.Int
		If myType.Equals(GetType(Integer)) Then
			mySqlDbType = SqlDbType.Int
		ElseIf myType.Equals(GetType(Long)) Then
			mySqlDbType = SqlDbType.BigInt
		ElseIf myType.Equals(GetType(Double)) Then
			mySqlDbType = SqlDbType.Float
		ElseIf myType.Equals(GetType(Decimal)) Then
			mySqlDbType = SqlDbType.Decimal
		ElseIf (myType.Equals(GetType(DateTime))) Then
			mySqlDbType = SqlDbType.DateTime
		ElseIf (myType.Equals(GetType(Boolean)) Or myType.Equals(GetType(Byte)) Or myType.Equals(GetType(Byte))) Then
			mySqlDbType = SqlDbType.Binary
		ElseIf (myType.Equals(GetType(String))) Then
			mySqlDbType = SqlDbType.Char
		Else
			Throw New ArgumentException("the DBtype could not be found!")
		End If
		Return mySqlDbType
	End Function

	Public Sub RowsDeleter(ByRef ArRws As DataRow())
		Dim i As Integer
		Dim Count As Integer
		Count = ArRws.Count
		For i = 0 To Count - 1
			ArRws(i).Delete()
		Next
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
		For j = 0 To UBound(aImpFunct)
			If aImpFunct(j).vImpFunct = True Then
				CmdSP.CommandText = aImpFunct(j).NImpFunct
				CmdSP.Transaction = CnSP.BeginTransaction(IsolationLevel.ReadUncommitted)
				Try
					CmdSP.ExecuteNonQuery()
					'CmdSP.Transaction.Rollback()
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
			Stop
			CmdSP.Transaction.Rollback()
			Stop
		End If
		CnSP.Close()
	End Sub

	Protected Sub MoveFile(ByVal PathToDir As String, ByVal FName As String, ByVal FDate As Date)
		'при отсутствии в родительском каталоге папки год и субпапки месяц, их создает и перемещает туда файл
		Dim DirYear As String
		Dim SubDirMonth As String

		FileIO.FileSystem.CurrentDirectory = FileIO.FileSystem.GetParentPath(PathToDir)
		DirYear = FileIO.FileSystem.CurrentDirectory + "\" + CStr(Year(FDate)) + "\"
		SubDirMonth = DirYear + MonthName(Month(FDate)) + "\"
		If Not FileIO.FileSystem.DirectoryExists(DirYear) Then FileIO.FileSystem.CreateDirectory(DirYear)
		If Not FileIO.FileSystem.DirectoryExists(SubDirMonth) Then FileIO.FileSystem.CreateDirectory(SubDirMonth)
		FileIO.FileSystem.MoveFile(PathToDir & FName, SubDirMonth & FName)
	End Sub

	Private Shared Sub OnInfoMessage(sender As Object,
	args As SqlInfoMessageEventArgs)
		Dim err As SqlError
		Dim ex As Exception
		For Each err In args.Errors
			Console.WriteLine("The {0} has received a severity {1}, _  
	      state {2} error number {3}\n" &
				"on line {4} of procedure {5} on server {6}:\n{7}",
				err.Source, err.Class, err.State, err.Number, err.LineNumber,
			err.Procedure, err.Server, err.Message)
		Next
		Stop
		'Return ex
	End Sub
End Class

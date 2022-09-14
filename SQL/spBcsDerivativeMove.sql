-- ================================================
-- Template generated from Template Explorer using:
-- Create Procedure (New Menu).SQL
--
-- Use the Specify Values for Template Parameters 
-- command (Ctrl-Shift-M) to fill in the parameter 
-- values below.
--
-- This block of comments will not be included in
-- the definition of the procedure.
-- ================================================
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Kutuev Rinat>
-- Create date: <2021-03-21>
-- Description:	<Ќа основе записей таблицы ImpBcsDerivativeMove дополн€ет новые инструменты, новые типы операций, новые торговые системы,'
--             'вводит операции покупки-продажи и переоценки позиций по производным инструментам >
-- =============================================
CREATE PROCEDURE spImpBcsDerivativeMove 
	-- Add the parameters for the stored procedure here
	--<@Param1, sysname, @p1> <Datatype_For_Param1, , int> = <Default_Value_For_Param1, , 0>, 
	--<@Param2, sysname, @p2> <Datatype_For_Param2, , int> = <Default_Value_For_Param2, , 0>
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
  DECLARE
    @DeliveryFutures Int = 3,
		@SetlementFutures Int = 4,
		@DefaultInstrType Int = 2, --индекс
		@DefaultActivType Int = 7, -- тип актива по умолчанию - cобственные средства
		@DefaultCur SmallInt = 840, 
		@DefaultLot Int = 1,
		@DefaultStep Int = 1,
		@DefaultStepPr Int = 1;
  DECLARE 
    @FinInstrBazActiv TABLE 
		  (
			  Id BIGINT, 
			  FinInstr NCHAR(50),
			  ActivTypeId SMALLINT,
			  InstrTypeId SMALLINT 
		  );
    
    	-- 1. ¬ставка новых групп производных инструментов в справочник групп  
	INSERT INTO RefDerivativeGr
	SELECT SUBSTRING(I.FinInstr, 1, CHARINDEX('-', I.FinInstr) -1) AS NameGr
	FROM ImpBcsDerivativeMove I
	LEFT JOIN RefDerivativeGr RDGr ON SUBSTRING(I.FinInstr, 1, CHARINDEX('-', I.FinInstr) -1) = RDGr.NameDerivativeGr
	WHERE RDGr.NameDerivativeGr IS NULL
	GROUP BY SUBSTRING(I.FinInstr, 1, CHARINDEX('-', I.FinInstr) -1);

  -- 2. ¬ставка нового базового актива в FinInstr
	INSERT INTO FinInstr (Id, FinInstr, ActivTypeId, BazActivId, InstrTypeId)
	OUTPUT Inserted.Id, Inserted.FinInstr, Inserted.ActivTypeId, Inserted.InstrTypeId INTO  @FinInstrBazActiv
	SELECT (SELECT (MAX(F1.Id))	FROM FinInstr AS F1) + ROW_NUMBER() OVER(ORDER BY RDGr.NameDerivativeGr) AS Id, 
			RDGr.NameDerivativeGr AS FinInstr, @DefaultActivType AS ActivTypeId, 
		   (SELECT (MAX(F1.Id))	FROM FinInstr AS F1) + ROW_NUMBER() OVER(ORDER BY RDGr.NameDerivativeGr) AS BazActivId, 
			@DefaultInstrType AS InstrTypeId, LEN(RDGR.NameDerivativeGr)
	FROM RefDerivativeGr RDGr
	LEFT JOIN DerivativeGr DGr ON RDGr.NameDerivativeGr = DGr.NameDerivativeGr
  LEFT JOIN FinInstr F ON RDGR.NameDerivativeGr = SUBSTRING(F.FinInstr,1,LEN(RTRIM(RDGR.NameDerivativeGr)))
	WHERE DGr.NameDerivativeGr IS NULL and F.FinInstr IS NULL;
	
END
GO

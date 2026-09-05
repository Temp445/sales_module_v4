/* All Scalar-valued Functions List */
GO

/****** Object:  UserDefinedFunction [dbo].[CDegree]    Script Date: 17/06/2026 11:43:44 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE Function [dbo].[CDegree](@Binary Float)
returns varchar(20)
as
begin 
	Declare @X float
  Declare @Y float
  Declare @Z float
	Declare @sDegree varchar(20)
  
  	select @X = convert(int,@Binary)
  	select @Y = @Binary - @X
		select @Y = convert(int,@Y * 60) 
    select @Z = @Binary - (@X + (@Y / 60))
    select @Z = convert(int,@Z * 60) 
  
    select @sDegree = convert(varchar,@X)+Char(176)+convert(varchar,round(@Y,2))+'''' + convert(varchar,round(@Z,2))+''''''
  
		return(@sDegree)
End
GO

/****** Object:  UserDefinedFunction [dbo].[fn_AbsentDays]    Script Date: 17/06/2026 11:43:44 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO





CREATE FUNCTION  [dbo].[fn_AbsentDays](@EmployeeCode varchar(10), @FDate Datetime, @TDate Datetime,@Branchid Varchar(3), @CompID varchar(3))  
RETURNS float 

AS  


BEGIN 
-- For Allowance 

	declare @AD as float
	declare @DOJ as Datetime

	select @DOJ = DateofJoining from PAY_Employee_Master where EmployeeCode =@EmployeeCode and Branchid = @Branchid 
	and Compid = @CompID
	
	if convert(varchar(7), @FDate,120) = convert(varchar(7), @DOJ,120) 

		select @AD = (Day(@DOJ) - count(LeaveDays) + 1)   
				from PAY_Attendance_Entry a where LeaveDate between @FDate and @TDate 
				and a.EmployeeCode =@EmployeeCode and a.AbsentType = 'A'  and a.Branchid = @Branchid and a.Compid = @CompID
	else

			select @AD = count(LeaveDays)  
				from PAY_Attendance_Entry a where LeaveDate between @FDate and @TDate 
				and a.EmployeeCode =@EmployeeCode and a.AbsentType = 'A'  and a.Branchid = @Branchid and a.Compid = @CompID
	
	return (@AD)

END










GO

/****** Object:  UserDefinedFunction [dbo].[fn_ActualPresentDays]    Script Date: 17/06/2026 11:43:44 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER ON
GO





CREATE FUNCTION  [dbo].[fn_ActualPresentDays] (@EmployeeCode varchar(10), @FDate Datetime, @TDate Datetime,@Branchid Varchar(3), @CompID varchar(3))  
RETURNS float 

AS  


BEGIN 
-- For Allowance 

	declare @APD as float
	declare @DOJ as Datetime

	select @DOJ = DateofJoining from PAY_Employee_Master where EmployeeCode =@EmployeeCode and Branchid = @Branchid 
	and Compid = @CompID
	
	if convert(varchar(7), @FDate,120) = convert(varchar(7), @DOJ,120) 

		select @APD = (Day(@DOJ) - count(LeaveDays) + 1)   
				from PAY_Attendance_Entry a where LeaveDate between @FDate and @TDate 
				and a.EmployeeCode =@EmployeeCode and a.AbsentType = 'A'  and a.Branchid = @Branchid and a.Compid = @CompID
	else

			select @APD = count(LeaveDays)  
				from PAY_Attendance_Entry a where LeaveDate between @FDate and @TDate 
				and a.EmployeeCode =@EmployeeCode and a.AbsentType = 'A'  and a.Branchid = @Branchid and a.Compid = @CompID
	
	return (@APD)

END









GO

/****** Object:  UserDefinedFunction [dbo].[fn_BalanceCOffDays]    Script Date: 17/06/2026 11:43:44 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION  [dbo].[fn_BalanceCOffDays] (@AsOnDate datetime, @EmployeeCode varchar(10), @Branchid Varchar(3), @CompID varchar(3))  
RETURNS float 
AS  
BEGIN 

Declare @COff float

select @COff = (COff+isNull((Select LeaveDays from PAY_Leave_Details 
			Where EmployeeCode=@EmployeeCode and LeaveCode='CO' and LeaveYear=Year(@AsOnDate) and BranchID=@BranchID and CompID=@CompID), 0)
			-COffTakenDays) 
 from (

	select sum(Case When ExtraHrs > FromMinutes then COff Else 0 End) COff,
		isnull((select sum(e.LeaveDays) 
			from PAY_Attendance_Entry e 
			where Year(e.LeaveDate) =  year(@AsOnDate) and e.AbsentType = 'A' and e.LeaveType = 'COF' 
				and e.EmployeeCode = @EmployeeCode and e.Branchid = @Branchid 
				and e.Compid = @CompID),0) COffTakenDays
	from
		(Select a.EmployeeCode, a.ExtraHrs, a.CurrentShift, isNull(b.HolidayType, 'Week Day') HolidayType, a.Branchid, a.Compid 
		from PAY_Time_entry a 
		left outer join PAY_Holiday_Dtl b
			on a.DateIn = b.HolidayDate and a.Branchid = b.Branchid and a.Compid = b.Compid
		where a.DateIn between DateAdd(m,-3,@AsOnDate) and @AsOnDate and a.Branchid = @Branchid and a.Compid = @CompID
		) a1 

	left outer join PAY_OT_Parameter c 
		on a1.HolidayType = c.HolidayType and c.Branchid = a1.Branchid and c.Compid = a1.Compid 
		where a1.EmployeeCode = @EmployeeCode and a1.Branchid = @Branchid and a1.Compid = @CompID
	) x

return (@COff)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_BalanceLeaveDays]    Script Date: 17/06/2026 11:43:44 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION  [dbo].[fn_BalanceLeaveDays] (@LeaveAsOnDate datetime, @EmployeeCode varchar(10), @LeaveCode varchar(3), @Branchid Varchar(3), @CompID varchar(3))  

RETURNS float 

 AS  


BEGIN 
	Declare @BLD float

	select @BLD = BalanceLeave 
	from (
		select case when x.CreditPolicy = 'E' then
				(round((x.TotalDays - x.AbsentDays)/ x.NoofDaysworked,-1)* x.NoofDaysCredited) - LeaveTakenDays
			else x.LeaveDays - x.LeaveTakenDays End BalanceLeave 
		from (			
		
			select a.LeaveCode, a.LeaveDays, b.NoofDaysworked,b.NoofDaysCredited,b.CreditPolicy,
			case when Year(c.DateofJoining) =  year(@LeaveAsOnDate) then DateDiff(d,c.DateofJoining,@LeaveAsOnDate) 
				else DateDiff(d, convert(varchar,year(@LeaveAsOnDate))+'-01-01',@LeaveAsOnDate) end TotalDays,		
			isNull((select sum(d.LeaveDays) from PAY_Attendance_Entry d 
			where d.AbsentType = 'A' and d.EmployeeCode = a.EmployeeCode and d.Branchid = a.Branchid and d.Compid = a.Compid), 0) AbsentDays,
		
			isnull((select sum(e.LeaveDays) from PAY_Attendance_Entry e 
				where Year(e.LeaveDate) =  year(@LeaveAsOnDate) and e.AbsentType = 'A' and e.LeaveType = a.LeaveCode  
					and LeaveDate <= @LeaveAsOnDate and e.EmployeeCode = a.EmployeeCode 
					and e.Branchid = a.Branchid and e.Compid = a.Compid),0) LeaveTakenDays		
			from PAY_Leave_Details a 
			left outer join PAY_Leave_Master b 
				on a.LeaveCode = b.LeaveCode and a.Branchid = b.Branchid and a.Compid = b.Compid
			left outer join PAY_Employee_Master c 
				on a.EmployeeCode = c.EmployeeCode and a.Branchid = c.Branchid and a.Compid = c.Compid
			where a.LeaveYear =  year(@LeaveAsOnDate) and a.EmployeeCode = @EmployeeCode and a.LeaveCode =@LeaveCode
				and a.Branchid = @Branchid and a.Compid = @CompID
			) x 
		) y

	return (@BLD)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_BaseConvert]    Script Date: 17/06/2026 11:43:44 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_BaseConvert] (@Input nvarchar(500), @BaseFrom tinyint, @BaseTo tinyint)  

RETURNS varchar(500)

 AS  

BEGIN 
	--BaseFrom and BaseTo values should be between 1 and 62

  	Declare @QuotValue float
  	Declare @RemValue nvarchar(500)
  	Declare @LpCnt int
  	Declare @DecValue float
  	Declare @singChar nChar(1)
	Declare @singCharAsc int
  	Declare @Result nvarchar(500)
  


	Select  @DecValue = 0
	Select @LpCnt = 1

	--Here Any converts any base to Base 10
	While @LpCnt <= Len(@Input)
		Begin  
			Select @singCharAsc = Ascii(Substring(@Input, @LpCnt, 1))

    			Select @DecValue = Case 
      						When @singCharAsc between 48 and 57 then
							((@singCharAsc - 48) * Power(@BaseFrom,  (Len(@Input) - @LpCnt))) + @DecValue
      						When @singCharAsc between 65 and 90 then
       							((@singCharAsc - 55) * Power(@BaseFrom, (Len(@Input) - @LpCnt))) + @DecValue
      						When @singCharAsc between 97 and 122 then
        							((@singCharAsc - 61) * Power(@BaseFrom,  (Len(@Input) - @LpCnt))) + @DecValue
    						End
 
			Select @LpCnt = @LpCnt + 1
 		End

	If (@BaseTo = 10)
		Select @Result = Convert(varchar, @DecValue)
	Else
		Begin
  			Select @RemValue = '', @QuotValue = @DecValue
 			
			While @QuotValue >= @BaseTo
				Begin
    					Select @singCharAsc =  Convert(int, @QuotValue) % @BaseTo

    					If (@singCharAsc between 10 and  35) 
      						Select @singChar = NChar(@singCharAsc  +  55)
					Else
    						If (@singCharAsc between 36 and 61) 
      							Select @singChar = NChar(@singCharAsc +  61)
						Else
							Select @singChar = Convert(nvarchar, @singCharAsc)
    
   					Select  @RemValue = @singChar + @RemValue
    					Select @QuotValue = Convert(Int, (@QuotValue / @BaseTo))
  				End

  			If (@QuotValue between 10 and  35) 
    				Select @singChar = NChar(@QuotValue + 55)
  			Else
				If (@QuotValue between 36 and 61) 
    					Select @singChar = NChar(@QuotValue + 61)
  				Else
    					Select @singChar = @QuotValue
  
  
  			Select @Result = @singChar + @RemValue
		End

	Return @Result
END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_BOM_Get_ItemRate]    Script Date: 17/06/2026 11:43:44 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_BOM_Get_ItemRate] (@fOpration Char(1), @ItemCode varchar(20), @AsOnDate datetime, @RptDate datetime, @BranchID varchar(3), @CompID varchar(3))      
RETURNS money  /* Returns ItemRate from Material Master*/    
AS      
BEGIN    
    
 DECLARE @ItemRate Float    
     
  Select @ItemRate=Rate      
  From INV_MATERIAL_MASTER a  --dbo.fn_BOM_Material_Master(@fOpration, @AsOnDate, @RptDate, @BranchID, @CompID) a    
  Where a.ItemCode=@ItemCode and a.BranchID=@BranchID and a.CompID=@CompID     
     
 RETURN @ItemRate;    
END  
GO

/****** Object:  UserDefinedFunction [dbo].[fn_BOM_Norms]    Script Date: 17/06/2026 11:43:44 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION  [dbo].[fn_BOM_Norms] (@FromItem varchar(20), @ToItem varchar(20), @AsOnDate datetime, 
@RptDate datetime, @Branchid Varchar(3), @CompID varchar(3))  
RETURNS float 
 AS  
 BEGIN 
	/*Declare @Norms money
	Select @Norms=IsNull( (Select Qty from dbo.fn_BOM_Structure_Dtl('A', @AsonDate, @RptDate, @BranchID, @CompID)
			            Where ParentCode = (Select Top 1 ItemCode From dbo.fn_BOM_MPS_Master('A', @AsonDate, @RptDate, @BranchID, @CompID)
						        Where StageCode=@ToItem and Branchid = @Branchid and compid=@CompID
						      )
				and ChildCode = @FromItem and Branchid = @Branchid and CompID=@CompID), 1)
	Return (@Norms)*/
	--Changed by rajapriyan on 04/02/2011 2:25 PM
	Declare @SNorms float --Set Norms
	Declare @GNorms float --Get Norms
	Declare @tmpitem varchar(20)
	/*Declare cur cursor for
	Select ItemCode from [dbo].[fn_BOM_ProductExplosion] (@ToItem,  @AsonDate, @RptDate, @BranchID, @CompID)  
	where ItemType='C'	
	open cur	
	Fetch Next From cur into @tmpitem*/
	Select @SNorms=1, @GNorms=1
	
	While isnull(@SNorms,-1)<>-1
	 Begin
		Select top 1 @GNorms=a.Qty,@FromItem=a.ParentCode 
		From [dbo].[fn_BOM_ProductExplosion] (@ToItem, @BranchID, @CompID) a 
		Where (a.ItemCode=@FromItem and a.ItemType='C') or (a.ItemCode=@FromItem and a.ItemType='S' and a.StructureLevel<>1) 
		Order by a.Itemtype

		If (@GNorms Is Null)
			break
		Else
			Select @SNorms=(@SNorms*@GNorms)
	
		Select @GNorms=null
	 End
 Return @SNorms 
END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_BOM_TotalRMCost]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_BOM_TotalRMCost] (@ItemCode varchar(20), @BranchID varchar(3),  @CompID varchar(3))  

RETURNS float 

AS  

BEGIN 

Declare @RMCost float

	Select @RMCost = Sum(m.Qty * n.Rate) 
 		From dbo.fn_BOM_GetBottomMostChild(@ItemCode, @BranchID, @CompID)m 
		Left Outer Join Inv_Material_Master n
		On n.ItemCode = m.ItemCode and m.BranchId = n.BranchID and m.CompID = n.CompID
		--Group by m.ParentCode

	Return @RMCost

END


GO

/****** Object:  UserDefinedFunction [dbo].[fn_COM_Get_DocNoWithPrefixSuffix]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_COM_Get_DocNoWithPrefixSuffix] (@DocType varchar(3), @DocNo int, @DocDate datetime, @PertainingToID int,   
@IsWidthAvoided bit='0', @VerticalID int, @ModuleCode int, @AsOnDate datetime, @RptDate datetime, @BranchID varchar(3),   
@CompID varchar(3), @Finyear varchar(20))   

RETURNS Varchar(40)     
   
AS      
BEGIN      

 Declare  @StrFix Varchar(40)
 Declare @DocNoFormatted As varchar(300);    
 Declare @WidthNo As int;   
 Declare @Prefix As varchar(15);   
 Declare @Suffix As varchar(15);   

 SET @StrFix =''
     
 --Select @WidthNo=Len(Convert(varchar(20),@DocNo)), @Prefix='', @Suffix='-%FS%/%FE%';     
 Select @WidthNo=0, @Prefix=@DocType + '-', @Suffix='-%FS%/%FE%';      
 Select @DocNoFormatted=dbo.GetFormatedDocNoWithPrefixSuffix(@DocDate, @DocNo, @WidthNo, @Prefix, @Suffix, @IsWidthAvoided)      
  
 /* Output */  
 Select @StrFix=IsNull(@DocNoFormatted, Convert(varchar, @DocNo)) --As DocNo       
      
/*  
 Declare @DocNoFormatted As varchar(300);      
      
 Select @DocNoFormatted=dbo.GetFormatedDocNoWithPrefixSuffix(@DocDate, @DocNo, WidthNo, Prefix, Suffix)      
 From dbo.fn_COM_DocumentType_Master('M', @AsOnDate, @RptDate, @BranchID, @CompID)      
 Where DocumentType = @DocType and PertainingToID =@PertainingToID and VerticalID = @VerticalID and ModuleCode=@ModuleCode      
      
 Select IsNull(@DocNoFormatted, Convert(varchar, @DocNo)) as DocNo;      
*/      
RETURN (@StrFix)  


END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_COM_Get_FinStart]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_COM_Get_FinStart](@TranDate datetime)
RETURNS datetime
AS
BEGIN
	DECLARE @FinStart datetime;

	If (Month(@TranDate)<4)
		SET @FinStart=Convert(datetime, Convert(varchar, (Year(@TranDate)-1)) + '-04-01', 120);
	Else
		SET @FinStart=Convert(datetime, Convert(varchar, Year(@TranDate)) + '-04-01', 120);
	
	RETURN (@FinStart)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_COM_Get_Finyear]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION  [dbo].[fn_COM_Get_Finyear](@TranDate datetime)    
RETURNS varchar(15)   
AS    
BEGIN   
 DECLARE @Finyear varchar(15);  
  
 If (Month(@TranDate)<4)  
  SET @Finyear='04-'+Convert(varchar, Year(@TranDate)-1)+' 03-'+Convert(varchar, Year(@TranDate));  
 Else  
  SET @Finyear='04-'+Convert(varchar, Year(@TranDate))+' 03-'+Convert(varchar, Year(@TranDate)+1);  
   
 RETURN (@Finyear)  
END 
GO

/****** Object:  UserDefinedFunction [dbo].[fn_COM_Get_HeatNo]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_COM_Get_HeatNo] (@DocType Varchar(3), @DocNo int, @DocDate datetime, @BranchID varchar(3),  
	@CompID varchar(3), @Finyear Varchar(20))    
RETURNS Varchar(20)     
AS      
BEGIN     
 Declare @HeatNo Varchar(20)    
    
 Select @HeatNo = IsNull(c.HeatNo,'')     
 From     
 (    
  Select TranDocId, ICode, DOCType, DOCNo, DOCDate, TransactionDate, TransactionTime, TransactionItem, TransactionQty, Norms,   
   FromWarehouseCode, FromRackBinCode, ToWarehouseCode, ToRackBinCode, ItemGroupKey, FromItem, ToItem, Justification, LIFO, FIFO, WA,   
   Userid, Branchid, Compid, Finyear 
  From INV_MaterialTran_Dtl a                  
  Where a.DocType=@DocType and a.DocNo=@DocNo and a.DocDate=@DocDate and a.BranchID=@BranchID and a.CompID=@CompID and a.Finyear=@Finyear    
    
  Union All    
    
  Select TranDocId, -(ICode), DOCType, DOCNo, DOCDate, TransactionDate, TransactionTime, TransactionItem, TransactionQty, Norms,   
   FromWarehouseCode, FromRackBinCode, ToWarehouseCode, ToRackBinCode, ItemGroupKey, FromItem, ToItem, Justification, LIFO, FIFO, WA,   
   Userid, Branchid, Compid, Finyear     
  From INV_WIPTRAN_DTL a         
  Where a.DocType=@DocType and a.DocNo=@DocNo and a.DocDate=@DocDate and a.BranchID=@BranchID and a.CompID=@CompID and a.Finyear=@Finyear                       
   and not exists    
   (    
    Select * From INV_MaterialTran_Dtl s    
    Where s.DocNo=a.DocNo and s.DocDate=a.DocDate and s.DocType=a.DocType and s.TransactionDate=a.TransactionDate     
     and s.TransactionTime=a.TransactionTime and s.TransactionItem=a.TransactionItem and s.FromwarehouseCode=a.FromWarehouseCode    
     and s.FromRackBinCode=a.FromRackBinCode and s.TowarehouseCode=a.TowarehouseCode and s.ToRackBinCode=a.ToRackBinCode     
     and s.ItemGroupKey=a.ItemGroupKey and s.FromItem=a.FromItem and s.ToItem=a.ToItem     
     and s.BranchID=a.BranchID and s.CompID=a.CompID and s.Finyear=a.Finyear    
   )     
 ) a      
 Left Outer Join COM_Trace_MaterialTran_TmpQty b    
  On a.ICode=b.TranID and a.BranchID=b.BranchID and a.CompID=b.CompID and a.Finyear=b.Finyear    
 Inner Join COM_Trace_LotMaster_TmpQty c    
  On b.LotID=c.LotID and b.BranchID=c.BranchID and b.CompID=c.CompID    
 Where a.DocType=@DocType and a.DocNo=@DocNo and a.DocDate=@DocDate and CONVERT(NUMERIC(38,0), CAST(b.Qty AS FLOAT))>0  
  and a.BranchID=@BranchID and a.CompID=@CompID and a.Finyear=@Finyear    
    
 RETURN @HeatNo;    
END 
GO

/****** Object:  UserDefinedFunction [dbo].[fn_COM_Get_LatestHeatNo_StockVerification]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_COM_Get_LatestHeatNo_StockVerification] (@ItemCode Varchar(20), @AsOnDate datetime, @BranchID varchar(3),  @CompID varchar(3), @Finyear Varchar(20))  
RETURNS Varchar(20)   
AS    
BEGIN   
	Declare @HeatNo Varchar(20)  
		Select @HeatNo = ''

	Select Top 1 @HeatNo = IsNull(HeatNo,'') 
	From COM_Trace_LotMaster_TmpQty a
	Where Transtamp>='2024-03-31 23:59:59.000' and a.Transtamp<=@AsOnDate and a.HeatNo<>'' and a.ItemCode=@ItemCode
	Order by a.Transtamp Desc

	IF @HeatNo=''
		BEGIN
			DECLARE @CurDate datetime
				Select @CurDate=getdate()

			DECLARE @TblStage Table (StageCode Varchar(20))	
				Insert into @TblStage
					Select b.StageCode  
					From fn_BOM_MPS_Master('A',@CurDate,@CurDate,@BranchID,@CompID)b	 
					Where b.ItemCode in (Select s.ItemCode  
											From fn_BOM_MPS_Master('A',@CurDate,@CurDate,@BranchID,@CompID)s	 
											Where s.StageCode=@ItemCode)

			Select Top 1 @HeatNo = IsNull(HeatNo,'') 
			From COM_Trace_LotMaster_TmpQty a
			Where Transtamp>='2024-03-31 23:59:59.000' and a.Transtamp<=@AsOnDate and HeatNo<>''
				and ItemCode in 
					(
						Select StageCode
						From @TblStage
						Union 
						Select ChildCode
						From BOM_Structure_Dtl a
						Where Exists (Select * From @TblStage s Where s.StageCode=a.ParentCode and a.BranchID=@BranchID and a.CompID=@CompID)
					)
			Order by a.Transtamp Desc
			
			/*
			Select Top 1 @HeatNo = IsNull(HeatNo,'') 
			From COM_Trace_LotMaster_TmpQty a
			Where Transtamp>='2024-03-31 23:59:59.000' and a.Transtamp<@AsOnDate and HeatNo<>''
				and ItemCode in (Select b.StageCode  
									From fn_BOM_Process(@BranchID,@CompID)b	 
									Where b.ParentCode in (Select s.ParentCode  
															From fn_BOM_Process(@BranchID,@CompID)s	 
															Where s.StageCode=@ItemCode)
								)
			Order by a.Transtamp Desc
			*/
		END
 
	RETURN @HeatNo;  
END  
GO

/****** Object:  UserDefinedFunction [dbo].[fn_COM_ItemRate_Pur]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_COM_ItemRate_Pur] (@ItemCode varchar(20), @RateType tinyint=0, @PeriodStart datetime=null,             
@PeriodEnd datetime=null, @isAnyRateReturn bit='0', @AsonDate datetime,@RptDate datetime,            
@BranchID varchar(3), @CompID varchar(3), @FinYear varchar(20))              
/* @RateType 0 - ItemMaster Rate; 1 - Weighted Average Rate for The Given Period ; 2 - Vendor Item Master Rate; 3 - Latest Order Rate of Last receipt / Inward in Given Period; 4 - Maximum Rate in Given Period Orders */              
RETURNS money   /* Returns Purchase Item Rate by given option (Order, VendorItemMaster, Wt.Avg.Rate, Material Master)*/              
AS              
BEGIN              
                
  Declare @Rate as money              
              
  If @RateType=4   /* 4 - Maximum Rate in Given Period Orders */              
   Begin              
    Select @Rate=Max(Rate/RatePer)              
    From PUR_PurchaseOrder_Dtl a               
    Where a.ItemCode=@ItemCode and a.BranchID = @BranchID  and a.CompID = @CompID and isActive='1'              
     and (a.OrderDate Between @PeriodStart and @PeriodEnd)              
    Group By a.ItemCode, a.BranchID, a.CompID              
   End              
              
  Else              
  If @RateType=3   /* 3 - Latest RR Accepted based Order Rate */              
 -- Inward Type 'U' added for Infant - MRP Analytical Report Few item Inwarded thru Customer Supplied Product           
 -- So this function return null value if Inward Type mentioned only 'P'          
   Begin              
    ;With MaxRRDate_CTE              
    AS              
    (              
     Select b.ItemCode, Max(a.TranStamp) as MaxTranStamp              
     From REC_RECEIPTENTRY_HDR a              
     Inner Join REC_RECEIPTENTRY_Dtl b              
      On a.ReceiptEntryNo=b.ReceiptEntryNo and a.ReceiptEntryDate=b.ReceiptEntryDate and a.ReceiptEntryType=b.ReceiptEntryType              
				and a.BranchID=b.BranchID and a.CompID=b.CompID              
		Where b.ItemCode=@ItemCode and a.InwardType in ('P' , 'U', 'V')          
		--and b.GroupCode='-1' -- not in ('-1', '-2', 'Scrap', 'Rejection')              
      and a.ReceiptEntryDate<=@PeriodEnd and a.BranchID=@BranchID and a.CompID=@CompID              
     Group By b.ItemCode              
    ),              
    MaxRRDoc_CTE              
    AS              
    (              
     Select b.ItemCode, a.ReceiptEntryType, a.ReceiptEntryNo, a.ReceiptEntryDate              
     From REC_RECEIPTENTRY_HDR a              
     Inner Join REC_RECEIPTENTRY_Dtl b              
      On a.ReceiptEntryNo=b.ReceiptEntryNo and a.ReceiptEntryDate=b.ReceiptEntryDate and a.ReceiptEntryType=b.ReceiptEntryType              
       and a.BranchID=b.BranchID and a.CompID=b.CompID              
     Inner Join MaxRRDate_CTE c              
      On b.ItemCode=c.ItemCode and a.TranStamp=c.MaxTranStamp              
     Where a.InwardType in ('P' , 'U', 'V')   ---and b.GroupCode='-1' -- not in ('-1', '-2', 'Scrap', 'Rejection')              
      and a.BranchID=@BranchID and a.CompID=@CompID              
    )              
              
    Select @Rate=Rate From       
 (Select Top 1 (a.Rate/a.RatePer)  as Rate            
    From fn_PUR_PurchaseOrder_Dtl(@AsonDate, @RptDate, @BranchID, @CompID, @FinYear) a              
    Where a.ItemCode=@ItemCode --and a.BranchID=@BranchID and a.CompID=@CompID              
     and Exists   
  (  
  Select * from MaxRRDoc_CTE b              
         --Inner Join REC_IncomingReceipt_Others c               
         -- On b.DocNo=c.DocNo and b.DocDate=c.DocDate and b.DocType=c.DocType              
         Inner Join REC_RECEIPTENTRY_Dtl d               
          On b.ReceiptEntryDate=d.ReceiptEntryDate and b.ReceiptEntryNo=d.ReceiptEntryNo and b.ReceiptEntryType=d.ReceiptEntryType and d.BranchId=@BranchID and d.CompID= @CompID  
                  
         Inner join dbo.fn_PUR_Schedule(@AsonDate, @RptDate, @BranchID, @CompID, @FinYear) e            
          On e.ScheduleDate=d.RefDate and e.ScheduleNo=d.RefNo and e.ScheduleType=d.RefType and e.ItemCode=d.ItemCode               
           and e.RequiredDate=d.ScheduleDate and e.BranchID=d.BranchId and e.CompID=d.CompID          
           
         Where a.OrderDate=e.OrderDate and a.OrderNo=e.OrderNo and a.OrderType=e.OrderType and a.ItemCode=e.ItemCode              
          and e.BranchID=a.BranchId and e.CompID=a.CompID              
)                
  --Union All      
      
  -- Select Top 1 (a.Rate/a.RatePer)  as Rate        
  --From fn_SAL_JobWorkOrder_Dtl(@AsonDate, @RptDate, @BranchID, @CompID, @FinYear) a          
  --Where a.RecdItemCode=@ItemCode           
  -- and Exists (      
  --Select * from MaxRRDoc_CTE b          
  --       Inner Join REC_IncomingReceipt_Others c           
  --        On b.DocNo=c.DocNo and b.DocDate=c.DocDate and b.DocType=c.DocType          
  --       Inner Join REC_IncomingReceipt_Dtl d           
  --        On c.DocDate=d.DocDate and c.DocNo=d.DocNo and c.DocType=d.DocType and c.BranchID=d.BranchId and c.CompID=d.CompID          
  --       Inner join dbo.fn_SAL_JobWorkOrder_Dtl(@AsonDate, @RptDate, @BranchID, @CompID, @FinYear) e          
  --        On e.OrderDate=c.RefDate and e.OrderNo=c.RefNo and e.OrderType=c.RefType and e.RecdItemCode=d.ItemCode           
  --         and e.OrderDate=d.ScheduleDate and e.BranchID=c.BranchId and e.CompID=c.CompID          
  --       Where a.OrderDate=e.OrderDate and a.OrderNo=e.OrderNo and a.OrderType=e.OrderType and a.RecdItemCode=e.RecdItemCode          
  --        and e.BranchID=a.BranchId and e.CompID=a.CompID       
        --)         
  )x        
              
/*             /* 3 - Latest Order Rate in Given Period */              
    Select Top 1 @Rate=Rate/RatePer              
    From PUR_PurchaseOrder_Dtl a               
             
    Inner Join  /* To Get Maximum Order Date of Last Receipt / Inward */              
     (Select ItemCode, Max(OrderDate) as OrderDate From PUR_PurchaseOrder_Dtl  f              
     Where ItemCode=@ItemCode and BranchID = @BranchID  and CompID = @CompID and isActive='1'              
      and (OrderDate Between @PeriodStart and @PeriodEnd)               
      /*               
      and Exists (Select * from REC_IncomingReceipt_Others c               
               Inner Join REC_IncomingReceipt_Dtl d               
               On c.DocDate=d.DocDate and c.DocNo=d.DocNo and c.DocType=d.DocType and c.BranchID=d.BranchId and c.CompID=d.CompID              
               Inner join dbo.fn_PUR_Schedule(@AsonDate, @RptDate, @BranchID, @CompID, @FinYear) e              
               On e.ScheduleDate=c.RefDate and e.ScheduleNo=c.RefNo and e.ScheduleType=c.RefType and e.ItemCode=d.ItemCode               
               and e.RequiredDate=d.ScheduleDate and e.BranchID=c.BranchId and e.CompID=c.CompID              
               Where f.OrderDate=e.OrderDate and f.OrderNo=e.OrderNo and f.OrderType=e.OrderType and f.ItemCode=e.ItemCode              
               and e.BranchID=f.BranchId and e.CompID=f.CompID              
              )              
     */              
     Group By ItemCode) b              
    On a.OrderDate=b.OrderDate and a.ItemCode=b.ItemCode              
              
    Inner Join  /* To Get Maximum Order No of Last Receipt / Inward */              
     (Select ItemCode, OrderDate, Max(OrderNo) as OrderNo From PUR_PurchaseOrder_Dtl  f              
     Where ItemCode=@ItemCode and BranchID = @BranchID  and CompID = @CompID and isActive='1'              
      and (OrderDate Between @PeriodStart and @PeriodEnd)              
      /*              
      and Exists (Select * from REC_IncomingReceipt_Others c               
               Inner Join REC_IncomingReceipt_Dtl d               
               On c.DocDate=d.DocDate and c.DocNo=d.DocNo and c.DocType=d.DocType and c.BranchID=d.BranchId and c.CompID=d.CompID              
               Inner join dbo.fn_PUR_Schedule(@AsonDate, @RptDate, @BranchID, @CompID, @FinYear) e              
               On e.ScheduleDate=c.RefDate and e.ScheduleNo=c.RefNo and e.ScheduleType=c.RefType and e.ItemCode=d.ItemCode               
               and e.RequiredDate=d.ScheduleDate and e.BranchID=c.BranchId and e.CompID=c.CompID              
               Where f.OrderDate=e.OrderDate and f.OrderNo=e.OrderNo and f.OrderType=e.OrderType and f.ItemCode=e.ItemCode              
             and e.BranchID=f.BranchId and e.CompID=f.CompID              
            )              
      */              
     Group By ItemCode, OrderDate) c              
    On a.ItemCode=c.ItemCode and a.OrderNo=c.OrderNo and b.OrderDate=c.OrderDate              
              
    Inner Join  /* To Get Maximum CreatedDate */              
     (Select ItemCode, OrderDate, OrderNo, Max(CreatedDate) as CreatedDate From PUR_PurchaseOrder_Dtl  f              
     Where ItemCode=@ItemCode and BranchID = @BranchID  and CompID = @CompID and isActive='1'              
      and (OrderDate Between @PeriodStart and @PeriodEnd)              
     Group By ItemCode, OrderDate, OrderNo) d              
    On a.ItemCode=c.ItemCode and a.OrderNo=c.OrderNo and b.OrderDate=c.OrderDate and a.CreatedDate=d.CreatedDate              
              
    Where a.ItemCode=@ItemCode and a.BranchID = @BranchID  and a.CompID = @CompID and isActive='1'              
     and (a.OrderDate Between @PeriodStart and @PeriodEnd)              
*/              
   End              
              
  Else              
  If @RateType=2   /* 2 - Maximum Rate in Vendor Master */              
   Begin              
    Select @Rate=Max(a.Rate/a.RatePer)               
    From PUR_Vendor_ItemMaster a              
    Where a.ItemCode=@ItemCode and a.BranchID=@BranchID and a.CompID=@CompID and a.isActive='1' --a.EffStatus='AA'              
     and a.ValidFrom=(Select Max(b.ValidFrom) From PUR_Vendor_ItemMaster b              
       Where a.VendorCode=b.VendorCode and a.ItemCode=b.ItemCode and b.isActive='1' -- b.EffStatus='AA'              
        and a.BranchID=b.BranchID and a.CompID=b.CompID)              
   End                
              
  Else              
  If @RateType=1   /* 1 - Weighted Average Rate */              
   Begin             
    Select @Rate=dbo.fn_WtAvgRate(@ItemCode, @PeriodStart, @PeriodEnd, @BranchID, @CompID, @Finyear)      
  --Select @Rate=dbo.fn_WtAvgRate(@ItemCode, @PeriodStart, @PeriodEnd, @AsonDate,@RptDate, @BranchID, @CompID, @Finyear)            
   End              
              
  Else              
  If @RateType=0   /* 0 - Material Master Rate */              
   Begin              
    Select @Rate=Rate              
    From INV_MATERIAL_MASTER a --dbo.fn_BOM_Get_Material_Master_All('A', @AsOnDate, @RptDate, @BranchID, @CompID) a               
    Where a.ItemCode=@ItemCode and a.BranchID = @BranchID  and a.CompID = @CompID              
   End              
              
  If (@isAnyRateReturn='1' AND @RateType>0 AND isNull(@Rate, 0)=0)              
   Begin              
    Select @Rate=dbo.fn_COM_ItemRate_Pur (@ItemCode, @RateType-1, @PeriodStart, @PeriodEnd, @isAnyRateReturn,  @AsonDate, @RptDate, @BranchID, @CompID, @FinYear)                
   End              
              
  RETURN (@Rate)              
              
END 
GO

/****** Object:  UserDefinedFunction [dbo].[fn_DD_KPI_ChartOptionMaster_Get_TypeValue]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION  [dbo].[fn_DD_KPI_ChartOptionMaster_Get_TypeValue](@ChartTypeId int, @TypeFlg Varchar(20), @PeriodType char(2), @BranchID Varchar(3), @CompID Varchar(3))    
RETURNS decimal(12,2)  
AS    
BEGIN   
  
 Declare @TypeValue decimal(12,2)  
  
 Select @TypeValue= Case When ISNUMERIC(TypeValue)='1'   
        Then Case When IsTypeValuePercentage='1' Then (Convert(decimal(12,2),TypeValue)/100.0)  
                  Else Convert(decimal(12,2),TypeValue) End  
        Else 0 End  
 --Select @TypeValue =  Case When ISNUMERIC('360000000.0120')='1' Then Convert(decimal(12,2),'360000000.0120') Else 0 End  
 From DD_KPI_ChartOption_Master  
 Where ChartTypeId=@ChartTypeId and TypeFlg=@TypeFlg and PeriodType=@PeriodType and BranchId=@BranchID and CompId=@CompID  
   
 RETURN (@TypeValue)  
END  
GO

/****** Object:  UserDefinedFunction [dbo].[fn_DD_KPI_FASReports]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_DD_KPI_FASReports](@ChartType int, @LrRoute Varchar(42), @Start datetime, @End datetime, @CmCode varchar(12))    
RETURNS money  
AS    
BEGIN  
  
    Declare @Amt as money   
  
    IF (@ChartType=11 OR  @ChartType=20 OR @ChartType=21 OR @ChartType=22)  
        BEGIN  
            Select @Amt=dbo.fn_FAS_Profit_Loss_Period_SalPurExpenses(@LrRoute, @Start, @End, @CmCode)  
        END   
  
    ELSE IF (@ChartType=23)     -- Gross Profit  
        BEGIN  
            Select @Amt=dbo.fn_FAS_Profit_Loss_Period_Gross(@Start, @End, @CmCode)  
        END   
      
    ELSE IF (@ChartType=3)      -- Net Profit  
        BEGIN  
            Select @Amt=dbo.fn_FAS_Profit_Loss_Period(@Start, @End, @CmCode)  
        END   
  
 RETURN (@Amt)  
  
END  
GO

/****** Object:  UserDefinedFunction [dbo].[fn_DD_KPI_GaugeChartData_AllowInsert]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_DD_KPI_GaugeChartData_AllowInsert](@ChartTypeId int, @AsOnDate datetime, @PeriodType char(2), @TargetValue money, @PlanValue money,   
               @PlanTillPeriodValue money, @AchievedValue money, @AchievedValue2 money, @AllowVariance money,  
               @StartDate datetime, @BranchID varchar(2), @CompID varchar(2))  
RETURNS bit  
AS  
BEGIN  
  
 DECLARE @IsAllowToInsert bit;  
  SET @IsAllowToInsert='0';  
  
 IF Not Exists  
  (  
   Select * From DD_KPI_GaugeChartData   
   Where ChartTypeId = @ChartTypeId and PeriodType=@PeriodType and AsOndate between @StartDate  and @AsOndate  
    and BranchId=@BranchId and CompId=@CompId  
  )    
    
  OR   
    
  Not Exists  
  (  
   Select * From DD_KPI_GaugeChartData a  
   Where ChartTypeId = @ChartTypeId and a.PeriodType=@PeriodType  
    and Transtamp = (Select Max(Transtamp) From DD_KPI_GaugeChartData s  
         Where  s.AsOnDate<=@AsOnDate and s.ChartTypeId=a.ChartTypeId and s.PeriodType=a.PeriodType   
          and s.BranchId=a.BranchId and s.CompId=a.CompId  
        )  
    and (a.TargetValue=@TargetValue and PlanValue=@PlanValue and  PlanTillPeriodValue=@PlanTillPeriodValue and  AchievedValue=@AchievedValue   
      and AchievedValue2=@AchievedValue2 and AllowVariance=@AllowVariance)  
    and  BranchID=@BranchID and CompID=@CompID  
  )  
  
  BEGIN  
   SET @IsAllowToInsert = '1'  
  END  
  
 RETURN @IsAllowToInsert  
  
END  
  
GO

/****** Object:  UserDefinedFunction [dbo].[fn_Decrypt]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO




CREATE FUNCTION [dbo].[fn_Decrypt] (@EnCryptedText varchar(1000)) 

RETURNS varchar(1000)

 AS  

BEGIN 

	Declare @EnCryptedLength 	int,
		@FirstLock		char(2),
		@Count 		int,
		@ShiftValue		int,
		@AsciiValue		int,
		@DecryptChar		char(1),
		@DecryptedText	varchar(1000)


 	Select @EnCryptedLength=len(@EnCryptedText)
	Select @FirstLock=right(@EnCryptedText, 2)
	Select @ShiftValue = Convert(int, char(ascii(substring(@FirstLock,1,1))-110)+char(ascii(substring(@FirstLock,2,1))-110))

	Select @DecryptedText =  ''

	Select @Count=1

	While (@Count <= @EnCryptedLength-2)
		Begin
			 Select  @AsciiValue=ascii(substring(@EnCryptedText, @Count, 1))
			 Select  @DecryptChar=char(@AsciiValue-@ShiftValue - @Count)
		  	 Select  @DecryptedText=@DecryptedText + @DecryptChar
			 Select  @Count=@Count+1
		End
	Return (@DecryptedText)
END


GO

/****** Object:  UserDefinedFunction [dbo].[fn_DigitSum]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_DigitSum] (@GivenNo varchar(100))  

RETURNS Int

AS  

BEGIN 

	Declare @Nos as Varchar(100)	
	Declare @digitsum bigint 
	Declare @Cnt as int	

	Declare @SumofNo as bigint
	Select @SumofNo = @GivenNo
	
	While @SumOfNo > 9

		

		Begin
			Select @Nos = @SumOfNo
			Select @Cnt = Len(@Nos)
			
			Select @digitsum = 0

			While   @Cnt > 0
		
				Begin
					Select @digitsum = @digitsum + Convert(bigint, Left(@Nos, 1))
					Select @Cnt = @Cnt - 1
					Select @Nos = Right(@Nos, @Cnt)			
					Select @SumOfNo = @digitsum
				End	
	
		End	

	

	Return isnull(@SumOfNo, 0)


END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_Encrypt]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO




CREATE FUNCTION [dbo].[fn_Encrypt] (@InText varchar(1000), @CurrentDate datetime)  

RETURNS varchar(1000) 

AS  

BEGIN 

	Declare @TextLength		int,
		@FirstLock 		char(2),
		@Count			int,
		@ShiftValue		char(2),
		@AsciiValue		int,
		@EncryptChar		char(1),
		@EncryptedText	varchar(1000)

	Select @ShiftValue=Convert(int, day(@CurrentDate))
    	Select @TextLength = Len(@InText)


	Select @EncryptedText = ''

	Select @Count = 1

	While @Count <=  @TextLength

		Begin
        			Select @AsciiValue = Ascii(Substring(@InText, @Count, 1))
        			Select @EncryptChar = Char(@AsciiValue + @ShiftValue + @Count)
        			Select @EncryptedText = @EncryptedText + @EncryptChar

   			Select @Count = @Count + 1
		End

    	If (@ShiftValue < 10)
        		Select @FirstLock = Char(Ascii('0') + 110)-- + Char(Ascii(Convert(varchar(2), @ShiftValue)) + 110)
    	Else
        		Select @FirstLock = Char(Ascii(Substring(Convert(varchar(2), @ShiftValue), 1, 1)) + 110) + Char(Ascii(Substring(Convert(varchar(2), @ShiftValue), 2, 1)) + 110)


    	Select @EncryptedText = @EncryptedText + @FirstLock

	Return (@EncryptedText)


END


GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_BankBook_Balance]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[fn_FAS_BankBook_Balance] (@BankCode int, @RptDate datetime, @CmCode varchar(12))  

RETURNS money 

AS  

BEGIN 
	Declare @BankAmt money

	Select @BankAmt=Sum(Amount)
	from
		(Select isNull(Sum(Amount), 0) as Amount 		--OB Amount
		 from 
			( Select (Sum(DrAmt)-Sum(CrAmt)) as Amount  
			  from FAS_Voucher_Vw 
			  where DocDate <= @RptDate and LrCode=@BankCode and CompCode=@CmCode
		
			  Union All 
			
			  Select OpBalance from FAS_LedgerMaster where LedgerCode=@BankCode and CompCode=@CmCode
		        	) a
	
		 Union All
		
	
		 Select isNull(Sum(Amount) , 0) as Amount		--   (+Issued - Receipt) Amount
		 from 
			( Select Case b.DrAmt When 0 Then  isNull(b.CrAmt, 0) Else isNull(-b.DrAmt, 0) End as Amount
			  from FAS_VoucherTranPay a Join FAS_Voucher_Vw b 
			  On a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.VrRoute=b.VrRoute and a.CompCode=b.CompCode 
			  Where b.LrCode=@BankCode and a.CompCode=@CmCode
			    and (a.BankDate is Null or a.BankDate > @RptDate) and b.DocDate <= @RptDate
			 
			  Union All
			
			  Select Case isReceipt When 0 then ChqAmount Else -ChqAmount End as Amount
			  from FAS_VoucherTranPay
			  Where DocNo=0 and VrRoute = 'OpBal-'+Convert(varchar, @BankCode) 
					and CompCode=@CmCode
			   and (BankDate is Null or BankDate > @RptDate)
			 ) b 
		) x

	RETURN (@BankAmt)
END

GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_BankBook_Value]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION  [dbo].[fn_FAS_BankBook_Value] (@BankCode int, @RptDate datetime, @PrStart datetime, @PrEnd datetime, @CmCode varchar(12))    
RETURNS money   
  
 AS    
  
BEGIN   
  
 Declare @CB money --, @CBDate datetime  
   
 --Select @CBDate = DateAdd(day, -1, @RptDate)  
  
 Select @CB=Sum(OpBalance)   
 from  
  (Select dbo.fn_FAS_LedgerOpBalance(@BankCode, @RptDate, '0', @CmCode) as Opbalance  
  
  Union All  
  
  Select Sum(CrAmt-DrAmt)   
  from FAS_VoucherTranSub a Left Outer Join FAS_VoucherTranPay b  
  On a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.VrRoute=b.VrRoute and a.CompCode=b.CompCode  
  Where a.LrCode=@BankCode   
   and a.DocDate < @RptDate  --and (a.DocDate between @PrStart and @PrEnd)    
   and (b.BankDate is Null or b.BankDate >= @RptDate)    
   and a.CompCode=@CmCode  
     
  
  Union All  
   
  Select Case When isReceipt='1' then -ChqAmount Else ChqAmount End  
  from FAS_VoucherTranPay  
  Where VrRoute='OpBal-' + Convert(varchar, @BankCode) --and DocDate=@RptDate   
   and (BankDate is Null or BankDate > @RptDate) and CompCode=@CmCode  
  ) x  
  
  
 return (@CB)  
  
END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_Budget_LedgerOpBalance]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_FAS_Budget_LedgerOpBalance] (@LrCode int, @OBDate datetime, @ActualsCutOff datetime, @isRealisedAmtOnly bit, @CmCode varchar(12))  

RETURNS money
		
AS  

BEGIN 

	Declare @OpBal as money		
	
	Declare @CBDate as datetime
	Select @CBDate = DateAdd(day, -1, @OBDate)
	
	Declare @isPLEntry bit
	
	If Exists (Select * from FAS_Budget_LedgerMaster a Left Outer Join FAS_GroupMaster b On a.LrRoute=b.GrRoute+b.GrWidth and a.CompCode=b.CompCode
					Where a.LedgerCode=@LrCode and a.CompCode=@CmCode and (a.LedgerCode=2 OR b.Master in ('00', '01'))
				)
		SET @isPLEntry='0'
	Else
		SET @isPLEntry='1'

	If @isPLEntry='1'
		SET @OpBal=0
	Else
		Select @OpBal=a.OpBalance+isNull(b.TranAmt, 0) 
		from FAS_Budget_LedgerMaster a
		Left Outer Join		(Select LrCode, Sum(Case When @isRealisedAmtOnly='1' and b.DocNo is Not Null and b.BankDate is Null Then 0 Else DrAmt-CrAmt End) as TranAmt 
							from dbo.fn_FAS_Budget_VoucherTranSub_UpTo_Lr(@LrCode, @CBDate, @ActualsCutOff, @CmCode) a Left Outer Join FAS_VoucherTranPay b
							On a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.VrRoute=b.VrRoute and a.CompCode=b.CompCode
							Where a.DocDate < @OBDate and a.CompCode=@CmCode
							Group By LrCode) b
		On a.LedgerCode=b.LrCode							
		where a.LedgerCode=@LrCode and a.CompCode=@CmCode

	Return @OpBal

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_Budget_Profit_Loss_OB]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_FAS_Budget_Profit_Loss_OB] (@OBDate datetime, @ActualsCutOff datetime, @CmCode varchar(12))  

RETURNS money

 AS  

BEGIN 

		Declare @Start datetime
		Select @Start=dbo.fn_FAS_Finyear_Start(@CmCode)

		Declare @End datetime
		Select @End=DateAdd(day, -1, @OBDate)

		Declare @Amt as money

		--Default Opening Balance of Profit & Loss A/c Ledger
		Select @Amt=OpBalance from FAS_LedgerMaster Where LedgerCode=2 and CompCode=@CmCode

		If @Start < @OBDate
			Begin
				Select @Amt=@Amt+isNull(Sum(Amt1), 0)
				from
					(Select Sum(ClosingBalance) as Amt1
					from dbo.fn_FAS_Budget_GroupClosingBalance_All(@Start, @End, @ActualsCutOff, '1',  @CmCode)
					where Path in ('0E', '0D', '05', '06', '07', '08') and Type='Group'
					
					Union All
			
					Select  Sum(OpeningStock)
					from dbo.fn_FAS_Budget_StockValue (@Start, @End, @ActualsCutOff, @CmCode)
			
					Union All
			
					Select -Sum(ClosingStock)
					from dbo.fn_FAS_Budget_StockValue (@Start, @End, @ActualsCutOff, @CmCode)

					Union All

					Select Sum(DrAmt-CrAmt) 
					From dbo.fn_FAS_Budget_VoucherTranSub_UpTo_Lr(2, @End, @ActualsCutOff, @CmCode) 
			     	) M

			End

	Return (@Amt)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_Budget_Profit_Loss_OB_SCH]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_FAS_Budget_Profit_Loss_OB_SCH] (@OBDate datetime, @ActualsCutOff datetime, @SuppliedOpeningStock money, @SuppliedClosingStock money, @CmCode varchar(12))  

RETURNS money

 AS  

BEGIN 

		Declare @Start datetime
		Select @Start=dbo.fn_FAS_Finyear_Start(@CmCode)

		Declare @End datetime
		Select @End=DateAdd(day, -1, @OBDate)

		Declare @Amt as money

		--Default Opening Balance of Profit & Loss A/c Ledger
		Select @Amt=OpBalance from FAS_Budget_LedgerMaster Where LedgerCode=2 and CompCode=@CmCode

		If @Start < @OBDate
			Begin
				Select @Amt=@Amt+isNull(Sum(Amt1), 0)
				from
					(Select Sum(ClosingBalance) as Amt1
					from dbo.fn_FAS_Budget_GroupClosingBalance_All(@Start, @End, @ActualsCutOff, '1',  @CmCode)
					where Path in ('0E', '0D', '05', '06', '07', '08') and Type='Group'
					
					Union All
			
					Select Case When @SuppliedOpeningStock>=0 Then @SuppliedOpeningStock 
						Else (Select  Sum(OpeningStock)	from dbo.fn_FAS_Budget_StockValue (@Start, @End, @ActualsCutOff, @CmCode)) End
			
					Union All
					
					Select Case When @SuppliedClosingStock>=0 Then -@SuppliedClosingStock 
								Else (Select -Sum(ClosingStock) from dbo.fn_FAS_Budget_StockValue (@Start, @End, @ActualsCutOff, @CmCode)) End

					Union All

					Select Sum(DrAmt-CrAmt) 
					from dbo.fn_FAS_Budget_VoucherTranSub_UpTo_Lr(2, @End, @ActualsCutOff, @CmCode)
			     		) M

			End

	Return (@Amt)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_Budget_Profit_Loss_Period]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_FAS_Budget_Profit_Loss_Period] (@Start datetime, @End datetime, @ActualsCutOff datetime, @CmCode varchar(12))  

RETURNS money

AS  

BEGIN 

  	Declare @Amt as money

		  Select @Amt=isNull(Sum(Amt1), 0)
		  from
		   (Select Sum(ClosingBalance) as Amt1
		   from dbo.fn_FAS_Budget_Rpt_TrialBalance('S', @Start, @End, @ActualsCutOff, @CmCode)
		   where Path in ('0E', '0D', '05', '06', '07', '08') and Type='Group'
		
		/*   Select Sum(ClosingBalance) as Amt1
		   from dbo.fn_FAS_GroupClosingBalance_All(@Start, @End, '1',  @CmCode)
		   where Path in ('0E', '0D', '05', '06', '07', '08') and Type='Group'
		*/   
		   Union All
		 
		   Select  Sum(OpeningStock)
		   from dbo.fn_FAS_Budget_StockValue (@Start, @End, @ActualsCutOff, @CmCode)
		 
		   Union All
		 
		   Select -Sum(ClosingStock)
		   from dbo.fn_FAS_Budget_StockValue (@Start, @End, @ActualsCutOff, @CmCode)
		
		        ) M
		
	Return (@Amt)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_Budget_Profit_Loss_Period_SCH]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_FAS_Budget_Profit_Loss_Period_SCH] (@Start datetime, @End datetime, @ActualsCutOff datetime, @SuppliedOpeningStock money,  @SuppliedClosingStock money, @CmCode varchar(12))  

RETURNS money

AS  

BEGIN 

  	Declare @Amt as money

		  Select @Amt=isNull(Sum(Amt1), 0)
		  from
		   (Select Sum(ClosingBalance) as Amt1
		   from dbo.fn_FAS_Rpt_Budget_TrialBalance('S', @Start, @End, @ActualsCutOff, @CmCode)
		   where Path in ('0E', '0D', '05', '06', '07', '08') and Type='Group'
		
		   Union All
		 
		   Select Case When @SuppliedOpeningStock>=0 Then @SuppliedOpeningStock 
				Else (Select  Sum(OpeningStock) from dbo.fn_FAS_Budget_StockValue (@Start, @End, @ActualsCutOff, @CmCode)) End
		 
		   Union All
		 
			Select Case When @SuppliedClosingStock>=0 Then -@SuppliedClosingStock 
				Else (Select -Sum(ClosingStock) from dbo.fn_FAS_Budget_StockValue (@Start, @End, @ActualsCutOff, @CmCode)) End
		
		        ) M
		
	Return (@Amt)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_Budget_ReturnVoucherType]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_FAS_Budget_ReturnVoucherType](@DrCode varchar(6), @CrCode varchar(6), @CmCode varchar(12))

RETURNS varchar(4)

AS

BEGIN
	Declare @DrLrType int	/* 1-Party; 2-Bank/Cash; 3-Sales; 4-Purchase; 5-Non Bank/Cash */
	Declare @CrLrType int	/* 1-Party; 2-Bank/Cash; 3-Sales; 4-Purchase; 5-Non Bank/Cash */
	Declare @VrRoute varchar(2)
	Declare @LrRoute varchar(4)
	
	/* To Find Debit Ledger Type */
	Select @LrRoute = Left(LrRoute, 4) 
	From FAS_Budget_LedgerMaster
	Where LedgerCode=Substring(@DrCode, 2, Len(@DrCode)) and CompCode=Left(@CmCode, 4)
	
	If (@LrRoute in ('0206', '0303'))
		Select @DrLrType=1
	Else
	If (@LrRoute in ('0202', '0201', '0B01'))
		Select @DrLrType=2
	Else
	If (@LrRoute = '0E')
		Select @DrLrType=3
	Else
	If (@LrRoute = '0D')
		Select @DrLrType=4
	Else
		Select @DrLrType=5
	
	/* To Find Credit Ledger Type */
	Select @LrRoute = Left(LrRoute, 4) 
	From FAS_Budget_LedgerMaster
	Where LedgerCode=SubString(@CrCode, 2, Len(@CrCode)) and CompCode=Left(@CmCode, 4)
	
	If (@LrRoute in ('0206', '0303'))
		Select @CrLrType=1
	Else
	If (@LrRoute in ('0202', '0201', '0B01'))
		Select @CrLrType=2
	Else
	If (@LrRoute = '0E')
		Select @CrLrType=3
	Else
	If (@LrRoute = '0D')
		Select @CrLrType=4
	Else
		Select @CrLrType=5
		
	
	If (@DrLrType=2)	/* Bank/Cash Debited */
		If (@CrLrType=2)		/* Bank/Cash Credited */
			Select @VrRoute='01'	/* Contra Voucher*/
		Else
			Select @VrRoute='0F'	/* Receipts Voucher */
	Else
	If (@CrLrType=2)	/* Bank/Cash Credited */
		If (@DrLrType=2)		/* Bank/Cash Credited */
			Select @VrRoute='01'	/* Contra Voucher*/
		Else
			Select @VrRoute='09'	/* Payment Voucher */
	Else
	If (@DrLrType=1)	/* Party Debited */
		If (@CrLrType=2)		/* Bank/Cash Credited */
			Select @VrRoute='09'	/* Payment Voucher */
		Else
		If (@CrLrType=3)		/* Sales A/c Credited */
			Select @VrRoute='0K'	/* Sales Voucher */
		Else
			Select @VrRoute='06'	/* Journal Voucher */
	Else
	If (@DrLrType=4)	/* Purchase A/c Debited */
		If (@CrLrType=2)	/* Party Credited */	
			Select @VrRoute='0B'	/* Purchase Voucher */
		Else
			Select @VrRoute='06'	/* Journal Voucher */
	Else
		Select @VrRoute='06'		/* Journal Voucher */
		

	RETURN('VV'+@VrRoute)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_DeleteGroupCheck_Tvw]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[fn_FAS_DeleteGroupCheck_Tvw] (@GrRoute varchar(44), @DeleteIt bit, @CmCode varchar(12))  

RETURNS varchar(500)

 AS  

BEGIN 

	Declare @Msg varchar(500)

	If Exists (Select * From FAS_VoucherTranSub a Where  Exists (Select LedgerCode From FAS_LedgerMaster b
						Where Left(LrRoute, Len(@GrRoute)) = @GrRoute and CompCode = @CmCode and b.LedgerCode = a.LrCode)
		and CompCode = @CmCode)

		Select @Msg =  'Vouchers Found for the Ledgers of this Selected Group'

	Else

		If Exists (Select * From FAS_Billdetails a Where Exists (Select LedgerCode From FAS_LedgerMaster b
						Where Left(LrRoute, Len(@GrRoute)) = @GrRoute and CompCode = @CmCode and b.LedgerCode = a.LrCode)
			and CompCode = @CmCode and VrRoute = 'OpBal')

			Select @Msg =  'Opening Pending Bills Found for the Ledgers of this Selected Group'

		Else
			If Exists (Select * From FAS_VoucherTranPay a Where Exists (Select LedgerCode From FAS_LedgerMaster b
						Where Left(LrRoute, Len(@GrRoute)) = @GrRoute and CompCode = @CmCode and a.VrRoute='OpBal-'+b.LedgerCode)
				and VrRoute like 'OpBal-%' and CompCode = @CmCode)

				Select @Msg =  'Opening Non Realised Cheques Found for the Ledgers of this Selected Group'

			Else
				If @DeleteIt=Convert(bit, '1')
					Begin
						Execute FAS_DeleteEntireGroup @GrRoute, @CmCode
						
						Select @Msg=''
					End

	Return (@Msg)


END










GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_Finyear_Start]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[fn_FAS_Finyear_Start] (@CmCode varchar(12))  

RETURNS datetime

 AS  

BEGIN 	

	Declare @Start datetime, @Result datetime

	Select @Start=Min(DocDate)  from FAS_VoucherTranSub Where CompCode=@CmCode

	If @Start is Not Null
		Begin
			If Month(@Start) Between 4 and 12
				Select @Result = Convert(varchar, Year(@Start)) + '-04-01'
			Else
				Select @Result =Convert(varchar,  Year(@Start) - 1) + '-04-01'
		End

	Return (@Result)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_Get_ExpenseBudget_SalesPlan]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_FAS_Get_ExpenseBudget_SalesPlan] (@AsOnDate datetime, @RptDate datetime, @PlanFor int, @StartDate datetime, @EndDate datetime, @BranchID varchar(3), @CompID varchar(3), @Finyear varchar(15))
/* @PlanFor = Year-1; Month-2; Week-3; Day-4; */

RETURNS money


 AS  

BEGIN 
	
	Declare @SalesPlanValue as money

	Select @SalesPlanValue=Round(sum(SalesPlanValue),2) 
	From
		(
			Select ItemCode, (PlanQty*Rate) SalesPlanValue 
			From 
			(
				Select ItemCode, PlanQty, dbo.fn_Sal_Item_Rate(a.ItemCode, 0, '', 0, Null, @BranchID, @CompID, @Finyear) Rate
				From fn_SAL_SalesPlan(@AsOnDate, @RptDate,  @BranchID, @CompID, @Finyear) a
				Where PlanQty > 0 and PlanFor = @PlanFor and PlanPeriod between @StartDate and @EndDate
			)x
		)y

	Return (@SalesPlanValue)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_Get_Width]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[fn_FAS_Get_Width] (@Input nvarchar(500), @NumToAdd tinyint, @WidthLen tinyint)  
RETURNS nvarchar(500) 

AS 

BEGIN 

	Declare @IpVal nvarchar(500)
  	Declare @Result nvarchar(500)

  	Select @IpVal = Convert(int, dbo.fn_BaseConvert(@Input, 36, 10)) + @NumToAdd
  	Select @IpVal = dbo.fn_BaseConvert(@IpVal, 10, 36)
  	Select @Result = Right((Replicate('0', @WidthLen) +  @IpVal), @WidthLen)

	Return @Result

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_GetMaxVoucherNo]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[fn_FAS_GetMaxVoucherNo] (@VrDate datetime, @VrRoute varchar(42), @VrNo as int, @isMemorandum bit, @CmCode varchar(12))  

RETURNS int 

AS  


BEGIN 

	Declare @Restart varchar(15), @isAutoNo bit

	Select @Restart=RestartAt, @isAutoNo=isAutoNumber From FAS_VoucherMaster Where VrRoute+VrWidth=@VrRoute and CompCode=@CmCode


	Declare @CurPrStart datetime

	Declare @MaxNo int


	If @Restart='FinYearly'
		Begin

			If Month(@VrDate) <= 3 
				Select @CurPrStart = Convert(datetime, Convert(varchar, Year(@VrDate)-1)+'04'+'01', 112)
			Else
				Select @CurPrStart = Convert(datetime, Convert(varchar, Year(@VrDate))+'04'+'01', 112)



			If @isMemorandum='1'
				Select @MaxNo=Max(DocNo) from FAS_MEM_VoucherTranMast
				Where CompCode=@CmCode and (DocDate Between @CurPrStart and @VrDate) and DocNo<>@VrNo
			Else
				Select @MaxNo=Max(DocNo) from FAS_VoucherTranMast
				Where VrRoute=@VrRoute and CompCode=@CmCode and (DocDate Between @CurPrStart and @VrDate) and DocNo<>@VrNo

		End

	Else
	If @Restart='Monthly'
		Begin
			Select @CurPrStart=Convert(datetime, Convert(varchar(7), @VrDate, 120)+'-01', 120)
	
			If @isMemorandum='1'
				Select @MaxNo=Max(DocNo) from FAS_MEM_VoucherTranMast
				Where CompCode=@CmCode and DocDate Between @CurPrStart and @VrDate and DocNo<>@VrNo
			Else
				Select @MaxNo=Max(DocNo) from FAS_VoucherTranMast
				Where VrRoute=@VrRoute and CompCode=@CmCode and DocDate Between @CurPrStart and @VrDate and DocNo<>@VrNo
		End
	Else
	If @Restart='Weekly'
		Begin			
			If @isMemorandum='1'
				Select @MaxNo=Max(DocNo) from FAS_MEM_VoucherTranMast
				Where CompCode=@CmCode and Datepart(wk, DocDate) = Datepart(wk, @VrDate) and DocNo<>@VrNo
			Else
				Select @MaxNo=Max(DocNo) from FAS_VoucherTranMast
				Where VrRoute=@VrRoute and CompCode=@CmCode and Datepart(wk, DocDate) = Datepart(wk, @VrDate) and DocNo<>@VrNo
		End
	Else
	If @Restart='Daily'
		Begin
			If @isMemorandum='1'
				Select @MaxNo=Max(DocNo) from FAS_MEM_VoucherTranMast
				Where CompCode=@CmCode and DocDate = @VrDate and DocNo<>@VrNo
			Else
				Select @MaxNo=Max(DocNo) from FAS_VoucherTranMast
				Where VrRoute=@VrRoute and CompCode=@CmCode and DocDate = @VrDate and DocNo<>@VrNo
		End
	Else		--Never
		Begin
			If @isMemorandum='1'
				Select @MaxNo=Max(DocNo) from FAS_MEM_VoucherTranMast
				Where CompCode=@CmCode and DocNo<>@VrNo
			Else
				Select @MaxNo=Max(DocNo) from FAS_VoucherTranMast
				Where VrRoute=@VrRoute and CompCode=@CmCode and DocNo<>@VrNo
		End



	RETURN isNull(@MaxNo, 0)


END

GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_GrossProfit_Loss_Period]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[fn_FAS_GrossProfit_Loss_Period] (@Start datetime, @End datetime, @CmCode varchar(12))  

RETURNS money

 AS  

BEGIN 
		Declare @Amt as money

		Select @Amt=isNull(Sum(Amt1), 0)
		from
			(Select Sum(ClosingBalance) as Amt1
			from dbo.fn_FAS_GroupClosingBalance_All(@Start, @End, '1',  @CmCode)
			where Path in ('0E', '0D', '05', '06') and Type='Group'
			
			Union All
	
			Select  Sum(OpeningStock)
			from dbo.fn_FAS_StockValue (@Start, @End, @CmCode)
	
			Union All
	
			Select -Sum(ClosingStock)
			from dbo.fn_FAS_StockValue (@Start, @End, @CmCode)

	     		) M

	Return (@Amt)

END













GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_GroupExist_Tvw]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[fn_FAS_GroupExist_Tvw] (@GrRoute Varchar(46), @CompCode Varchar(12))  

RETURNS int

 AS  

BEGIN 

	Declare @Result bit	

	If Exists (Select * From FAS_VoucherTranSub a Where  Exists (Select LedgerCode From FAS_LedgerMaster b
						Where Left(LrRoute, Len(@GrRoute)) = @GrRoute and CompCode = @CompCode and b.LedgerCode = a.LrCode)
		and CompCode = @CompCode)

		--if it Return '1' the Ledger of the Route have Transaction in VoucherTranSub
		Set @Result =  1 

	Else

		If Exists (Select * From FAS_Billdetails a Where Exists (Select LedgerCode From FAS_LedgerMaster b
						Where Left(LrRoute, Len(@GrRoute)) = @GrRoute and CompCode = @CompCode and b.LedgerCode = a.LrCode)
			and CompCode = @CompCode and VrRoute = 'OpBal')

			 --if it Return '1' the Ledger of the Route have Transaction in Billdetails
			Set @Result =  1
		Else
			If Exists (Select * From FAS_VoucherTranPay a Where Exists (Select LedgerCode From FAS_LedgerMaster b
						Where Left(LrRoute, Len(@GrRoute)) = @GrRoute and CompCode = @CompCode)
				and CompCode = @CompCode and VrRoute = 'OpBal-%' )

				--if it Return '1' the Ledger of the Route have Transaction in VoucherTranPay
				Set @Result =  1 

			Else

				--if it Return '0' the Ledger of the Route have No Transaction
				Set @Result =  0  


	Return (@Result)


END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_LedgerExist_Tvw]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[fn_FAS_LedgerExist_Tvw] (@LrCode int, @CompCode Varchar(12))  

RETURNS int

 AS  

BEGIN 
	
	Declare @Result int
	
	If Exists (Select * From FAS_VoucherTranSub a Where  CompCode = @CompCode and a.LrCode = @LrCode)		

		Set @Result =  1  -- If it Return '1'  this Ledger have Transaction in VoucherTranSub 

	Else

		If Exists (Select * From FAS_Billdetails a Where CompCode = @CompCode and a.LrCode = @LrCode and VrRoute = 'OpBal')			

			Set @Result = 1 -- If it Return '1'  this Ledger have Transaction in Billdetails
		Else
			If Exists (Select * From FAS_VoucherTranPay a Where CompCode = @CompCode and  VrRoute = 'OpBal-' + @LrCode)
				
				Set @Result =  1 -- If it Return '1'  this Ledger have Transaction in VoucherTranPay
			Else
			
				Set @Result =  0 -- If it Return '0'   Ledger have No Transaction


	Return (@Result)


END









GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_LedgerOpBalance]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_FAS_LedgerOpBalance] (@LrCode int, @OBDate datetime, @isRealisedAmtOnly bit, @CmCode varchar(12))  

RETURNS money
	
		
AS  

BEGIN 

	Declare @OpBal as money		
	
	Declare @isPLEntry bit
	
	If Exists (Select * from FAS_LedgerMaster a Left Outer Join FAS_GroupMaster b On a.LrRoute=b.GrRoute+b.GrWidth and a.CompCode=b.CompCode
					Where a.LedgerCode=@LrCode and a.CompCode=@CmCode and (a.LedgerCode=2 OR b.Master in ('00', '01'))
				)
		SET @isPLEntry='0'
	Else
		SET @isPLEntry='1'

	If @isPLEntry='1'
		SET @OpBal=0
	Else
		Select @OpBal=a.OpBalance+isNull((Select Sum(Case When @isRealisedAmtOnly='1' and b.DocNo is Not Null and b.BankDate is Null Then 0 Else DrAmt-CrAmt End) as TranAmt 
							from FAS_VoucherTranSub a Left Outer Join FAS_VoucherTranPay b
							On a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.VrRoute=b.VrRoute and a.CompCode=b.CompCode
							Where a.LrCode=@LrCode and a.DocDate < @OBDate and a.CompCode=@CmCode), 0) 
		from FAS_LedgerMaster a
		where a.LedgerCode=@LrCode and a.CompCode=@CmCode

	Return @OpBal

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_Profit_Loss_OB]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[fn_FAS_Profit_Loss_OB] (@OBDate datetime, @CmCode varchar(12))  

RETURNS money

 AS  

BEGIN 

		Declare @Start datetime
		Select @Start=dbo.fn_FAS_Finyear_Start(@CmCode)

		Declare @End datetime
		Select @End=DateAdd(day, -1, @OBDate)

		Declare @Amt as money

		--Default Opening Balance of Profit & Loss A/c Ledger
		Select @Amt=OpBalance from FAS_LedgerMaster Where LedgerCode=2 and CompCode=@CmCode

		If @Start < @OBDate
			Begin
				Select @Amt=@Amt+isNull(Sum(Amt1), 0)
				from
					(Select Sum(ClosingBalance) as Amt1
					from dbo.fn_FAS_GroupClosingBalance_All(@Start, @End, '1',  @CmCode)
					where Path in ('0E', '0D', '05', '06', '07', '08') and Type='Group'
					
					Union All
			
					Select  Sum(OpeningStock)
					from dbo.fn_FAS_StockValue (@Start, @End, @CmCode)
			
					Union All
			
					Select -Sum(ClosingStock)
					from dbo.fn_FAS_StockValue (@Start, @End, @CmCode)

					Union All

					Select Sum(DrAmt-CrAmt) 
					from FAS_VoucherTranSub Where LrCode=2 and DocDate < @OBDate and CompCode=@CmCode
			     		) M

			End

	Return (@Amt)

END









GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_Profit_Loss_OB_SCH]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_FAS_Profit_Loss_OB_SCH] (@OBDate datetime, @SuppliedOpeningStock money, @SuppliedClosingStock money, @CmCode varchar(12))  

RETURNS money

 AS  

BEGIN 

		Declare @Start datetime
		Select @Start=dbo.fn_FAS_Finyear_Start(@CmCode)

		Declare @End datetime
		Select @End=DateAdd(day, -1, @OBDate)

		Declare @Amt as money

		--Default Opening Balance of Profit & Loss A/c Ledger
		Select @Amt=OpBalance from FAS_LedgerMaster Where LedgerCode=2 and CompCode=@CmCode

		If @Start < @OBDate
			Begin
				Select @Amt=@Amt+isNull(Sum(Amt1), 0)
				from
					(Select Sum(ClosingBalance) as Amt1
					from dbo.fn_FAS_GroupClosingBalance_All(@Start, @End, '1',  @CmCode)
					where Path in ('0E', '0D', '05', '06', '07', '08') and Type='Group'
					
					Union All
			
					Select Case When @SuppliedOpeningStock>=0 Then @SuppliedOpeningStock 
						Else (Select  Sum(OpeningStock)	from dbo.fn_FAS_StockValue (@Start, @End, @CmCode)) End
			
					Union All
					
					Select Case When @SuppliedClosingStock>=0 Then -@SuppliedClosingStock 
								Else (Select -Sum(ClosingStock) from dbo.fn_FAS_StockValue (@Start, @End, @CmCode)) End

					Union All

					Select Sum(DrAmt-CrAmt) 
					from FAS_VoucherTranSub Where LrCode=2 and DocDate < @OBDate and CompCode=@CmCode
			     		) M

			End

	Return (@Amt)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_Profit_Loss_Period]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_FAS_Profit_Loss_Period] (@Start datetime, @End datetime, @CmCode varchar(12))    
RETURNS money  
AS    
BEGIN   
  
   Declare @Amt as money  
  
    Select @Amt=isNull(Sum(Amt1), 0)  
    from  
     (Select Sum((DrAmt-CrAmt)) as Amt1  
   From FAS_Voucher_Vw a   
   Where Exists (Select *   
       From FAS_LedgerMaster_AuthTbl b  
       Where a.LrCode=b.LedgerCode and a.CompCode=b.CompCode  
        and Left(b.LrRoute, 2) in  ('0E', '0D', '05', '06', '07', '08')  
      )  
       
	and (a.DocDate Between @Start and @End) and a.CompCode=@CmCode  
            
    Union All  
     
     Select  Sum(OpeningStock)  
     from dbo.fn_FAS_StockValue (@Start, @End, @CmCode)  
     
     Union All  
     
     Select -Sum(ClosingStock)  
     from dbo.fn_FAS_StockValue (@Start, @End, @CmCode)  
    
          ) M  
    

 /* Existing Query
	
		Select @Amt=isNull(Sum(Amt1), 0)  
    from  
     (Select Sum(ClosingBalance) as Amt1  
     from dbo.fn_FAS_Rpt_TrialBalance('S', @Start, @End,  @CmCode)  
     where Path in ('0E', '0D', '05', '06', '07', '08') and Type='Group'  
    
  /*   Select Sum(ClosingBalance) as Amt1  
     from dbo.fn_FAS_GroupClosingBalance_All(@Start, @End, '1',  @CmCode)  
     where Path in ('0E', '0D', '05', '06', '07', '08') and Type='Group'  
  */     
     Union All  
     
     Select  Sum(OpeningStock)  
     from dbo.fn_FAS_StockValue (@Start, @End, @CmCode)  
     
     Union All  
     
     Select -Sum(ClosingStock)  
     from dbo.fn_FAS_StockValue (@Start, @End, @CmCode)  
    
          ) M  
		  */
 Return (@Amt)  
  
END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_Profit_Loss_Period_Gross]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_FAS_Profit_Loss_Period_Gross] (@Start datetime, @End datetime, @CmCode varchar(12))    
RETURNS money  
AS    
BEGIN   
  
    Declare @Amt as money  
  
    Select @Amt=isNull(Sum(Amt1), 0)  
    From  
    (  
        Select Sum(DrAmt-CrAmt) as Amt1  
        From FAS_Voucher_Vw a   
        Where Exists   
            (  
                Select *   
                From FAS_LedgerMaster_AuthTbl b  
                Where a.LrCode=b.LedgerCode and a.CompCode=b.CompCode  
                and Left(b.LrRoute, 2) in ('0E', '06', '0D', '05')  
            )    
            and (a.DocDate Between @Start and @End) and a.CompCode=@CmCode  
       
        Union All  
     
        Select  Sum(OpeningStock)-Sum(ClosingStock)  
        from dbo.fn_FAS_StockValue (@Start, @End, @CmCode)  
    ) M  
    
 Return (@Amt)  
  
END  
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_Profit_Loss_Period_SalPurExpenses]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_FAS_Profit_Loss_Period_SalPurExpenses] (@LrRoute Varchar(42), @Start datetime, @End datetime, @CmCode varchar(12))    
RETURNS money  
AS    
BEGIN   
    Declare @Amt as money;
    
    Select @Amt=isNull(Sum(DrAmt-CrAmt), 0)  
    From FAS_Voucher_Vw a   
    Where Exists   
        (  
            Select *   
            From FAS_LedgerMaster_AuthTbl b  
            Where a.LrCode=b.LedgerCode and a.CompCode=b.CompCode  
            and Left(b.LrRoute, 2) in (@LrRoute)  
        )    
        and (a.DocDate Between @Start and @End) and a.CompCode=@CmCode  
  
    RETURN (@Amt)  
  
END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_Profit_Loss_Period_SCH]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_FAS_Profit_Loss_Period_SCH] (@Start datetime, @End datetime, @SuppliedOpeningStock money,  @SuppliedClosingStock money, @CmCode varchar(12))  

RETURNS money

AS  

BEGIN 

  	Declare @Amt as money

		  Select @Amt=isNull(Sum(Amt1), 0)
		  from
		   (Select Sum(ClosingBalance) as Amt1
		   from dbo.fn_FAS_Rpt_TrialBalance('S', @Start, @End,  @CmCode)
		   where Path in ('0E', '0D', '05', '06', '07', '08') and Type='Group'
		
		   Union All
		 
		   Select Case When @SuppliedOpeningStock>=0 Then @SuppliedOpeningStock 
				Else (Select  Sum(OpeningStock) from dbo.fn_FAS_StockValue (@Start, @End, @CmCode)) End
		 
		   Union All
		 
			Select Case When @SuppliedClosingStock>=0 Then -@SuppliedClosingStock 
				Else (Select -Sum(ClosingStock) from dbo.fn_FAS_StockValue (@Start, @End, @CmCode)) End
		
		        ) M
		
	Return (@Amt)

END

GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_ValidateVoucher]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION  [dbo].[fn_FAS_ValidateVoucher](@VrNo int, @VrDate Datetime, @VrRoute Varchar(42), @CmCode varchar(12))  
RETURNS int 
/* Returns 0-Valid Voucher Other '0' means Invalid Error Returns */
AS  


BEGIN 

	declare @InvalidCriteriaNo as int
	
	/* Voucher Number Not generated Properly */
	If @VrNo < 1 
		Select @InvalidCriteriaNo = 1
		

	/* Voucher Route Validation */
	If Not Exists (Select * From FAS_VoucherMaster 
					Where VrRoute+VrWidth=@VrRoute and isActive='1'
						and CompCode=@CmCode
					)
		Select @InvalidCriteriaNo = 1
		

	/* In Tran Master No Data Found for the Given Voucher Details */
	If Not Exists (Select * From FAS_VoucherTranMast Where DocNo=@VrNo and DocDate=@VrDate and VrRoute=@VrRoute and CompCode=@CmCode)
		Select @InvalidCriteriaNo = 1	

	
	/* In Tran Sub No Data Found for the Given Voucher Details */
	If Not Exists (Select * From FAS_VoucherTranSub 
					Where DocDate=@VrDate and DocNo=@VrNo and VrRoute=@VrRoute and CompCode=@CmCode
					)
		Select @InvalidCriteriaNo = 2	
	
	
	/* Active Ledger Code Validation */
	If Exists (Select * From FAS_VoucherTranSub a Left Outer Join FAS_LedgerMaster b
					On a.LrCode=b.LedgerCode and a.CompCode=b.CompCode
				Where b.IsActive='0' 
					and a.CompCode=@CmCode
				)
		Select @InvalidCriteriaNo = 2
	
	
	/* In Tran Sub Negative or Zero Amount Found for the Given Voucher Details */
	If Exists (Select * From FAS_VoucherTranSub 
				Where DocDate=@VrDate and DocNo=@VrNo and VrRoute=@VrRoute 
					and (DrAmt <= 0 OR CrAmt <= 0)
					and CompCode=@CmCode
				)
		Select @InvalidCriteriaNo = 3		


	/* In Tran Sub Debit & Credit Amount not Matched for the Given Voucher Details */
	If Exists (Select DocNo, DocDate, VrRoute, CompCode From FAS_VoucherTranSub 
				Where DocNo=@VrNo and DocDate=@VrDate and VrRoute=@VrRoute 
					and CompCode=@CmCode
					Group By DocNo, DocDate, VrRoute, CompCode
					Having Sum(DrAmt)<>Sum(CrAmt)
				)
		Select @InvalidCriteriaNo = 4		
		
	
	/* In Tran Sub Invalid LedgerCode Found for the Given Voucher Details */	
	If Exists (Select * From FAS_VoucherTranSub 
				Where DocNo=@VrNo and DocDate=@VrDate and VrRoute=@VrRoute 
					and LrCode <= 0
					and CompCode=@CmCode
				)
		Select @InvalidCriteriaNo = 5	


	/* Payment Entries Not Matched for the Given Voucher Details */
	/* If Payment Cheque or Draft Data found in FAS_VoucherTranMast Table */	
	If Exists (Select * From FAS_VoucherTranMast 
				Where DocNo=@VrNo and DocDate=@VrDate and VrRoute=@VrRoute 
					and PayMode in ('Cheque', 'Draft')
					and CompCode=@CmCode
				)
		/* If Data Found FAS_VoucherTranRef Table Then the BillAmount should be the single 'New Ref' Entry and Amount should be matched with Ledger Amount */
		If Not Exists (Select * From FAS_VoucherTranPay
						Where DocNo=@VrNo and DocDate=@VrDate and VrRoute=@VrRoute 
							and CompCode=@CmCode
						)
		Select @InvalidCriteriaNo = 6		
	
	
	/* If Posted From Sales or Purchase Bill Entries Not Matched for the Given Voucher Details */
	/* If Posted From Sales or Purchase Data found in FAS_VoucherTranRef Table */
	If Exists (Select * From FAS_VoucherTranRef 
				Where DocNo=@VrNo and DocDate=@VrDate and VrRoute=@VrRoute 
					and CompCode=@CmCode
				)
		/* If Data Found FAS_VoucherTranRef Table Then the BillAmount should be the single 'New Ref' Entry and Amount should be matched with Ledger Amount */
		If Exists (Select a.DocNo, a.DocDate, a.VrRoute, a.CompCode From FAS_BillDetails a Inner Join FAS_VoucherTranSub b
						On a.DocDate=b.DocDate and a.DocNo=b.DocNo and a.VrRoute=b.VrRoute and a.VTSNo=b.VTSNo and a.LrCode=b.LrCode and a.CompCode=b.CompCode
						Where a.DocNo=@VrNo and a.DocDate=@VrDate and a.VrRoute=@VrRoute 
							and a.Type='New Ref'
							and a.CompCode=@CmCode
						Group By a.DocNo, a.DocDate, a.VrRoute, a.CompCode
						Having Sum(a.Amount)<>Sum(b.DrAmt-b.CrAmt)
					)
		Select @InvalidCriteriaNo = 7		
	
	
	/* Cost Centre Validation */
	If Exists (Select * from
			(Select x.LedgerName, a.LrCode, a.VTSNo, w.CategoryCode, Abs(a.DrAmt-a.CrAmt) as VrAmt,
			isNull(Sum(b.Amount), 0) as CostAmt, w.isMandatory
			from FAS_VoucherTranSub a Left Outer Join FAS_VoucherTranCostCentre b
			On a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.VrRoute=b.VrRoute and a.LrCode=b.LrCode and a.VTSNo=b.VTSNo and a.CompCode=b.CompCode
			Left Outer Join FAS_CostCentre_Vw w  On b.CostCentreCode=w.CostCentreCode and a.CompCode=w.CompCode
			Left Outer Join FAS_LedgerMaster x On a.LrCode=x.LedgerCode and a.CompCode=x.CompCode
			where a.DocNo=@VrNo and a.DocDate=@VrDate and a.VrRoute=@VrRoute and x.isCostCentreApplicable='1' and a.CompCode=@CmCode
			Group by x.LedgerName, a.LrCode, a.VTSNo, w.CategoryCode, Abs(a.DrAmt-a.CrAmt), w.isMandatory) y
			Where y.VrAmt <> y.CostAmt
		)
		Select @InvalidCriteriaNo = 8
	

	return @InvalidCriteriaNo
END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FAS_VoucherValidate_AfterSave]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[fn_FAS_VoucherValidate_AfterSave] (@VrNo int, @VrDate datetime, @VrRoute varchar(42), @PayMode varchar(10) = 'None', @FinStart datetime, @CmCode varchar(12))  
RETURNS varchar(200) 

AS  

BEGIN 
	Declare @RetStr as varchar(200)


	--Proper Voucher No Validation
	If @VrNo < 1 
		Select @RetStr='Voucher No. not generated Properly.'
	Else

	--Voucher Date Validation
	If @VrDate < @FinStart
		Select @RetStr='Voucher Date is less than Financial year Start.'
	Else

	--Debit Credit Equality Validation
	If Exists (Select Sum(CrAmt) as SubCrAmt, Sum(DrAmt) as SubDrAmt, 
			isNull((Select Amount from FAS_VoucherTranMast	Where docNo=a.docNo  and DocDate = a.DocDate and VrRoute=a.VrRoute and CompCode=a.CompCode), 0) as Amount
		From FAS_VoucherTranSub a 
		Where a.DocNo=@VrNo and a.DocDate=@VrDate and a.VrRoute=@VrRoute and a.CompCode=@CmCode 
		Group by a.DocNo, a.DocDate, a.VrRoute, a.CompCode
		Having Sum(CrAmt)<>Sum(DrAmt)
		)

		Select @RetStr='Debit & Credit Amount not Matched'
	Else
	
	--All LrCode properly Updated Validation
	If Exists (Select LrCode from FAS_VoucherTranSub a 
		where a.DocNo=@VrNo and a.DocDate=@VrDate and a.VrRoute=@VrRoute and LrCode<=0 and a.CompCode=@CmCode 
		)

		Select @RetStr='Unable to Save Accounts Name Properly in Detail Table.'
	Else
	
	--Cheque Entry Validation
	If @PayMode = 'Cheque' 
		If Exists (Select Sum(a.ChqAmount) as ChqAmt from FAS_VoucherTranPay a
			Where a.DocNo=@VrNo and a.DocDate=@VrDate and a.VrRoute=@VrRoute and a.CompCode=@CmCode
			Having Sum(ChqAmount)<=0
			)	

		Select @RetStr='Cheque Amount not saved Properly.'
	Else
	
	--Bill Details Validation
	If Exists (Select * from
			(Select a.VTSNo, x.LedgerName, a.LrCode, (a.DrAmt-a.CrAmt) as VrAmt, isNull(Sum(b.Amount), 0) as BillAmt, x.isBillwise
			from FAS_VoucherTranSub a Left Outer Join FAS_BillDetails b
			On a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.VrRoute=b.VrRoute and a.LrCode=b.LrCode and a.VTSNo=b.VTSNo and a.CompCode=b.CompCode
			Left Outer Join FAS_LedgerMaster x On a.LrCode=x.LedgerCode and a.CompCode=x.CompCode
			where a.docno=@VrNo and a.DocDate=@VrDate and a.VrRoute=@VrRoute and x.isBillwise='1' and a.CompCode=@CmCode
			Group by a.VTSNo, x.LedgerName, a.LrCode, (a.DrAmt-a.CrAmt), x.isBillwise) y
			Where y.VrAmt <> y.BillAmt
		)
	

		Select @RetStr='Voucher Amount not matched with Bill Amount.'
	Else
	
	--Cost Centre Validation
	If Exists (Select * from
			(Select x.LedgerName, a.LrCode, a.VTSNo, w.CategoryCode, Abs(a.DrAmt-a.CrAmt) as VrAmt,
			isNull(Sum(b.Amount), 0) as CostAmt, w.isMandatory
			from FAS_VoucherTranSub a Left Outer Join FAS_VoucherTranCostCentre b
			On a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.VrRoute=b.VrRoute and a.LrCode=b.LrCode and a.VTSNo=b.VTSNo and a.CompCode=b.CompCode
			Left Outer Join FAS_CostCentre_Vw w  On b.CostCentreCode=w.CostCentreCode and a.CompCode=w.CompCode
			Left Outer Join FAS_LedgerMaster x On a.LrCode=x.LedgerCode and a.CompCode=x.CompCode
			where a.DocNo=@VrNo and a.DocDate=@VrDate and a.VrRoute=@VrRoute and x.isCostCentreApplicable='1' and a.CompCode=@CmCode
			Group by x.LedgerName, a.LrCode, a.VTSNo, w.CategoryCode, Abs(a.DrAmt-a.CrAmt), w.isMandatory) y
			Where y.VrAmt <> y.CostAmt
		)
	
		Select @RetStr='Voucher Amount not matched with CostCentre Amount.'


	Select @RetStr=isNull(@RetStr, '')
	
	RETURN (@RetStr)
END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_findnumeric]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_findnumeric] (@GivenStr varchar(100))  

RETURNS varchar(10)

AS  

BEGIN 

	Declare @ResultStr as Varchar(100)		
	Declare @lenstrs as int	

	Declare @fstr as varchar(100)
	Declare @tmpStr as varchar(100)

	Select @fStr = @GivenStr
	
	If @fStr <> ''		

		Begin
			Select @tmpStr = @fStr
			Select @lenstrs = Len(@tmpStr)
			
			--Select @digitsum = 0

			While   @lenstrs > 0
		
				Begin
					if  Isnumeric(Left(@tmpStr, 1)) = 1
						 Begin
							Select @ResultStr = @ResultStr + Left(@tmpStr, 1)
							Select @lenstrs = @lenstrs - 1
							Select @tmpStr = Right(@tmpStr, @lenstrs)
							Select @fStr = @ResultStr
						End
					Else
						Select @lenstrs = @lenstrs - 1
						
						
				End	
	
		End	

	

	Return isnull(@fStr, '')


END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_FrontEndExpression]    Script Date: 17/06/2026 11:43:45 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE FUNCTION [dbo].[fn_FrontEndExpression] (@ModuleCode int, @FrmID int, @ExprSeq int, @BranchID varchar(3), @CompID varchar(3))  

RETURNS varchar(1000)

AS  

BEGIN 
	
	DECLARE @Block as varchar(100), @Expr as varchar(1000), @BlockSeq as int
	
	Select @Expr = ''

	DECLARE Expr_Cursor CURSOR FOR

		Select  
			Case d.BlockType 
					When 'ExprSeq' then Case When d.Block < 0 then 'SUB TOTAL' Else '#'+e.ExprGroupName+'#' End
					When 'Param' Then Case When d.isTypePercentage='1' Then 
									Case When d.ValueUsed<>0 Then d.Block + '[' + Convert(varchar, d.ValueUsed) + '%]' 
										Else d.Block+'[%]' 
									End 
								Else 
									Case When d.ValueUsed <>0 Then d.Block + '[' + Convert(varchar, d.ValueUsed) + ']' 
										Else d.Block
									End
							     End
					Else d.Block 
			End as Block,
			d.BlockSeq
		from 
			(Select a.ModuleCode, a.BlockSeq, a.BlockType, Case a.BlockType When 'Param' Then b.ParamShortName 
									When 'ExprSeq' Then c.ExprRoute 
									Else a.Block 
							End as Block, 
				a.ValueUsed, a.isTypePercentage, a.BranchID, a.CompID 
			from COM_FormulaMast_Block_Dtl a Left Outer Join COM_FrmParamMast_Hdr b 
				On a.Block=Convert(varchar, b.ParamID) and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID
			Left Outer Join COM_FormulaMast_Dtl c
				On a.ModuleCode=c.ModuleCode and a.FrmID=c.FrmID and a.Block=Convert(varchar, c.ExprSeq) and a.BranchID=c.BranchID and a.CompID=c.CompID
			Where a.ModuleCode=@ModuleCode and a.FrmID=@FrmID and a.ExprSeq=@ExprSeq and a.BranchID=@BranchID and a.CompID=@CompID
			) d 
			Left Outer Join COM_FormulaExpr_Group e
				On d.Block=e.Route+e.Width and d.ModuleCode=e.ModuleCode and d.BranchID=e.BranchID and d.CompID=e.CompID
		Order By d.BlockSeq


	Open Expr_Cursor

	FETCH NEXT FROM Expr_Cursor into @Block, @BlockSeq
		WHILE @@FETCH_STATUS = 0
			BEGIN
				If (@Block=')')
					Select @Expr = @Expr + @Block
				Else 
					If (Right(@Expr, 1)='(')
						Select @Expr = @Expr + @Block 
					Else Select @Expr = @Expr + ' ' + @Block 
	
				FETCH NEXT FROM Expr_Cursor into @Block, @BlockSeq
			END
	CLOSE Expr_Cursor
	DEALLOCATE Expr_Cursor

	Return @Expr

END

GO

/****** Object:  UserDefinedFunction [dbo].[fn_GetExprName]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[fn_GetExprName] (@ModuleCode int, @FrmID int, @ExprSeq int, @BranchID varchar(3), @CompID varchar(3))  

RETURNS varchar(100)

AS  

BEGIN 

	Declare @ExprName varchar(100)

	Select Top 1 @ExprName=b.ExprGroupName 
	from COM_FormulaMast_Dtl a Left Outer Join COM_FormulaExpr_Group b
		On a.ExprRoute=b.Route+b.Width and a.BranchID=b.BranchID and a.CompID=b.CompID
	Where a.ExprSeq=@ExprSeq and a.FrmID=@FrmID and a.ModuleCode=@ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID
	
	Return @ExprName

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_GrandTotal]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_GrandTotal] (@ModuleCode int, @DocType varchar(3), @DocNo int, @DocDate datetime, @BranchID varchar(3), @CompID varchar(3), @Finyear varchar(20))  

RETURNS float 

AS  

BEGIN 

	Declare @GrandTotal  float

	Select @GrandTotal = Sum(GrandTotalFlag * ResultValue) 
	From COM_FormulaCalc_Hdr
	Where DocType=@DocType and DocNo=@DocNo and DocDate=@DocDate
		and ModuleCode=@ModuleCode and BranchID=@BranchID and CompID=@CompID 
/*		and Finyear=@Finyear		*/

	Return (@GrandTotal)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_ItemNorms]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION  [dbo].[fn_ItemNorms] (@SentItem varchar(20), @RecdItem varchar(20), @BranchID varchar(3), @CompID varchar(3))  
RETURNS float 

 AS  

BEGIN 

	Declare @Norms float

	Select top 1 @Norms= m.Qty 
	from BOM_Structure_Dtl m 
	where m.parentcode=
				(Select top 1 p.StageCode 
				from BOM_Mps_Master p 
				where p.ItemCode =
							
(Select top 1 n.ItemCode 
						from BOM_Mps_Master n
						where n.StageCode=@RecdItem and n.BranchID=@BranchID and n.CompID=@CompID
						group by n.ItemCode, n.RevisionNo
						having n.RevisionNo=Max(n.RevisionNo)) 
					and p.OperationSequence=1  
					and p.BranchID=@BranchID and p.CompID=@CompID
				group by p.StageCode, p.RevisionNo
				having p.RevisionNo=Max(p.RevisionNo))
		and m.ChildCode=@SentItem and m.BranchID=@BranchID and m.CompID=@CompID
	group by m.Qty, m.RevisionNo
	having m.RevisionNo=Max(m.RevisionNo)


	return (@Norms)

END

GO

/****** Object:  UserDefinedFunction [dbo].[fn_LeaveDays]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION  [dbo].[fn_LeaveDays] (@EmployeeCode varchar(10), @FDate datetime, @TDate datetime, @Branchid Varchar(3), @CompID varchar(3))    
RETURNS float   
  
 AS    
  
  
BEGIN   
  
--For Leave Days  
  
 Declare @LD as float  
  
  select @LD = sum(LeaveDays)  from PAY_Attendance_Entry Where EmployeeCode =@EmployeeCode   
  and LeaveDate between @FDate and @TDate and AbsentType = 'A'  and LeaveType not in ('LOP', 'ESL')
  and Branchid = @Branchid and Compid =@CompID   
  
 return (@LD)  
  
END  
  
  
GO

/****** Object:  UserDefinedFunction [dbo].[fn_LOPDays]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION  [dbo].[fn_LOPDays] (@EmployeeCode varchar(10), @FDate datetime,  @TDate datetime, @Branchid Varchar(3), @CompID varchar(3))    
RETURNS float   
  
 AS    
  
  
BEGIN   
  
--For LOP  
  
 Declare @LOP as float  
   
 Select @LOP = Sum(LeaveDays)  from PAY_Attendance_Entry Where EmployeeCode =@EmployeeCode   
 and LeaveDate between @FDate  and @TDate and AbsentType = 'A'  and LeaveType in ('LOP', 'ESL') and Branchid = @Branchid   
 and Compid =@CompID  
   
 return (@LOP)  
  
END
  
  
  
GO

/****** Object:  UserDefinedFunction [dbo].[fn_Norms]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE FUNCTION  [dbo].[fn_Norms] (@FromItem varchar(15), @ToItem varchar(15),@Branchid Varchar(3), @CompID varchar(3))  
RETURNS float 

 AS  

BEGIN 

	Declare @Norms money

	Select @Norms=isNull( (Select Qty from BOM_STRUCTURE_DTL
				where ParentCode = (Select top 1 ItemCode from BOM_MPS_MASTER 
								where StageCode=@ToItem and Branchid = @Branchid and compid=@CompID)
				and ChildCode = @FromItem and Branchid = @Branchid and CompID=@CompID), 1)
	return (@Norms)

END


GO

/****** Object:  UserDefinedFunction [dbo].[fn_OpeningBalance]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION  [dbo].[fn_OpeningBalance] (@LrCode int, @AsOnDate datetime, @TranCmCode varchar(12))  
RETURNS float 

 AS  

BEGIN 

	Declare @CBal float
	Declare @OPBal float
	Declare @Cr float
	Declare @Dr float	

	Select @OPBal = OPbalance  from FAS_LedgerMaster where LedgerCode = @LrCode and CompCode = Left(@TranCmCode, 4)

	Select @Cr = isNull(Sum(CrAmt), 0), @Dr = isNull(Sum(DrAmt), 0) from FAS_Voucher_Vw where LrCode = @LrCode and isCancelled='0' and CompCode = @TranCmCode

	Select @CBal = (@OPBal + @Dr) - @Cr

	return (@CBal)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_OperationCost]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION  [dbo].[fn_OperationCost] (@ItemCode varchar(15), @BranchID varchar(3), @CompID varchar(3))  
RETURNS float 

AS  

BEGIN 

	Declare @Cost money


	Select @Cost = isNull(Sum(c.OperationCost), 0)
	from dbo.fn_BOM_ProductExplosion(@ItemCode, @BranchID, @CompID) a 
		Left Outer Join BOM_MPS_Master b On a.ItemCode=b.StageCode and a.BranchID=b.BranchID and a.CompID=b.CompID
			Left Outer Join (Select * from BOM_ItemwiseMachinewiseCycleTime Where Priority=1 and BranchID=@BranchID and CompID=@CompID) c
				On b.StageCode=c.ItemCode and c.BranchID=b.BranchID and c.CompID=b.CompID

/*
	Select  @Cost=Sum(f.Operationcost)
	from bom_itemwisemachinewisecycletime f Right Outer Join 
	(
		Select StageCode, OperationSequence
		From
			(Select StageCode, Case OperationSequence When 0 Then 999 Else OperationSequence End as OperationSequence
			from Bom_Mps_Master 
			Where  ItemCode in 
						(Select ItemCode 
						from Bom_Mps_Master 
						where StageCode =@ItemCode and Compid =@CompID
						Group by ItemCode, RevisionNo Having RevisionNo = max(RevisionNo)
						)
			and Compid =@CompID
			Group by StageCode, OperationSequence, RevisionNo Having RevisionNo = max(RevisionNo)
			) m
		Where OperationSequence <= (Select distinct Case OperationSequence When 0 Then 999 Else OperationSequence End as Seq from Bom_Mps_Master 
				Where StageCode =@ItemCode and Compid =@CompID
				Group by ItemCode, OperationSequence, RevisionNo Having RevisionNo = max(RevisionNo))
	) n
	On f.ItemCode=n.StageCode

*/



	 Return(@Cost)

End
GO

/****** Object:  UserDefinedFunction [dbo].[fn_PasswordVerify]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO



CREATE FUNCTION  [dbo].[fn_PasswordVerify] (@Userid varchar(10), @Password varchar(10), @CurDate datetime, @Branchid Varchar(3), @CompID varchar(3))  
RETURNS bit

 AS  

BEGIN 
	declare @DPwd varchar(50)
	declare @UPwd varchar(50)
	declare @Rst bit

	select @DPwd = dbo.fn_encrypt(@Password,@CurDate) 

	select @UPwd = Password from Sec_Users where Userid = @Userid and Branchid = @Branchid  and Compid = @Compid
	
	if @DPwd = @UPwd 	
		set @Rst = 1
	else
		set @Rst = 0

	Return @Rst
END


GO

/****** Object:  UserDefinedFunction [dbo].[fn_PAY_ApplicableWeekDayCount]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_PAY_ApplicableWeekDayCount] (@FromDate datetime, @ToDate datetime, @Branchid varchar(3), @Compid varchar(3))  

RETURNS tinyint

AS  

BEGIN 

	  Declare @WHCnt tinyint
	  Declare @WH int
	  Declare @RH int
	  Declare @EffDate datetime
	  Declare @FDate datetime
	  Declare @TDate datetime

	Select  @WHCnt = 0	

	Declare curRec cursor for

	Select WeeklyHoliday, RoutineHoliday, Max(EffDate) EffDate  from PAY_Holiday_Hdr where EffDate < @FromDate and Branchid = @ToDate
	and Branchid = @Branchid and Compid = @Compid Group by WeeklyHoliday, RoutineHoliday
	Union All
	Select WeeklyHoliday, RoutineHoliday, EffDate  from PAY_Holiday_Hdr where (EffDate Between  @FromDate and @ToDate)
	and Branchid = @Branchid and Compid = @Compid Order by EffDate Desc

	open curRec
	
	Fetch next from curRec into @WH,@RH, @EffDate

	if @EffDate <= @FromDate
		Select @FDate = @FromDate
	else
		Select @FDate = @EffDate

	
	Select @TDate = @ToDate


	while @@Fetch_Status=0

		Begin
			
			Select @WHCnt = @WHCnt +DateDiff(ww, @FDate, @TDate) - Case When @WH = Datepart(w, @FDate) Then -1 Else 0 End


			Fetch next from curRec into @WH,@RH, @EffDate


			if @@Fetch_Status=0
		
				begin
		
					select @TDate = DateAdd("d", -1, @FDate)  
		
					if @TDate < @FromDate
						select @TDate = @FromDate
					else
						select @TDate =@TDate
					
					
					if @EffDate > @FromDate
						Select @FDate = @FromDate
					else
						Select @FDate = @EffDate
				end

		End


		Close curRec
    		Deallocate curRec
 
		RETURN @WHCnt

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_PAY_CheckIsLeave]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION  [dbo].[fn_PAY_CheckIsLeave] (@EmpCode varchar(10), @GDate Datetime, @BranchID Varchar(3), @CompID varchar(3))  

RETURNS varchar(20) 


AS  


BEGIN 
	DECLARE @Result varchar(20)
	
	Declare @JoinDate datetime, @RelDate datetime

	Select @JoinDate=DateOfJoining, @RelDate=RelievedDate 
	From PAY_Employee_Master 
	Where EmployeeCode=@EmpCode and BranchID=@BranchID and CompID=@CompID

	IF @JoinDate > @GDate
		SELECT @Result='Before Joining'
		
	ELSE
	IF @RelDate < @GDate
		SELECT @Result='Relieved'

	ELSE
	IF EXISTS (SELECT * FROM PAY_ShiftExceptionDetails
			WHERE EmployeeCode=@EmpCode AND ExcepDate=@GDate AND  Flg='SH' AND Branchid=@BranchID AND Compid=@CompID)
			
		SELECT @Result='Exception Holiday'
		
	ELSE
	IF EXISTS (SELECT * FROM PAY_ShiftExceptionDetails
		WHERE EmployeeCode=@EmpCode AND ExcepDate=@GDate AND  Flg<>'SH' AND Branchid=@BranchID AND Compid=@CompID)
		
		SELECT @Result='Exception WorkDay'
	ELSE
	BEGIN
		DECLARE @WCode int
	
		SELECT @WCode=pagm.WeekGroupCode FROM PAY_AssignedGroup_Master pagm
		WHERE pagm.EmployeeCode=@EmpCode AND pagm.Branchid=@BranchID AND pagm.Compid=@CompID
			AND pagm.EffDate=(SELECT Max(pagm2.EffDate) FROM PAY_AssignedGroup_Master pagm2
				                  WHERE pagm2.EmployeeCode=@EmpCode AND pagm2.Branchid=@BranchID AND pagm2.Compid=@CompID
				                  	AND pagm2.EffDate<=@GDate
						)
			
		DECLARE @WDayCode int, @WNo int
		SELECT @WDayCode = DATEPART(dw, @GDate)
		SELECT @WNo=((Day(@GDate) - 1) / 7) + 1
			
		IF EXISTS (SELECT * FROM PAY_WeekGroupMaster_Dtl pwgmd 
		           		WHERE pwgmd.WeekGroupCode=@WCode AND pwgmd.WeekNo=@WNo AND pwgmd.WHCode=@WDayCode AND pwgmd.OffValue=1
		           			AND pwgmd.Branchid=@BranchID AND pwgmd.Compid=@CompID
				)
		
			SELECT @Result='Weekly Off'
		ELSE
		IF EXISTS (SELECT * FROM PAY_Holiday_Dtl phd
		           WHERE phd.HolidayDate=@GDate AND phd.Branchid=@BranchID AND phd.Compid=@CompID)			
			
			SELECT @Result='Public Holiday'
		ELSE
			SELECT @Result='Working Day'
	END
			


	RETURN (@Result)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_PAY_CompanyWorkingDays]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION  [dbo].[fn_PAY_CompanyWorkingDays] (@YearMonth Varchar(7), @BranchID Varchar(3), @CompID varchar(3))    
  
RETURNS float  
  
AS  
  
Begin  

	Declare @SDate datetime  
	Declare @EDate datetime  
	Declare @WDays float  
	Declare @TotalDays int  
  
	select @SDate = convert(datetime,@YearMonth + '-01',111)  
	select @EDate = DateAdd("d", -1, DateAdd("m", 1, @SDate))  
	select @TotalDays = Day(@EDate)  
   
	select @WDays=0   
	select @WDays = @WDays + count(*) from fn_PAY_ApplicableWeeklyHolidays_Unicorn(@SDate,@EDate,@BranchID,@CompID) 
									--fn_PAY_ApplicableWeeklyHolidays(@SDate,@EDate,@BranchID,@CompID)  
   
	select @WDays = @WDays + count(HolidayDate) from PAY_Holiday_Dtl where PaidLeave = 1   
	and HolidayDate between @SDate and @EDate  
	and HolidayType<>'Exceptional Holiday' and Branchid = @BranchID and Compid = @CompID  
  
 select @WDays = (@TotalDays - @WDays)  
  
 Return (@WDays)  
  
end  
GO

/****** Object:  UserDefinedFunction [dbo].[fn_PAY_NthOccurrenceOfDay]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_PAY_NthOccurrenceOfDay] (@FromDate datetime, @GivenDay tinyint, @NthOccurrence tinyint)  

RETURNS datetime 

AS  

BEGIN 

	Declare @tmpDatePart int
	Declare @Result datetime
	
	Select @tmpDatePart=Datepart(dw, @FromDate)
	Select @Result=DateAdd(d, (Case When @tmpDatePart > @GivenDay Then (@GivenDay + 7)-@tmpDatePart Else @GivenDay-@tmpDatePart End)+((@NthOccurrence-1)*7), @FromDate)

	Return @Result

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_PAYFrontEndExpression]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO




CREATE FUNCTION [dbo].[fn_PAYFrontEndExpression] (@FrmID int, @ExprSeq int, @BranchID varchar(3), @CompID varchar(3))  

RETURNS varchar(1000)

AS  

BEGIN 
	
	DECLARE @Block as varchar(100), @Expr as varchar(1000), @BlockSeq as int
	
	Select @Expr = ''

	DECLARE Expr_Cursor CURSOR FOR

		Select 
			Case d.BlockType 
					When 'ExprSeq' then e.ExprGroupName 
					When 'Param' Then Case When d.isTypePercentage='1' Then 
									Case When d.ValueUsed<>0 Then d.Block + '[' + Convert(varchar, d.ValueUsed) + '%]' 
										Else d.Block+'[%]' 
									End 
								Else 
									Case When d.ValueUsed <>0 Then d.Block + '[' + Convert(varchar, d.ValueUsed) + ']' 
										Else d.Block
									End
							     End
					Else d.Block 
			End as Block,
			d.BlockSeq
		from 
			(Select a.BlockSeq, a.BlockType, Case a.BlockType When 'Param' Then b.ParamShortName 
									When 'ExprSeq' Then c.ExprRoute 
									Else a.Block 
							End as Block, 
				a.ValueUsed, a.isTypePercentage, a.BranchID, a.CompID 
			from PAY_FormulaMast_Block_Dtl a Left Outer Join PAY_FrmParamMast_Hdr b 
				On a.Block=Convert(varchar, b.ParamID) and a.BranchID=b.BranchID and a.CompID=b.CompID
			Left Outer Join PAY_FormulaMast_Dtl c
				On a.FrmID=c.FrmID and a.Block=Convert(varchar, c.ExprSeq) and a.BranchID=c.BranchID and a.CompID=c.CompID
			Where a.FrmID=@FrmID and a.ExprSeq=@ExprSeq and a.BranchID=@BranchID and a.CompID=@CompID
			) d 
			Left Outer Join PAY_FormulaExpr_Group e
				On d.Block=e.Route+e.Width and d.BranchID=e.BranchID and d.CompID=e.CompID
		Order By d.BlockSeq


	Open Expr_Cursor

	FETCH NEXT FROM Expr_Cursor into @Block, @BlockSeq
		WHILE @@FETCH_STATUS = 0
			BEGIN
				If (@Block=')')
					Select @Expr = @Expr + @Block
				Else 
					If (Right(@Expr, 1)='(')
						Select @Expr = @Expr + @Block 
					Else Select @Expr = @Expr + ' ' + @Block 
	
				FETCH NEXT FROM Expr_Cursor into @Block, @BlockSeq
			END
	CLOSE Expr_Cursor
	DEALLOCATE Expr_Cursor

	Return @Expr

END

GO

/****** Object:  UserDefinedFunction [dbo].[fn_PAYFrontEndExpression1]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_PAYFrontEndExpression1] (@FrmID int, @ExprSeq int, @BranchID varchar(3), @CompID varchar(3))      
    
RETURNS varchar(1000)    
    
AS      
    
BEGIN     
     
 DECLARE @Block as varchar(100), @Expr as varchar(1000), @BlockSeq as int    
     
 Select @Expr = ''    
    
 DECLARE Expr_Cursor CURSOR FOR    
    
  Select     
   Case d.BlockType     
     When 'ExprSeq' then e.ExprGroupName     
     When 'Param' Then Case When d.isTypePercentage='1' Then     
         Case When d.ValueUsed<>0 Then d.Block + '[' + Convert(varchar, d.ValueUsed) + '%]'     
          Else d.Block+'[%]'     
         End     
        Else     
         Case When d.ValueUsed <>0 Then d.Block + '[' + Convert(varchar, d.ValueUsed) + ']'     
          Else d.Block    
         End    
            End    
     Else d.Block     
   End as Block,    
   d.BlockSeq    
  from     
   (Select a.BlockSeq, a.BlockType, Case a.BlockType When 'Param' Then b.ParamShortName     
         When 'ExprSeq' Then c.ExprRoute     
         Else a.Block     
       End as Block,     
    a.ValueUsed, a.isTypePercentage, a.BranchID, a.CompID     
   from PAY_FormulaMast_Block_Dtl1 a Left Outer Join PAY_FrmParamMast_Hdr b     
    On a.Block=Convert(varchar, b.ParamID) and a.BranchID=b.BranchID and a.CompID=b.CompID    
   Left Outer Join PAY_FormulaMast_Dtl1 c    
    On a.FrmID=c.FrmID and a.Block=Convert(varchar, c.ExprSeq) and a.BranchID=c.BranchID and a.CompID=c.CompID    
   Where a.FrmID=@FrmID and a.ExprSeq=@ExprSeq and a.BranchID=@BranchID and a.CompID=@CompID    
   ) d     
   Left Outer Join PAY_FormulaExpr_Group e    
    On d.Block=e.Route+e.Width and d.BranchID=e.BranchID and d.CompID=e.CompID    
  Order By d.BlockSeq    
    
    
 Open Expr_Cursor    
    
 FETCH NEXT FROM Expr_Cursor into @Block, @BlockSeq    
  WHILE @@FETCH_STATUS = 0    
   BEGIN    
    If (@Block=')')    
     Select @Expr = @Expr + @Block    
    Else     
     If (Right(@Expr, 1)='(')    
      Select @Expr = @Expr + @Block     
     Else Select @Expr = @Expr + ' ' + @Block     
     
    FETCH NEXT FROM Expr_Cursor into @Block, @BlockSeq    
   END    
 CLOSE Expr_Cursor    
 DEALLOCATE Expr_Cursor    
    
 Return @Expr    
    
END    



GO

/****** Object:  UserDefinedFunction [dbo].[fn_PreviousFinyearProcess_Get_BasedOnGivenFinyear]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_PreviousFinyearProcess_Get_BasedOnGivenFinyear] (@Finyear varchar(20))        
RETURNS varchar(500)      
 AS        
BEGIN     
 /* Get Previous Finyear Based On Given Param Data */    
  
 Declare @PrevFinyear Varchar(20);  
 Declare @StartFin As varchar(10);  
 Declare @EndFin As varchar(10);  
 Declare @TmpStr As varchar(10);  
 Declare @TmpDate As datetime;  
       
 -- Get Previous Start Finyear  
 SET @TmpStr='01-'+Left(@Finyear,7)          
 SET @TmpDate=(DateAdd(Year,-1,Convert(datetime,@TmpStr,105))) ;      
 SET @StartFin=(Case When Len(Convert(varchar(2),Month(@TmpDate)))=1 Then ('0'+Convert(varchar(1),Month(@TmpDate))) Else Convert(varchar(2),Month(@TmpDate)) End) + '-' + Convert(varchar(4),Year(@TmpDate));  
  
 -- Get Previous End Finyear       
 SET @TmpStr='01-'+SubString(@Finyear,9,7)          
 SET @TmpDate=(DateAdd(Year,-1,Convert(datetime,@TmpStr,105))) ;      
 SET @EndFin=(Case When Len(Convert(varchar(2),Month(@TmpDate)))=1 Then ('0'+Convert(varchar(1),Month(@TmpDate))) Else Convert(varchar(2),Month(@TmpDate)) End) + '-' + Convert(varchar(4),Year(@TmpDate));  
 SET @PrevFinyear=(@StartFin + ' ' + @EndFin);   
    
 Return @PrevFinyear    
END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_Pur_AdvanceAmt_Order]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[fn_Pur_AdvanceAmt_Order] (@ItemCode varchar(20),@Price float,@OrderType varchar(3),@OrderNo int,@OrderDate datetime,@BranchID varchar(3), @CompID varchar(3))  
RETURNS float

AS  

BEGIN 

		Declare @AdvanceAmt  as float

		select @AdvanceAmt=e.AdvanceAmt  
		from
		(
			select c.ItemCode, isnull(d.AdvancePerc,0) as AdvancePerc, ((AdvancePerc/100)*@Price) as AdvanceAmt , c.Stage , c.Branchid, c.Compid 
			from
			(
				select 	case when a.challanQty <> 0 and a.Receiptdate is NULL then '110' 
					else
						case when a.AcceptedDate is NULL then '120' 
						else 
							case when a.AcceptedDate is not NULL then '130' else '0' end 
						end
					end as Stage, a.ItemCode, b.OrderType, b.OrderNo, b.OrderDate, b.Branchid, b.Compid 
				from REC_RECEIPTENTRY_Dtl a
					
				left Outer join
				(
					select DcsrType,DcsrNo,DcsrDate,OrderType,OrderNO,OrderDate,Branchid,Compid  from Pur_Dcsr_Hdr
						
					union all
						
					select ScheduleType,ScheduleNo,ScheduleDate,OrderType,OrderNo,Orderdate,Branchid,Compid from Pur_Schedule
						
				) b 
				on b.DcsrType=a.RefType and b.DcsrNo=a.RefNo and b.DcsrDate=a.RefDate and b.Branchid=a.Branchid and b.Compid=a.Compid
				where a.Branchid=@Branchid and a.Compid=@Compid
			) c
			
			left outer join  Pur_Order_Advance d on c.OrderType=d.OrderType 
			and c.OrderNo=d.OrderNo and c.OrderDate=d.OrderDate and c.Stage=d.WorkPerc and d.Branchid=c.Branchid and d.Compid=c.Compid
			
			where c.OrderType=@OrderType and c.OrderNo=@OrderNo and c.OrderDate=@OrderDate and c.Branchid=@Branchid and c.Compid=@Compid
		) e

	Return (@AdvanceAmt)

END





GO

/****** Object:  UserDefinedFunction [dbo].[fn_Pur_Item_Rate]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_Pur_Item_Rate] (@ItemCode varchar(20), @VendorCode int, @OrderType varchar(3), @OrderNo int, @OrderDate datetime,  @AsOnDate datetime, @RptDate datetime, @BranchID varchar(3), @CompID varchar(3), @FinYear varchar(15))  

RETURNS money

AS
	BEGIN 
		
		Declare @Rate as money		
		
		If @VendorCode = 0 and @OrderDate is Null	/* No Vendor & Order Specified */

			Begin
				Select  @Rate=Rate
				From dbo.fn_PUR_PurchaseOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a 
				Inner Join 
					(
						Select Max(b.OrderDate) OrderDate
						From dbo.fn_PUR_PurchaseOrder_Hdr (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) b
						
						Right Outer Join dbo.fn_PUR_PurchaseOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)c
							On b.OrderType=c.OrderType and b.OrderNo=c.OrderNo and b.OrderDate=c.OrderDate
	
						Where c.ItemCode = @ItemCode and b.BranchID=@BranchID and b.CompID=@CompID
					) b
				On a.OrderDate=b.OrderDate
				Where a.ItemCode = @ItemCode and a.BranchID = @BranchID  and a.CompID = @CompID

				If (isNull(@Rate, 0) <= 0)			
					Select  @Rate=Max((Rate/RatePer))
					From dbo.fn_PUR_LaborOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a 
					Inner Join 
						(
							Select Max(b.OrderDate) OrderDate
							From dbo.fn_PUR_LaborOrder_Hdr (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) b
							Right Outer Join dbo.fn_PUR_LaborOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)c
								On b.OrderType=c.OrderType and b.OrderNo=c.OrderNo and b.OrderDate=c.OrderDate
							Where c.RecdItemcode = @ItemCode and b.BranchID=@BranchID and b.CompID=@CompID
						) b
					On a.OrderDate=b.OrderDate
					Where a.RecdItemcode = @ItemCode and a.BranchID = @BranchID  and a.CompID = @CompID
					
				/*Else
				If (isNull(@Rate, 0) <= 0)	
					Select  @Rate=Max(Rate)
					From dbo.fn_PUR_PurchaseOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a Inner Join 
						(Select OrderType, Max(OrderNo) OrderNo, Max(OrderDate) OrderDate
						From dbo.fn_PUR_PurchaseOrder_Hdr (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) b
						Where BranchID=@BranchID and CompID=@CompID
						Group By OrderType
						) b
					On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate
					Where a.ItemCode = @ItemCode  and a.OrderDate >= dateadd(month, -3, @Rptdate) 						
						and a.BranchID = @BranchID  and a.CompID = @CompID	*/
				Else
				If (isNull(@Rate, 0) <= 0)						
					Select  @Rate=Rate
					From dbo.fn_PUR_ReworkOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
					Where RecdItemcode = @ItemCode --and OrderType=@OrderType and OrderDate=@OrderDate and OrderNo=@OrderNo
							and BranchID = @BranchID and CompID = @CompID
				Else
				If (isNull(@Rate, 0) <= 0)					
					Select  @Rate=Rate
					From dbo.fn_PUR_DismantlingOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
					Where RecdItemcode = @ItemCode  --and OrderType=@OrderType and OrderDate=@OrderDate and OrderNo=@OrderNo
						and BranchID = @BranchID and CompID = @CompID
				Else
				If (isNull(@Rate, 0) <= 0)						
					Select  @Rate=Rate
					From Pur_Vendor_ItemMaster
					Where Itemcode = @ItemCode and BranchID = @BranchID and CompID = @CompID
				Else
				If (isNull(@Rate, 0) <= 0)						
					Select  @Rate=Max((Rate/Rateper))
					From Pur_SubCont_ItemMaster
					Where RecdItemcode = @ItemCode and BranchID = @BranchID and CompID = @CompID
					Group By RecdItemcode					
				Else
				If (isNull(@Rate, 0) <= 0)	
					Select @Rate=dbo.fn_WtAvgRate(@ItemCode, Convert(datetime, '01-' + Left(@FinYear, 7), 105), Convert(datetime, '31-' + Right(@FinYear, 7), 105), @BranchID, @CompID, @FinYear)
				Else
				If (isNull(@Rate, 0) <= 0)						
					Select  @Rate=Rate
					From Inv_Material_Master 
					Where ItemCode = @ItemCode and  BranchID = @BranchID and CompID = @CompID
			End
		Else
			If @OrderDate is Null 			/* No Order Specified */
			
				Begin
					Select  @Rate=Rate
					From dbo.fn_PUR_PurchaseOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a Inner Join 
						(Select OrderType, Max(OrderNo) OrderNo, Max(OrderDate) OrderDate
						From dbo.fn_PUR_PurchaseOrder_Hdr (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) b
						Where VendorCode=@VendorCode and BranchID=@BranchID and CompID=@CompID
						Group By OrderType
						) b
					On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate
					Where a.ItemCode = @ItemCode and a.BranchID = @BranchID  and a.CompID = @CompID
	
					If (isNull(@Rate, 0) <= 0)			
						Select  @Rate=Max((Rate/RatePer))
						From dbo.fn_PUR_LaborOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a Inner Join 
							(Select OrderType, Max(OrderNo) OrderNo, Max(OrderDate) OrderDate
							From dbo.fn_PUR_LaborOrder_Hdr (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) b
							Where VendorCode=@VendorCode and BranchID=@BranchID and CompID=@CompID
							Group By OrderType
							) b
						On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate
						Where a.RecdItemcode = @ItemCode and a.BranchID = @BranchID  and a.CompID = @CompID
						
					Else
					If (isNull(@Rate, 0) <= 0)	
						Select  @Rate=Max(Rate)
						From dbo.fn_PUR_PurchaseOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a Inner Join 
							(Select OrderType, Max(OrderNo) OrderNo, Max(OrderDate) OrderDate
							From dbo.fn_PUR_PurchaseOrder_Hdr (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) b
							Where VendorCode=@VendorCode and BranchID=@BranchID and CompID=@CompID
							Group By OrderType
							) b
						On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate
						Where a.ItemCode = @ItemCode  and a.OrderDate >= dateadd(month, -3, @Rptdate) 						
							and a.BranchID = @BranchID  and a.CompID = @CompID	
					Else
					If (isNull(@Rate, 0) <= 0)						
						Select  @Rate=Rate
						From dbo.fn_PUR_ReworkOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
						Where RecdItemcode = @ItemCode --and OrderType=@OrderType and OrderDate=@OrderDate and OrderNo=@OrderNo
								and BranchID = @BranchID and CompID = @CompID
					Else
					If (isNull(@Rate, 0) <= 0)					
						Select  @Rate=Rate
						From dbo.fn_PUR_DismantlingOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
						Where RecdItemcode = @ItemCode  --and OrderType=@OrderType and OrderDate=@OrderDate and OrderNo=@OrderNo
							and BranchID = @BranchID and CompID = @CompID
					Else
					If (isNull(@Rate, 0) <= 0)						
						Select  @Rate=Rate
						From Pur_Vendor_ItemMaster
						Where Itemcode = @ItemCode and VendorCode=@VendorCode and BranchID = @BranchID and CompID = @CompID
					Else
					If (isNull(@Rate, 0) <= 0)						
						Select  @Rate=Max((Rate/Rateper))
						From Pur_SubCont_ItemMaster
						Where RecdItemcode = @ItemCode and SubContCode = @VendorCode and BranchID = @BranchID and CompID = @CompID
						Group By RecdItemcode					
					Else
					If (isNull(@Rate, 0) <= 0)	
						Select @Rate=dbo.fn_WtAvgRate(@ItemCode, Convert(datetime, '01-' + Left(@FinYear, 7), 105), Convert(datetime, '31-' + Right(@FinYear, 7), 105), @BranchID, @CompID, @FinYear)
					Else
					If (isNull(@Rate, 0) <= 0)						
						Select  @Rate=Rate
						From Inv_Material_Master 
						Where ItemCode = @ItemCode and  BranchID = @BranchID and CompID = @CompID
				End

		Else	/* Order specified */

				Begin
					Select @Rate=Rate
							From dbo.fn_PUR_PurchaseOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a
							Where ItemCode = @ItemCode and OrderType =  @OrderType and OrderNo = @OrderNo and OrderDate = @OrderDate
										and BranchID = @BranchID  and CompID = @CompID
					If (isNull(@Rate, 0) <= 0)
						Select  @Rate=Max((Rate/RatePer))
						From dbo.fn_PUR_LaborOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
						Where  RecdItemcode = @ItemCode and OrderType =  @OrderType and OrderNo = @OrderNo and OrderDate = @OrderDate
							and BranchID = @BranchID and CompID = @CompID
						Group By RecdItemcode
					Else
					If (isNull(@Rate, 0) <= 0)	
						Select @Rate=Max(Rate)
						From dbo.fn_PUR_PurchaseOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a
						Where ItemCode = @ItemCode and OrderType =  @OrderType and OrderNo = @OrderNo and OrderDate = @OrderDate
							and a.OrderDate >= dateadd(month, -3, @Rptdate) 
							and a.BranchID = @BranchID  and a.CompID = @CompID
					Else
					If (isNull(@Rate, 0) <= 0)	
						Select  @Rate=Rate 
						From dbo.fn_PUR_ReworkOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
						Where RecdItemcode = @ItemCode and OrderType=@OrderType and OrderDate=@OrderDate and OrderNo=@OrderNo
							and BranchID = @BranchID and CompID = @CompID
					Else
					If (isNull(@Rate, 0) <= 0)				
						Select  @Rate=Rate
						From dbo.fn_PUR_DismantlingOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
						Where RecdItemcode = @ItemCode  and OrderType=@OrderType and OrderDate=@OrderDate and OrderNo=@OrderNo
							and BranchID = @BranchID and CompID = @CompID
					Else
					If (isNull(@Rate, 0) <= 0)	
						Select  @Rate=Rate
						From Pur_Vendor_ItemMaster
						Where Itemcode = @ItemCode and VendorCode = @VendorCode and BranchID = @BranchID and CompID = @CompID
					Else
					If (isNull(@Rate, 0) <= 0)	
						Select  @Rate=Max((Rate/Rateper)) 
						From Pur_SubCont_ItemMaster
						Where RecdItemcode = @ItemCode and SubContCode = @VendorCode and BranchID = @BranchID and CompID = @CompID
						Group By RecdItemcode
					Else
					If (isNull(@Rate, 0) <= 0)	
						Select @Rate=dbo.fn_WtAvgRate(@ItemCode, Convert(datetime, '01-' + Left(@FinYear, 7), 105), Convert(datetime, '31-' + Right(@FinYear, 7), 105), @BranchID, @CompID, @FinYear)
					Else
					If (isNull(@Rate, 0) <= 0)	
						Select  @Rate=Rate
						From Inv_Material_Master 
						Where ItemCode = @ItemCode and  BranchID = @BranchID and CompID = @CompID
				End

		Return (IsNull(@Rate, 0))		
	End

/*
BEGIN 
		
		Declare @Rate as money		
		
		if @OrderDate is Not Null
			Begin
				Select @Rate=Rate
						From dbo.fn_PUR_PurchaseOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a
						Where ItemCode = @ItemCode and OrderType =  @OrderType and OrderNo = @OrderNo and OrderDate = @OrderDate
									and BranchID = @BranchID  and CompID = @CompID
				If (isNull(@Rate, 0) <= 0)
					Select  @Rate=Max((Rate/RatePer))
					From dbo.fn_PUR_LaborOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
					Where  RecdItemcode = @ItemCode and OrderType =  @OrderType and OrderNo = @OrderNo and OrderDate = @OrderDate
						and BranchID = @BranchID and CompID = @CompID
					Group By RecdItemcode
				Else
				If (isNull(@Rate, 0) <= 0)	
					Select @Rate=Max(Rate)
					From dbo.fn_PUR_PurchaseOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a
					Where ItemCode = @ItemCode and OrderType =  @OrderType and OrderNo = @OrderNo and OrderDate = @OrderDate
						and a.OrderDate >= dateadd(month, -3, @Rptdate) 
						and a.BranchID = @BranchID  and a.CompID = @CompID
				Else
				If (isNull(@Rate, 0) <= 0)	
					Select  @Rate=Rate 
					From dbo.fn_PUR_ReworkOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
					Where RecdItemcode = @ItemCode and OrderType=@OrderType and OrderDate=@OrderDate and OrderNo=@OrderNo
						and BranchID = @BranchID and CompID = @CompID
				Else
				If (isNull(@Rate, 0) <= 0)				
					Select  @Rate=Rate
					From dbo.fn_PUR_DismantlingOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
					Where RecdItemcode = @ItemCode  and OrderType=@OrderType and OrderDate=@OrderDate and OrderNo=@OrderNo
						and BranchID = @BranchID and CompID = @CompID
				Else
				If (isNull(@Rate, 0) <= 0)	
					Select  @Rate=Rate
					From Pur_Vendor_ItemMaster
					Where Itemcode = @ItemCode and VendorCode = @VendorCode and BranchID = @BranchID and CompID = @CompID
				Else
				If (isNull(@Rate, 0) <= 0)	
					Select  @Rate=Max((Rate/Rateper)) 
					From Pur_SubCont_ItemMaster
					Where RecdItemcode = @ItemCode and SubContCode = @VendorCode and BranchID = @BranchID and CompID = @CompID
					Group By RecdItemcode
				Else
				If (isNull(@Rate, 0) <= 0)	
					Select @Rate=dbo.fn_WtAvgRate(@ItemCode, Convert(datetime, '01-' + Left(@FinYear, 7), 105), Convert(datetime, '31-' + Right(@FinYear, 7), 105), @BranchID, @CompID, @FinYear)
				Else
				If (isNull(@Rate, 0) <= 0)	
					Select  @Rate=Rate
					From Inv_Material_Master 
					Where ItemCode = @ItemCode and  BranchID = @BranchID and CompID = @CompID
			End

		Else

			Begin
				Select  @Rate=Rate
				From dbo.fn_PUR_PurchaseOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a Inner Join 
					(Select OrderType, Max(OrderNo) OrderNo, Max(OrderDate) OrderDate
					From dbo.fn_PUR_PurchaseOrder_Hdr (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) b
					Where VendorCode=@VendorCode and BranchID=@BranchID and CompID=@CompID
					Group By OrderType
					) b
				On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate
				Where a.ItemCode = @ItemCode and a.BranchID = @BranchID  and a.CompID = @CompID

				If (isNull(@Rate, 0) <= 0)			
					Select  @Rate=Max((Rate/RatePer))
					From dbo.fn_PUR_LaborOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a Inner Join 
						(Select OrderType, Max(OrderNo) OrderNo, Max(OrderDate) OrderDate
						From dbo.fn_PUR_LaborOrder_Hdr (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) b
						Where VendorCode=@VendorCode and BranchID=@BranchID and CompID=@CompID
						Group By OrderType
						) b
					On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate
					Where a.RecdItemcode = @ItemCode and a.BranchID = @BranchID  and a.CompID = @CompID
					
				Else
				If (isNull(@Rate, 0) <= 0)	
					Select  @Rate=Max(Rate)
					From dbo.fn_PUR_PurchaseOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a Inner Join 
						(Select OrderType, Max(OrderNo) OrderNo, Max(OrderDate) OrderDate
						From dbo.fn_PUR_PurchaseOrder_Hdr (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) b
						Where VendorCode=@VendorCode and BranchID=@BranchID and CompID=@CompID
						Group By OrderType
						) b
					On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate
					Where a.ItemCode = @ItemCode  and a.OrderDate >= dateadd(month, -3, @Rptdate) 						
						and a.BranchID = @BranchID  and a.CompID = @CompID	
				Else
				If (isNull(@Rate, 0) <= 0)						
					Select  @Rate=Rate
					From dbo.fn_PUR_ReworkOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
					Where RecdItemcode = @ItemCode --and OrderType=@OrderType and OrderDate=@OrderDate and OrderNo=@OrderNo
							and BranchID = @BranchID and CompID = @CompID
				Else
				If (isNull(@Rate, 0) <= 0)					
					Select  @Rate=Rate
					From dbo.fn_PUR_DismantlingOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
					Where RecdItemcode = @ItemCode  --and OrderType=@OrderType and OrderDate=@OrderDate and OrderNo=@OrderNo
						and BranchID = @BranchID and CompID = @CompID
				Else
				If (isNull(@Rate, 0) <= 0)						
					Select  @Rate=Rate
					From Pur_Vendor_ItemMaster
					Where Itemcode = @ItemCode and VendorCode=@VendorCode and BranchID = @BranchID and CompID = @CompID
				Else
				If (isNull(@Rate, 0) <= 0)						
					Select  @Rate=Max((Rate/Rateper))
					From Pur_SubCont_ItemMaster
					Where RecdItemcode = @ItemCode and SubContCode = @VendorCode and BranchID = @BranchID and CompID = @CompID
					Group By RecdItemcode					
				Else
				If (isNull(@Rate, 0) <= 0)	
					Select @Rate=dbo.fn_WtAvgRate(@ItemCode, Convert(datetime, '01-' + Left(@FinYear, 7), 105), Convert(datetime, '31-' + Right(@FinYear, 7), 105), @BranchID, @CompID, @FinYear)
				Else
				If (isNull(@Rate, 0) <= 0)						
					Select  @Rate=Rate
					From Inv_Material_Master 
					Where ItemCode = @ItemCode and  BranchID = @BranchID and CompID = @CompID
			End

		Return (IsNull(@Rate, 0))		
	End
*/


/*
	BEGIN 
		
		Declare @Rate as money		
		
		if @OrderDate is Not Null
			Begin
				Select top 1 @Rate=Rate
				from
					(Select Rate, 1 as Priority
							From dbo.fn_PUR_PurchaseOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a
							Where ItemCode = @ItemCode and OrderType =  @OrderType and OrderNo = @OrderNo and OrderDate = @OrderDate
										and BranchID = @BranchID  and CompID = @CompID
					
					Union All
		
					Select  Max((Rate/RatePer)) as Rate, 2 as Priority
					From dbo.fn_PUR_LaborOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
					Where  RecdItemcode = @ItemCode and OrderType =  @OrderType and OrderNo = @OrderNo and OrderDate = @OrderDate
						and BranchID = @BranchID and CompID = @CompID
					Group By RecdItemcode

					Union All
	
					Select Max(Rate), 3 as Priority
					From dbo.fn_PUR_PurchaseOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a
					Where ItemCode = @ItemCode and OrderType =  @OrderType and OrderNo = @OrderNo and OrderDate = @OrderDate
						and a.OrderDate >= dateadd(month, -3, @Rptdate) 
						and a.BranchID = @BranchID  and a.CompID = @CompID

					Union All
					
					Select  Rate, 4 as Priority 
					From dbo.fn_PUR_ReworkOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
					Where RecdItemcode = @ItemCode and OrderType=@OrderType and OrderDate=@OrderDate and OrderNo=@OrderNo
							and BranchID = @BranchID and CompID = @CompID
					
					Union All

					Select  Rate, 5 as Priority 


					From dbo.fn_PUR_DismantlingOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
					Where RecdItemcode = @ItemCode  and OrderType=@OrderType and OrderDate=@OrderDate and OrderNo=@OrderNo
						and BranchID = @BranchID and CompID = @CompID
				
				
					Union All
					
					Select  Rate, 6 as Priority
					From Pur_Vendor_ItemMaster
					Where Itemcode = @ItemCode and VendorCode = @VendorCode and BranchID = @BranchID and CompID = @CompID

					Union All 
					
					Select  Max((Rate/Rateper)) Rate, 7 as Priority
					From Pur_SubCont_ItemMaster
					Where RecdItemcode = @ItemCode and SubContCode = @VendorCode and BranchID = @BranchID and CompID = @CompID
					Group By RecdItemcode

					Union All

					Select dbo.fn_WtAvgRate(@ItemCode, Convert(datetime, '01-' + Left(@FinYear, 7), 105), Convert(datetime, '31-' + Right(@FinYear, 7), 105), @BranchID, @CompID, @FinYear), 8 as Priority

					Union All

					Select  Rate, 9 as Priority
					From Inv_Material_Master 
					Where ItemCode = @ItemCode and  BranchID = @BranchID and CompID = @CompID


					)x
				Where isNull(Rate, 0)<>0
				Order by Priority
			End

		else

			Begin
				Select top 1 @Rate=Rate
				from
					(Select  Rate, 1 as Priority
					From dbo.fn_PUR_PurchaseOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a Inner Join 
						(Select OrderType, Max(OrderNo) OrderNo, Max(OrderDate) OrderDate
						From dbo.fn_PUR_PurchaseOrder_Hdr (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) b
						Where VendorCode=@VendorCode and BranchID=@BranchID and CompID=@CompID
						Group By OrderType
						) b
					On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate
					Where a.ItemCode = @ItemCode and a.BranchID = @BranchID  and a.CompID = @CompID

					Union All
		
					Select  Max((Rate/RatePer)) as Rate, 2 as Priority
					From dbo.fn_PUR_LaborOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a Inner Join 
						(Select OrderType, Max(OrderNo) OrderNo, Max(OrderDate) OrderDate
						From dbo.fn_PUR_LaborOrder_Hdr (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) b
						Where VendorCode=@VendorCode and BranchID=@BranchID and CompID=@CompID
						Group By OrderType
						) b
					On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate
					Where a.RecdItemcode = @ItemCode and a.BranchID = @BranchID  and a.CompID = @CompID
					
					Union All 

					Select  Max(Rate), 3 as Priority
					From dbo.fn_PUR_PurchaseOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) a Inner Join 
						(Select OrderType, Max(OrderNo) OrderNo, Max(OrderDate) OrderDate
						From dbo.fn_PUR_PurchaseOrder_Hdr (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear) b
						Where VendorCode=@VendorCode and BranchID=@BranchID and CompID=@CompID
						Group By OrderType
						) b
					On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate
					Where a.ItemCode = @ItemCode  and a.OrderDate >= dateadd(month, -3, @Rptdate) 						
						and a.BranchID = @BranchID  and a.CompID = @CompID
	
					Union All
					
					Select  Rate, 4 as Priority 
					From dbo.fn_PUR_ReworkOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
					Where RecdItemcode = @ItemCode --and OrderType=@OrderType and OrderDate=@OrderDate and OrderNo=@OrderNo
							and BranchID = @BranchID and CompID = @CompID

					Union All
				
					Select  Rate, 5 as Priority 
					From dbo.fn_PUR_DismantlingOrder_Dtl (@AsOnDate, @RptDate, @BranchID, @CompID, @FinYear)
					Where RecdItemcode = @ItemCode  --and OrderType=@OrderType and OrderDate=@OrderDate and OrderNo=@OrderNo
						and BranchID = @BranchID and CompID = @CompID

					Union All
					
					Select  Rate, 6 as Priority
					From Pur_Vendor_ItemMaster
					Where Itemcode = @ItemCode and VendorCode=@VendorCode and BranchID = @BranchID and CompID = @CompID

					Union All 
					
					Select  Max((Rate/Rateper)) Rate, 7 as Priority
					From Pur_SubCont_ItemMaster
					Where RecdItemcode = @ItemCode and SubContCode = @VendorCode and BranchID = @BranchID and CompID = @CompID
					Group By RecdItemcode
					
					Union All

					Select dbo.fn_WtAvgRate(@ItemCode, Convert(datetime, '01-' + Left(@FinYear, 7), 105), Convert(datetime, '31-' + Right(@FinYear, 7), 105), @BranchID, @CompID, @FinYear),8 as Priority

					Union All
					
					Select  Rate, 9 as Priority
					From Inv_Material_Master 
					Where ItemCode = @ItemCode and  BranchID = @BranchID and CompID = @CompID
					)x
				Where isNull(Rate, 0)<>0
				Order by Priority
			End

		Return (Isnull(@Rate, 0))		
	End


*/


GO

/****** Object:  UserDefinedFunction [dbo].[fn_Pur_LOMiscItem_Rate]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_Pur_LOMiscItem_Rate] (@ItemCode varchar(20),@RptDate datetime, @AsOnDate datetime, @FromDate datetime, @ToDate datetime, @BranchID varchar(3), @CompID varchar(3), @FinYear varchar(15))  

RETURNS money
	AS

	BEGIN 
		
	Declare @Rate as money	
	
		Select @Rate=Rate
		From dbo.fn_PUR_PurchaseOrder_Dtl(@AsOnDate, @RptDate, @BranchID, @CompID, @Finyear)
		Where ItemCode = @ItemCode  and BranchID = @BranchID  and CompID = @CompID
		
		if (IsNull(@Rate,0) <=0) 
			Select @Rate=Rate
			From dbo.fn_PUR_LaborOrder_Dtl( @AsOnDate, @RptDate, @BranchID, @CompID, @Finyear)  
			Where RecdItemCode=@ItemCode and BranchID = @BranchID  and CompID = @CompID
		Else
		if (IsNull(@Rate,0) <= 0)
			Select  @Rate=Rate
			From Pur_Vendor_ItemMaster
			Where Itemcode = @ItemCode and BranchID = @BranchID and CompID = @CompID

		Else
		if (IsNull(@Rate,0) <= 0)
			Select  @Rate=Rate
			From PUR_SubCont_ItemMaster
			Where RecdItemcode = @ItemCode and BranchID = @BranchID and CompID = @CompID		

		Else
		if (IsNull(@Rate,0) <= 0)

			Select  @Rate=Rate
			From Inv_Material_Master 
			Where ItemCode = @ItemCode and  BranchID = @BranchID and CompID = @CompID
		
		Return (IsNull(@Rate, 0))	
	
		End
GO

/****** Object:  UserDefinedFunction [dbo].[fn_Pur_PurVsCons_Rate]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_Pur_PurVsCons_Rate] (@ItemCode varchar(20), @AsOnDate datetime, @FromDate datetime, @ToDate datetime, @BranchID varchar(3), @CompID varchar(3), @FinYear varchar(15))  

RETURNS money
	AS

	BEGIN 
		
	Declare @Rate as money	
	
		Select @Rate=Rate
		From PUR_PurchaseOrder_Dtl
		Where ItemCode = @ItemCode  and OrderDate =(Select Max(b.OrderDate) From PUR_PurchaseOrder_Dtl b
								Where b.OrderDate <=@ToDate and  ItemCode = b.ItemCode )and BranchID = @BranchID  and CompID = @CompID
		
		if (IsNull(@Rate,0) <=0) 
			Select  @Rate=Rate
			From Pur_Vendor_ItemMaster
			Where Itemcode = @ItemCode and BranchID = @BranchID and CompID = @CompID

		Else
		if (IsNull(@Rate,0) <= 0)
			Select  @Rate=Rate
			From Inv_Material_Master 
			Where ItemCode = @ItemCode and  BranchID = @BranchID and CompID = @CompID

		Return (IsNull(@Rate, 0))			
	
		End
GO

/****** Object:  UserDefinedFunction [dbo].[fn_Qc_InstrumentUseQty]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_Qc_InstrumentUseQty] (@InstrumentCode varchar(20),@FromDate  datetime, @BranchID varchar(3), @CompID  varchar(3))  
 
RETURNS float 

AS  

BEGIN 

	Declare @Qty  float

	Select @Qty= Sum(CheckingQty) from Qc_InstrumentTran_Dtl a 
							Where a.ItemCode=@InstrumentCode and a.Branchid=@BranchID and a.compid=@CompID and a.RefDate>=@FromDate
							group by a.ItemCode
							
	Return @Qty 

end
GO

/****** Object:  UserDefinedFunction [dbo].[fn_RemoveSpaces]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[fn_RemoveSpaces] (@String varchar(300))  

RETURNS varchar(300) 

AS  

BEGIN 
	
	Declare @Index int
	Select @Index=isNull(CharIndex(' ', @String), 0)
	
	While @Index > 0
	Begin
		Select @String=Stuff(@String, @Index, 1, '')
	
		Select @Index=isNull(CharIndex(' ', @String), 0)
	End

	Return(@String)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_RoundOff]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION  [dbo].[fn_RoundOff] (@Amount float, @RoundingValue float, @NoOfDecimals int, @Direction varchar(1))  
RETURNS float 

 AS  

BEGIN 

	Declare @Result float 
	Declare @NoOfBlocks int
  	Declare @tmpAmt float
	Declare @tmpStr varchar(300)


 	Select @NoOfBlocks = Convert(int, (@Amount / @RoundingValue))
  
  	Select @tmpAmt = @NoOfBlocks * @RoundingValue
  
  	If @tmpAmt <> @Amount 
		Begin
			Select @tmpStr = @RoundingValue / 2.0
   			
			Select @Result = Case @Direction When 'B' then 
						Case When (Round((@Amount - @tmpAmt), @NoOfDecimals) >= (@RoundingValue / 2.0)) then (@NoOfBlocks + 1) * @RoundingValue 
							else @tmpAmt
							end
				When 'U' then (@NoOfBlocks + 1) * @RoundingValue
				When 'D' then @tmpAmt
				Else @Amount
				End
		End
	else
		Select @Result=@Amount

	return (@Result)

END


GO

/****** Object:  UserDefinedFunction [dbo].[fn_Sal_Fg_Item_Rate]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_Sal_Fg_Item_Rate] (@ItemCode varchar(20), @BranchID varchar(3), @CompID varchar(3), @FinYear varchar(15))  
  
RETURNS money 
AS 
BEGIN
  Declare @Rate as money  

	select  @Rate=a.Rate from [SAL_INVOICE_Dtl]a 
	Left Outer Join [SAL_INVOICE_HDR] b on b.invoiceDate=a.InvoiceDate and b.InvoiceType=a.InvoiceType and b.InvoiceNo=a.InvoiceNo and 
	b.Branchid=a.Branchid and b.Compid=a.Compid and b.Finyear=a.Finyear  Where b.Branchid=@BranchID and b.Compid=@CompID and b.Finyear=@FinYear and 
	b.InvoiceTime = (Select Max(c.InvoiceTime) from  [SAL_INVOICE_HDR]c               
	Where c.Branchid=a.Branchid and c.Compid=a.Compid and c.Finyear=a.Finyear and c.InvoiceDate=a.InvoiceDate )and a.ItemCode=@ItemCode 
	Group By a.ItemCode,a.Rate
	Order By a.Rate
	Return @Rate

End;
GO

/****** Object:  UserDefinedFunction [dbo].[fn_Sal_Item_Rate]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_Sal_Item_Rate] (@ItemCode varchar(20), @CustomerCode int, @OrderType varchar(3), @OrderNo int, @OrderDate datetime, @BranchID varchar(3), @CompID varchar(3), @FinYear varchar(15))    
  
RETURNS float  
  
AS  
  
 BEGIN   
    
  Declare @Rate as float  
    
  If @CustomerCode = 0 and @OrderDate is Null  
     
   Begin  
    Select top 1 @Rate=Rate  
    from  
     (Select  top 1 x.Rate, 1 as Priority  
     From Sal_Order_Dtl x Inner Join   
      (Select Max(a.OrderDate) OrderDate  
      From Sal_Order_Hdr b Right Outer Join Sal_Order_Dtl a  
      On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate  
      Where a.ItemCode = @ItemCode and a.BranchID = @BranchID  and a.CompID = @CompID  
      ) y  
     On x.OrderDate=y.OrderDate  
     Where x.ItemCode = @ItemCode and x.BranchID = @BranchID  and x.CompID = @CompID  
  
     Union All  
    
     Select  Max((Rate/RatePer)) as Rate, 2 as Priority  
     From Sal_JobWorkOrder_Dtl x Inner Join   
      (Select Max(a.OrderDate) OrderDate  
      From Sal_JobWorkOrder_Hdr b Right Outer Join Sal_JobworkOrder_Dtl a  
      On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate  
      Where a.SentItemCode = @ItemCode and a.BranchID = @BranchID and a.CompID = @CompID  
         ) y  
     On x.OrderDate = y.OrderDate  
     Where x.SentItemCode = @ItemCode and x.BranchID = @BranchID  and x.CompID = @CompID  
       
     Union All   
/*  
     Select  Rate, 3 as Priority  
     From Sal_Order_Dtl x Inner Join   
      (Select Max(a.OrderDate) OrderDate  
      From Sal_JobWorkOrder_Hdr b Right Outer Join Sal_JobworkOrder_Dtl a  
      On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate  
      Where a.SentItemCode = @ItemCode and a.BranchID = @BranchID and a.CompID = @CompID   
         ) y  
     On x.OrderDate = y.OrderDate  
     Where x.SentItemCode = @ItemCode and x.BranchID = @BranchID  and x.CompID = @CompID   
  
     Union All  
*/       
     Select  Rate, 4 as Priority  
     From Sal_Customer_ItemMaster  
     Where Itemcode = @ItemCode and BranchID = @BranchID and CompID = @CompID  
       
     Union All   
       
     Select  Max((Rate/Rateper)) Rate,5 as Priority  
     From Sal_Jobwork_ItemMaster  
     Where SentItemCode = @ItemCode and BranchID = @BranchID and CompID = @CompID  
     Group By SentItemCode  
       
     Union All  
       
     Select  Rate, 6 as Priority  
     From Inv_Material_Master   
     Where ItemCode = @ItemCode and  BranchID = @BranchID and CompID = @CompID  
     )x  
    Where isNull(Rate, 0)<>0  
    Order by Priority  
   End  
  Else  
   if @OrderDate is Null --No Order Specified  
   
    Begin  
     Select top 1 @Rate=Rate  
     from  
      (
				/* Clarification required. if OrderDate is null then how we can take Rate from Order related objects
					Commented on 16/07/2015 by sujatha*/
				/*
				Select  top 1 x.Rate, 1 as Priority  
				From Sal_Order_Dtl x Inner Join   
				 (Select Max(a.OrderDate) OrderDate  
				 From Sal_Order_Hdr b Right Outer Join Sal_Order_Dtl a  
				 On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate  
				 Where b.CustomerCode=@CustomerCode and a.ItemCode = @ItemCode 
				 and a.BranchID = @BranchID  and a.CompID = @CompID  
       ) y  
      On x.OrderDate=y.OrderDate  
      Where x.ItemCode = @ItemCode and x.BranchID = @BranchID  and x.CompID = @CompID  
   
      Union All  
     
      Select  Max((Rate/RatePer)) as Rate, 2 as Priority  
      From Sal_JobWorkOrder_Dtl x Inner Join   
       (Select Max(a.OrderDate) OrderDate  
       From Sal_JobWorkOrder_Hdr b Right Outer Join Sal_JobworkOrder_Dtl a  
       On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate  
       Where b.CustomerCode = @CustomerCode and a.SentItemCode = @ItemCode and a.BranchID = @BranchID and a.CompID = @CompID  
          ) y  
      On x.OrderDate = y.OrderDate  
      Where x.SentItemCode = @ItemCode and x.BranchID = @BranchID  and x.CompID = @CompID  
        
      Union All   
				*/
 /*  
      Select  Rate, 3 as Priority  
      From Sal_Order_Dtl x Inner Join   
       (Select Max(a.OrderDate) OrderDate   
  
  
       From Sal_JobWorkOrder_Hdr b Right Outer Join Sal_JobworkOrder_Dtl a  
       On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate  
       Where b.CustomerCode = @CustomerCode and a.SentItemCode = @ItemCode and a.BranchID = @BranchID and a.CompID = @CompID   
          ) y  
      On x.OrderDate = y.OrderDate  
      Where x.SentItemCode = @ItemCode and x.BranchID = @BranchID  and x.CompID = @CompID   
   
      Union All  
 */       
      Select  Rate, 4 as Priority  
      From Sal_Customer_ItemMaster  
      Where Itemcode = @ItemCode and CustomerCode=@CustomerCode and BranchID = @BranchID and CompID = @CompID  
        
      Union All   
        
      Select  Max((Rate/Rateper)) Rate,5 as Priority  
      From Sal_Jobwork_ItemMaster  
      Where SentItemCode = @ItemCode and CustomerCode=@CustomerCode and BranchID = @BranchID and CompID = @CompID  
      Group By SentItemCode  
        
      Union All  
        
      Select  Rate, 6 as Priority  
      From Inv_Material_Master   
      Where ItemCode = @ItemCode and  BranchID = @BranchID and CompID = @CompID  
      )x  
     Where isNull(Rate, 0)<>0  
     Order by Priority  
    End  
     
   else  --Order specified  
   
    Begin  
     Select top 1 @Rate=Rate  
     from  
      (Select Rate, 1 as Priority  
      From Sal_Order_Dtl a  
      Where a.ItemCode = @ItemCode and OrderType =  @OrderType and OrderNo = @OrderNo and OrderDate = @OrderDate  
        and OrderDate = (Select Max(OrderDate) From Sal_Order_Dtl Where ItemCode = a.ItemCode        
       and BranchID = a.BranchID and CompID = a.CompID)  
       and a.BranchID = @BranchID  and a.CompID = @CompID  
     
      Union All  
     
      Select  Max((Rate/RatePer)) as Rate, 2 as Priority  
      From Sal_JobWorkOrder_Dtl  
      Where  SentItemCode = @ItemCode and OrderType =  @OrderType and OrderNo = @OrderNo and OrderDate = @OrderDate  
       and BranchID = @BranchID and CompID = @CompID  
      Group By SentItemCode  
        
      Union All   
 /*  
      Select Rate,3 as Priority  
      From Sal_Order_Dtl a  
      Where  ItemCode = @ItemCode ---and OrderType =  @OrderType and OrderNo = @OrderNo and  OrderDate =  @OrderDate  
        and OrderDate = (Select Max(OrderDate) From Sal_Order_Dtl Where ItemCode = a.ItemCode  
       and OrderType = a.OrderType and OrderNo = a.OrderNo and OrderDate = a.OrderDate  
       and BranchID = a.BranchID and CompID = a.CompID)  
        and BranchID = @BranchID  and CompID = @CompID  
   
      Union All  
 */       
      Select  Rate,4 as Priority  
      From Sal_Customer_ItemMaster  
      Where Itemcode = @ItemCode and CustomerCode=@CustomerCode and BranchID = @BranchID and CompID = @CompID  
        
      Union All   
        
      Select  Max((Rate/Rateper)) Rate, 5 as Priority  
      From Sal_Jobwork_ItemMaster  
      Where SentItemCode = @ItemCode and CustomerCode=@CustomerCode and BranchID = @BranchID and CompID = @CompID  
      Group By SentItemCode  
        
      Union All  
        
      Select  Rate, 6 as Priority  
      From Inv_Material_Master   
      Where ItemCode = @ItemCode and  BranchID = @BranchID and CompID = @CompID  
      )x  
     Where isNull(Rate, 0)<>0  
     Order by Priority  
    End  
  
  Return (Isnull(@Rate, 0))  
    
 End  
  
GO

/****** Object:  UserDefinedFunction [dbo].[fn_Sal_ItemRate_Latest]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[fn_Sal_ItemRate_Latest] (@ItemCode varchar(20), @BranchID varchar(3), @CompID varchar(3), @FinYear varchar(15))  
  
RETURNS money 
AS 
BEGIN

Declare @Rate as money  

Select Top 1 @Rate= x.Rate     
      From (Select a.InvoiceDate, a.Invoicetime, b.Itemcode, b.Rate, a.BranchId, a.CompID, a.Finyear
                From SAL_INVOICE_HDR a   Left Outer Join SAL_INVOICE_Dtl b
                On  a.InvoiceNo=b.InvoiceNo and a.InvoiceDate=b.InvoiceDate and a.InvoiceType=b.InvoiceType      
                                Where b.ItemCode=@ItemCode and a.BranchID=@BranchID  and a.CompID=@CompID --and a.Finyear=@FinYear 
             )x
      Inner Join ( Select b.InvoiceDate, Max(b.InvoiceTime) InvoiceTime      
                    From SAL_INVOICE_HDR b Right Outer Join SAL_INVOICE_Dtl a      
                    On  a.InvoiceNo=b.InvoiceNo and a.InvoiceDate=b.InvoiceDate and a.InvoiceType=b.InvoiceType      
                    Where a.ItemCode=@ItemCode and a.BranchID=@BranchID  and a.CompID=@CompID --and b.Finyear=@FinYear --and b.InvoiceDate='2019-05-04'   
                    and exists (Select Max(a.InvoiceDate) InvoiceDate      
                                From SAL_INVOICE_HDR b Right Outer Join SAL_INVOICE_Dtl a      
                                On  a.InvoiceNo=b.InvoiceNo and a.InvoiceDate=b.InvoiceDate and a.InvoiceType=b.InvoiceType      
                                Where a.ItemCode=@ItemCode and a.BranchID=@BranchID  and a.CompID=@CompID --and b.Finyear=@FinYear --and b.InvoiceDate='2019-05-04'   
                               )   
                              Group by b.InvoiceDate
                ) m      
      On m.InvoiceDate=x.InvoiceDate and x.InvoiceTime=m.InvoiceTime 
      Where x.BranchID=@BranchID  and x.CompID=@CompID 

      
Return @Rate

End
GO

/****** Object:  UserDefinedFunction [dbo].[fn_ToolBalance_Qty]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION  [dbo].[fn_ToolBalance_Qty] (@Regrind int, @Reform int,@BalQty int)  

RETURNS int 

 AS  

BEGIN 

	Declare @Qty  int
	Declare @RefTbl table
		(Qty int )
	
	insert into @RefTbl Select @Regrind

	insert into @RefTbl Select @Reform

	insert into @RefTbl Select @BalQty

	
	Select @Qty=min(Qty) from @RefTbl where Qty<>-1
	 

	return (@Qty)

END
GO

/****** Object:  UserDefinedFunction [dbo].[fn_WtAvgRate]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[fn_WtAvgRate] (@ItemCode  Varchar(20), @FromDate  datetime, @ToDate  datetime, @BranchID varchar(3), @CompID  varchar(3), @Finyear varchar(15))  
 
RETURNS float

AS  

BEGIN 

	Declare @WtAvgRate  float

	Select    @WtAvgRate = ( Sum(x.TransactionQty*x.Rate)/Sum(x.TransactionQty))  
	from 
		(

		Select a.TransactionQty, (select top 1 d.Rate/d.RatePer from PUR_PurchaseOrder_Dtl d where d.OrderNo = c.OrderNo and d.ItemCode = c.ItemCode and d.OrderType = c.OrderType and d.OrderDate = c.OrderDate and  d.RateEffDate <= a.transactionDate
			and c.CompID = d.CompID 	
/*			and c.Finyear = d.Finyear		 */
			order by d.RateEffDate desc) Rate
		from Inv_MaterialTran_Dtl a 
		Left Outer Join ReceiptEntry b
			On a.DocNo=b.ReceiptEntryNo and a.TransactionItem = b.ItemCode and a.DocDate=b.ReceiptEntryDate and a.DocType=b.ReceiptEntryType and a.CompID = b.CompID 
/*			and a.Finyear = b.Finyear		*/
		Left Outer Join dbo.fn_PUR_SCHEDULE(@ToDate, @ToDate, @BranchID, @CompID, @Finyear) c
			On b.RefNo = c.ScheduleNo and a.TransactionItem = c.ItemCode and b.RefDate = c.ScheduleDate  and  b.RefType = c.ScheduleType and a.CompID = c.CompID 
/*			and a.Finyear = c.Finyear		*/
		Where a.Doctype in (Select DocumentType from COM_DocumentType_Master Where PertainingTo='R' and ModuleName='REC' and BranchID=@BranchID and CompID=@CompID)
			and a.DocDate Between @FromDate and @ToDate and a.TransactionItem =@ItemCode and b.InwardType = 'P' and a.ToWarehouseCode = 'UI'    
			and a.BranchID=@BranchID and a.Compid = @CompID 
		) x

	If isNull(@WtAvgRate, 0) = 0 

	Select @WtAvgRate = isNull(Rate, 0) from Inv_Material_Master Where ItemCode =  @ItemCode and BranchID=@BranchID and Compid = @CompID

--	Select @WtAvgRate=isNull(@WtAvgRate, 0)

	return (@WtAvgRate)


END
GO

/****** Object:  UserDefinedFunction [dbo].[GetDateFormatStringArray]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[GetDateFormatStringArray](@DocDate datetime, @FormatStr varchar(100))  
RETURNS VARCHAR(200)  
AS  
BEGIN  

 DECLARE @BuildStr varchar(200);    
 DECLARE @tmpStr varchar(10);  
 DECLARE @rPos int;  
 DECLARE @rPos2 int;  
 DECLARE @Cnt int;  
 DECLARE @isLoopContinue bit;   
 DECLARE @MATCHSTRING varchar(50);  

 SET @MATCHSTRING=N'%FSYY%FS%FEYY%FE%YYYY%YY%YY+%MMMM%MMM%MM%M%WK%DD%D%';          
 SET @isLoopContinue='1';  
 SET @BuildStr='';  
 SET @Cnt=1;  
  
 If (LTrim(@FormatStr)<>'')  
  
 WHILE (@isLoopContinue='1')  
 BEGIN  
  SET @rPos=CharIndex(N'%', @FormatStr, @Cnt);  
  SET @rPos2=CharIndex(N'%', @FormatStr, @rPos+1);  
  
  If (@rPos>0 AND @rPos2>0)  
   BEGIN  
    SET @tmpStr=SubString(@FormatStr, @rPos, (@rPos2-@rPos+1));  
  
    If (CharIndex(@tmpStr, @MATCHSTRING, 1)>0)  
     BEGIN  
      SET @BuildStr = @BuildStr + SubString(@FormatStr, @Cnt, (@rPos-@Cnt)) + dbo.GetDateFormattedString(@DocDate, @tmpStr);  
     END  
    Else  
     BEGIN  
      SET @BuildStr=@BuildStr + Substring(@FormatStr, @Cnt, (@rPos2-@Cnt+1));  
      If (@rPos2=LEN(@FormatStr))  
       SET @isLoopContinue='0';  
     END  
  
    SET @Cnt=@rPos2+1;  
   END  
  Else  
   BEGIN  
    SET @BuildStr=@BuildStr + Substring(@FormatStr, @Cnt, (Len(@FormatStr)-@Cnt+1));  
    SET @isLoopContinue='0';  
   END    
 END  
  
 RETURN @BuildStr;  

/*
 DECLARE @BuildStr varchar(200);  
  
 DECLARE @tmpStr varchar(10);  
 DECLARE @rPos int;  
 DECLARE @rPos2 int;  
 DECLARE @Cnt int;  
  
 DECLARE @MATCHSTRING varchar(50);  
 SET @MATCHSTRING=N'%FSYY%FS%FEYY%FE%YYYY%YY%YY+%MMMM%MMM%MM%M%WK%DD%D%';  
  
 DECLARE @isLoopContinue bit;  
   
  
 SET @isLoopContinue='1';  
 SET @BuildStr='';  
 SET @Cnt=1;  
  
 If (LTrim(@FormatStr)<>'')  
  
 WHILE (@isLoopContinue='1')  
 BEGIN  
  SET @rPos=CharIndex(N'%', @FormatStr, @Cnt);  
  SET @rPos2=CharIndex(N'%', @FormatStr, @rPos+1);  
  
  If (@rPos>0 AND @rPos2>0)  
   BEGIN  
    SET @tmpStr=SubString(@FormatStr, @rPos, (@rPos2-@rPos+1));  
  
    If (CharIndex(@tmpStr, @MATCHSTRING, 1)>0)  
     BEGIN  
      SET @BuildStr = @BuildStr + SubString(@FormatStr, @Cnt, (@rPos-@Cnt)) + dbo.GetDateFormattedString(@DocDate, @tmpStr);  
     END  
    Else  
     BEGIN  
      SET @BuildStr=@BuildStr + Substring(@FormatStr, @Cnt, (@rPos2-@Cnt+1));  
      If (@rPos2=LEN(@FormatStr))  
       SET @isLoopContinue='0';  
     END  
  
    SET @Cnt=@rPos2+1;  
   END  
  Else  
   BEGIN  
    SET @BuildStr=@BuildStr + Substring(@FormatStr, @Cnt, (Len(@FormatStr)-@Cnt+1));  
    SET @isLoopContinue='0';  
   END  
  
 END  
  
 RETURN @BuildStr;  
*/
END
GO

/****** Object:  UserDefinedFunction [dbo].[GetDateFormattedString]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE Function [dbo].[GetDateFormattedString](@DocDate datetime, @FormatStr varchar(10)) 
RETURNS varchar(20)
As
BEGIN
 /* It should call from GetDateFormatPrefixSuffix */

  DECLARE @tmp varchar(20)
  
  SET @tmp='';
  
  Select @tmp=Case LTrim(@FormatStr)
  				When N'%FSYY%'
				  Then Case When Month(@DocDate)<4 Then Convert(varchar, Year(@DocDate)-1) Else Convert(varchar, Year(@DocDate)) End
				When N'%FS%'
				  Then Case When Month(@DocDate)<4 Then Right(Convert(varchar, Year(@DocDate)-1), 2) Else Right(Convert(varchar, Year(@DocDate)), 2) End
  				When N'%FEYY%'
				  Then Case When Month(@DocDate)<4 Then Convert(varchar, Year(@DocDate)) Else Convert(varchar, Year(@DocDate)+1) End
				When N'%FE%'
				  Then Case When Month(@DocDate)<4 Then Right(Convert(varchar, Year(@DocDate)), 2) Else Right(Convert(varchar, Year(@DocDate)+1), 2) End
				When N'%YYYY%'
				  Then Convert(varchar, Year(@DocDate))
				When N'%YY%'
				  Then Right(Convert(varchar, Year(@DocDate)), 2)
				When N'%YY+%'
				  Then Right(Convert(varchar, Year(@DocDate)+1), 2)
				When N'%MMMM%'
				  Then Convert(varchar, DATENAME(mm, @DocDate))
				When N'%MMM%'
				  Then Left(DATENAME(mm, @DocDate), 3)
				When N'%MM%'
				  Then Right('0'+ Convert(varchar, MONTH(@DocDate)), 2)
				When N'%M%'
				  Then Convert(varchar, MONTH(@DocDate))
				When N'%WK%'
				  Then Convert(varchar, DatePart(wk, @DocDate))
				When N'%DD%'
				  Then Right('0'+ Convert(varchar, Day(@DocDate)), 2)
				When N'%D%'
				  Then Convert(varchar, Day(@DocDate))
			  End 

  RETURN UPPER(@tmp);
END
GO

/****** Object:  UserDefinedFunction [dbo].[GetFormatedDocNoWithPrefixSuffix]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[GetFormatedDocNoWithPrefixSuffix](@DocDate datetime, @DocNo int, @DocNoWidth tinyint=0, 
@Prefix varchar(50)='', @Suffix varchar(50)='',@IsWidthAvoided bit='0')
 RETURNS VARCHAR(200)
 AS
 BEGIN
	DECLARE @tmpFormatStr varchar(200);

	DECLARE @PrefixStr varchar(200), @SuffixStr varchar(200);
	SELECT @PrefixStr='', @SuffixStr='';

	If (@Prefix<>'')
		SET @PrefixStr=dbo.GetDateFormatStringArray(@DocDate, @Prefix);

	If (@Suffix<>'')
		SET @SuffixStr=dbo.GetDateFormatStringArray(@DocDate, @Suffix);

	If @IsWidthAvoided='1'
	 Begin
		SET @tmpFormatStr=@PrefixStr + @SuffixStr;
	 End
	Else
	 Begin
		If (@DocNoWidth>0)
			SET @tmpFormatStr=@PrefixStr + RIGHT(REPLICATE('0', @DocNoWidth) + Convert(varchar, @DocNo), @DocNoWidth) + @SuffixStr;
		Else
			SET @tmpFormatStr=@PrefixStr + Convert(varchar, @DocNo) + @SuffixStr;
	 End

	RETURN @tmpFormatStr;
END
GO

/****** Object:  UserDefinedFunction [dbo].[RoundingOff]    Script Date: 17/06/2026 11:43:46 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO





CREATE FUNCTION  [dbo].[RoundingOff] (@Amt float, @RoundingValue float, @Direction int)  
RETURNS float 

 AS  

BEGIN 

	 declare @NoOfBlocks int
	 declare @tmpAmt float
	--  @Direction = 0 then None 
	--  @Direction = 1 then Nearest 
	--  @Direction = 2 then Upward 
	--  @Direction=3 then Downward
  	
  	If @Direction <> 0
		Begin
			select @NoOfBlocks = convert(int, (@Amt / @RoundingValue) )
			select @tmpAmt = @NoOfBlocks * @RoundingValue

			If (@tmpAmt <> @Amt)
			    	Begin     
					if @Direction = 1
					         	if ((@Amt - @tmpAmt) >= (@RoundingValue / 2) ) 
							select @tmpAmt =  (@NoOfBlocks + 1) * @RoundingValue 
					else
						if @Direction=2 
						          select  @tmpAmt =  (@NoOfBlocks + 1) * @RoundingValue
						else
							if @Direction=0 
						          		select @tmpAmt = @Amt
 
		 	    	End

		End


	return (@tmpAmt)

END



GO


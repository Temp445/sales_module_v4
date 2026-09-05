/* All Views List */
GO

/****** Object:  View [dbo].[Inspection_Parameter]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE view [dbo].[Inspection_Parameter]
as
select distinct x.InspectionTypeCode,x.SamplingPlanTypeCode,x.QualityPlanTypeCode,ShortName,ParameterCode,Parameter,ItemCode,InstrumentType,DimenSion,case when isnumeric(FromValue)=0 then case when dimension='Angle' then dbo.cDegree(FromValue) else FromValue end else case when dimension='Angle' then dbo.cDegree(FromValue) else FromValue end +' / '+case when dimension='Angle' then dbo.cDegree(ToValue) else ToValue end  end Specification,
case when isnumeric(FromValue)=0 then FromValue else Mean end Mean,Defect,ValiDate,Sequence,x.InspectionTypeLongName,x.InspectionTypeDescription,x.Noreading,Userid,Branchid,Compid from (
select ShortName,ParameterCode,Parameter,ItemCode,(select InstrumentType from BOM_InstrumentType_Master where InstrumentTypeCode=a.InstruMentTypeCode  and Branchid=a.Branchid and Compid=a.Compid)InstrumentType,DimenSion,FromValue,ToValue,
Mean,Defect,case when ValiDate=1 then 'YES' else 'NO'end ValiDate,Sequence,a.InspectionTypeCode,b.InspectionTypeDescription,b.InspectionTypeLongName,b.Noreading,b.SamplingPlanTypeCode,b.QualityPlanTypeCode,a.Userid,a.Branchid,a.Compid
from Qc_Inspection_Master a 
left outer join Qc_InspectionType_Master b on  a.InspectionTypeCode=b.InspectionTypeCode  and a.Branchid=b.Branchid and a.Compid=b.Compid
where a.Inspect=1)x
GO

/****** Object:  View [dbo].[MPS_Process]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE view [dbo].[MPS_Process] as
select *,isnull(dbo.fn_ItemNorms (x.StageCode,(select distinct top 1 StageCode from MPSMaster where ItemCode=x.ItemCode and OperationSequence=1),x.Branchid,x.CompID),1)Norms   from (
select ItemCode,StageCode,OperationSequence,OperationCode,Branchid,Compid from MPsMaster where Status=1
union
select * from (select Distinct (select top 1 ItemCode from MPsMaster where StageCode=ParentCode)ItemCode ,childCode,
convert(int,0) as operationSequence,'0' as OperationCode,Branchid,Compid from BOM_Structure_Dtl where Priority=1) y where y.ItemCode is not null)x
GO

/****** Object:  View [dbo].[PPAP]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[PPAP] as
select Distinct z.CustomerCode, z.ItemCode,z.DocType, z.StandardLName,
Case when Z.DocType in ('MTS','OPS','PAS','PRS','STS') then 
(select top 1 Userid from BOM_Inspection_Parameter where ItemCode=z.ItemCode) when Z.DocType='PFD' then 
(select top 1 Approvedby from ProcessFlow where ItemCode=z.ItemCode) when Z.DocType='PRD' then 
(select top 1 Userid from Control_Plan where ItemCode=z.ItemCode) when Z.DocType='FMA' then 
(select top 1 Userid from PFMEA where ItemCode=z.ItemCode) else ''
end PreparedBy, z.Branchid, z.Compid

 from (
Select Distinct 
(select top 1 s.CustomerCode from 
(select CustomerCode, ItemCode, Branchid, Compid from SAL_customer_ItemMaster
union 
select distinct CustomerCode, ReceivedItemCode, Branchid, Compid from SAL_JobWork_ItemMaster) s 
where isnull(r.ItemCode,y.ItemCode) = s.ItemCode and isnull(r.Branchid,y.Branchid) = s.Branchid and isnull(r.Compid,y.Compid) = s.Compid) CustomerCode,
r.ItemCode, y.ItemCode as StageCode, y.ItemDescription,y.DocType,

r.OperationSequence, y.StandardLname,y.Branchid,y.Compid from (
Select * from (
select a.ITemCode,c.ItemDescription,b.StandardsName as DocType,a.DrawingNo as Docno,DrawingDate as DocDate,
Case when ItemNature in ('CR','R') then Case when StandardCode in (4,2) then b.StandardLname else null end  
		 when ItemNature in ('T') then  Case when StandardCode in (5) then b.StandardLname else null end
		 when ItemNature in ('C','CC') then  Case when StandardCode not in (4,5) then b.StandardLname else null end
		 when ItemNature in ('F') then  Case when StandardCode not in (4,5) then b.StandardLname else null end
	else b.StandardLname end as StandardLname,a.Branchid,a.Compid
 from BOM_Drawing_Master a 

left outer join PAP_Standard b
	on b.Standardtype=a.Type

left outer join Inv_Material_Master c
	on c.ItemCode=a.ItemCode and c.Branchid=a.Branchid and c.Compid=a.Compid
)x where StandardLname is not null

and exists (
	Select * from BOM_Inspection_Parameter where ItemCode=x.ItemCode and branchid=x.Branchid and Compid=x.Compid and 
Type=(case x.StandardLname when 'MATERIAL STANDARD' then 'PT' when 'OPERATION STANDARD' then 'PT' when 'PROCESS STANDARD' then 'PS'
			when 'PACKING STANDARD' then 'PG' when 'SETUP STANDARD' then 'SP' end)
)


Union

select Distinct  a.ItemCode,c.ItemDescription, PFMEAType,PFMEANo,PFMEADate,'PFMEA',a.Branchid,a.Compid  from PFMEA a

left outer join Inv_Material_Master c
	on c.ItemCode=a.ItemCode and c.Branchid=a.Branchid and c.Compid=a.Compid

Union

Select Distinct a.ItemCode,c.ItemDescription, ProcessType,ProcessNo,ProcessDate,'PROCESS FLOW DIAGRAM',a.Branchid,a.Compid  from ProcessFlow a

left outer join Inv_Material_Master c
	on c.ItemCode=a.ItemCode and c.Branchid=a.Branchid and c.Compid=a.Compid

Union

Select Distinct a.ItemCode,c.ItemDescription, ControlType,ControlNo,ControlDate,'CONTROL PLAN',a.Branchid,a.Compid from Control_plan a 

left outer join Inv_Material_Master c
	on c.ItemCode=a.ItemCode and c.Branchid=a.Branchid and c.Compid=a.Compid


)y 

left outer join MPS_Process r on y.ItemCode = r.StageCode and y.Branchid = r.Branchid and y.Compid = r.Compid) z
where CustomerCode <> ''

union


select Distinct d.CustomerCode,a.ItemCode,InspectionTypeShortName,b.InspectionTypeLongName,a.UserId,a.Branchid,a.Compid from Inspection_Parameter a
left outer join Qc_InspectionType_Master b on b.InspectionTypeCode=a.InspectionTypeCode And b.Branchid=a.Branchid and b.Compid=a.Compid
left outer join Material_MAster c on c.ItemCode=a.ItemCode and c.Branchid=a.Branchid and c.Compid=a.Compid
left outer join Sal_Customer_Itemmaster d on d.ItemCode=a.ItemCode and d.Branchid=a.Branchid and d.Compid=a.Compid
where c.Itemnature='F' and c.Status=1
GO

/****** Object:  View [dbo].[FAS_Budget_LedgerMaster]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[FAS_Budget_LedgerMaster]

AS

	Select LedgerCode, LedgerName, Alias, LrRoute, MaxReached, isBillwise, CreditDays, OpBalance, 
		isCostCentreApplicable, isActive, CompCode, UserID, TranStamp
	From FAS_LedgerMaster

	Union All
	
	Select LedgerCode, LedgerName, '', LrRoute, '0', '0', 0, 0, '0', isActive, CompCode, UserID, TranStamp
	From FAS_LedgerMaster_Assumed
	Where TransferedCode=0
GO

/****** Object:  View [dbo].[FAS_Budget_GroupLedgers]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE View [dbo].[FAS_Budget_GroupLedgers]

as 

	Select top 100 percent Path, Type, Code, Particulars, CompCode  From 
	
		(
		Select GroupCode Code, GroupName Particulars, (GrRoute+GrWidth) as Path, 'Group' as Type, CompCode
		From FAS_GroupMaster 	
	
		Union All 
	
		Select LedgerCode, LedgerName Particulars, LrRoute Path, 'Ledger' as Type, CompCode
		from FAS_Budget_LedgerMaster 	
		) m
	
	Order by Path, Type
GO

/****** Object:  View [dbo].[Tmp_Stock_Tran]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

Create View [dbo].[Tmp_Stock_Tran] as 
	Select * From Inv_MaterialTran_Dtl a
	Where convert(varchar(10), TransactionDate,120) + ' '+ convert(varchar(8),TransactionTime,108) >'2015-04-01 17:30:00' 
		and Finyear='04-2015 03-2016' 
		and Not Exists
			(
				Select * From StockList_20152016 s
				Where s.ItemCode=a.TransactionItem and s.WHSCode=a.FromWarehouseCode and s.RackBin=a.FromRackBinCode
			)
			and a.FromWarehouseCode<>'WIP'
			and (IsNumeric(a.FromWarehouseCode)=0 and IsNumeric(a.ToWarehouseCode)=0)
	 
GO

/****** Object:  View [dbo].[Tmp_Stock_StockList]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

Create View [dbo].[Tmp_Stock_StockList] as 
Select TransactionItem, FromWarehouseCode, FromRackBinCode, Sum(TransactionQty) as TransactionQty
From Tmp_Stock_Tran 
Group by TransactionItem, FromWarehouseCode, FromRackBinCode
GO

/****** Object:  View [dbo].[AssemblyWorkOrder]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO





CREATE view [dbo].[AssemblyWorkOrder] as 
select a.AssemblyWorkOrderType,a.AssemblyWorkOrderNo,a.AssemblyWorkOrderDate,
a.ParentCode,a.WorkOrderQty,a.CompletedQty,a.CompletedStatus, a.WOShortClose,
a.StoreRequestStatus,a.RejectionQty, a.ScrapQty, b.ItemCode,b.RequiredQty,b.Norms,
b.IssuedQty,b.MRNStatus,a.Branchid,a.Compid,b.Finyear
from PPC_ASSEMBLYWORKORDER_HDR a,PPC_ASSEMBLYWORKORDER_DTL b where 
a.AssemblyWorkOrderType=b.AssemblyWorkOrderType 
and a.AssemblyWorkOrderNo=b.AssemblyWorkOrderNo 
and a.Branchid=b.Branchid and a.Compid=b.Compid 
and a.Finyear=b.Finyear






GO

/****** Object:  View [dbo].[itemnorms]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[itemnorms] as
select  processWorkOrderType,processWorkOrderNo,processWorkOrderDate,
ParentCode,WorkOrderQty,CompletedQty,CompletedStatus, WOShortClose,
StoreRequestStatus,RejectionQty, ScrapQty, fromstagecode,RequiredQty,Norms,
IssuedQty,MRNStatus,Branchid,Compid,Finyear from PPC_ProcessWorkOrder
union 
select  AssemblyWorkOrderType,AssemblyWorkOrderNo,AssemblyWorkOrderDate,
ParentCode,WorkOrderQty,CompletedQty,CompletedStatus, WOShortClose,
StoreRequestStatus,RejectionQty, ScrapQty, ItemCode,RequiredQty,Norms,
IssuedQty,MRNStatus,Branchid,Compid,Finyear from AssemblyWorkOrder
---where ParentCode like '6u0%'





GO

/****** Object:  View [dbo].[BOM_Item_Norms]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO





CREATE View [dbo].[BOM_Item_Norms] as

Select a.ParentCode, a.ChildCode, Qty, BranchID, CompID 
from Bom_structure_dtl a

Union All

Select b.ItemCode, b.StageCode, 1, BranchID, CompID 
from Bom_MPS_Master b 








GO

/****** Object:  View [dbo].[BOM_Structure]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create view [dbo].[BOM_Structure] as
select a.ParentCode, a.ChildCode, a.Qty, a.RevisionNo, a.Explodable, a.Percentage, 
a.Priority, isNull(d.BaseQty,1) BaseQty, d.EffectiveDate, d.Status, a.Branchid, a.Compid from 
BOM_Structure_Dtl a left outer join BOM_Structure_Hdr d on a.ParentCode = d.ParentCode 
and a.RevisionNo = d.RevisionNo and a.Branchid = d.Branchid and a.Compid = d.Compid 
where d.Status = 1 and a.RevisionNo = (select Max(b.RevisionNo) from BOM_Structure_Dtl b where 
b.ParentCode = a.ParentCode and b.Branchid = a.Branchid and b.Compid = a.Compid)
and a.Priority = (select Min(c.Priority) from BOM_Structure_Dtl c where 
c.ParentCode = a.ParentCode and c.ItemGroupKey = a.ItemGroupKey and c.Branchid = a.Branchid 
and c.Compid = a.Compid) 
GO

/****** Object:  View [dbo].[Calibration_Parameter]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE View [dbo].[Calibration_Parameter] as
Select Distinct a.ItemCode,a.Sno,a.Parameter,convert(varchar,a.FromValue) +' / '+ convert(varchar,a.ToValue) as Specification,
 a.Branchid,a.Compid from Qc_Calibration_Parameter a 
union
Select Distinct ItemCode,'O' as Sno,(Select Description From TermsMaster Where TypeFlg = 'Other Parameters' and Code = b.OtherParameter and BranchID = b.BranchID and CompID = b.CompID)OtherParamName,'OK / NOT OK' as Specification
,Branchid,Compid from Qc_CalibrationStatus_OtherParameter b
GO

/****** Object:  View [dbo].[CalibrationIssueParameter]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE View [dbo].[CalibrationIssueParameter] as 
Select Distinct IssueType,IssueNo,IssueDate,a.ItemCode,a.Sno,a.Parameter,convert(varchar,b.FromValue) +' / '+ convert(varchar,b.ToValue) as Specification, ObservedValue,a.Branchid,a.Compid from Qc_CalibrationStatus_Parameter a 
left outer join Qc_Calibration_Parameter b on a.ItemCode=b.ItemCode and a.Sno=b.sno and a.Parameter=b.Parameter
and a.Branchid=b.Branchid and a.Compid=b.Compid
Union
Select Distinct IssueType,IssueNo,IssueDate,ItemCode,'O' as Sno,(Select Description From TermsMaster Where TypeFlg = 'Other Parameters' and Code = b.OtherParameter and BranchID = b.BranchID and CompID = b.CompID)OtherParamName,'OK / NOT OK' as Specification,Status
,Branchid,Compid from Qc_CalibrationStatus_OtherParameter b
GO

/****** Object:  View [dbo].[ComplaintRegister]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO





CREATE view [dbo].[ComplaintRegister] as 
select x.ComplaintRegisterType,x.ComplaintRegisterNo,x.ComplaintRegisterDate,
x.RefNo,x.RefDate,x.ComplaintReceivedDate,x.TargetDate,x.ComplaintNature,x.ItemCode,
(select d.ItemDescription from  INV_Material_Master d where d.ItemCode = x.ItemCode
and d.Branchid = x.Branchid and d.Compid = x.Compid) ItemDescription,x.Remarks,
x.ComplaintDetails,x.BaseRefNo,x.BaseRefDate,x.ActionTaken,x.Feedback,x.FeedbackDate,
case when  a.SourceTag = 'E' then (select b.PartyName from FAS_Party_Master b where 
b.PartyCode = a.Source and b.Branchid = a.Branchid and b.Compid = a.Compid) else (select 
c.Description from  TermsMaster c where c.TypeFlg = 'Department' and c.Code = a.Source  
and c.Compid = a.Compid) end SourceName, case when 
a.DestinationTag = 'E' then (select b1.PartyName from FAS_Party_Master b1 where 
b1.PartyCode = a.Destination and b1.Branchid = a.Branchid and b1.Compid = a.Compid) 
else (select  c1.Description from TermsMaster c1 where c1.TypeFlg = 'Department' and c1.Code = a.Destination 
and c1.Compid = a.Compid) end DestinationName,x.Branchid,
x.Compid,x.Finyear from QC_ComplaintRegister x left outer join QC_ComplaintType_Master a 
on a.ComplaintType = x.ComplaintType and a.Branchid = x.Branchid and a.Compid = x.Compid 











GO

/****** Object:  View [dbo].[Control_Plan]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE VIEW [dbo].[Control_Plan] As    
Select Distinct a.ControlType, a.ControlNo, a.ControlDate, a.ItemCode, a.KeyContact, a.KeyContactPh,   
  a.SupplierApprovalDate, a.CustomerEngApprovalDate, a.CustomerQtlyApprovalDate, a.OtherApprovalDate,  
  b.ProcessNo, b.ProcessName, b.StageCode, b.Revisionno, b.Machine, b.Chrno, b.ChrProduct, b.ChrProcess,   
  b.Splchrclass, b.Specification, b.Evaluation, b.SampleSize ,b.SampleFreq, b.ControlMethod, b.ReactionPlan,   
  a.UserID, a.BranchID, a.CompID,b.TranStamp   
From PAP_ControlPlan_Hdr a     
Left Outer Join PAP_ControlPlan_Dtl b     
  On a.ControlNo=b.ControlNo and a.ControlDate=b.ControlDate and a.ControlType=b.Controltype    
  and a.RevisionNo=b.RevisionNo and a.BranchID=b.BranchID and a.CompID=b.CompID      
 Where a.Revisionno=(Select Max(h1.Revisionno) From PAP_ControlPlan_Hdr h1    
         Where h1.ControlNo=a.ControlNo and h1.ControlDate=a.ControlDate and h1.ControlType=a.ControlType            
          and h1.BranchID=a.BranchID and h1.CompID=a.CompID    
        ) 
GO

/****** Object:  View [dbo].[ControlPlan_Dtl]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create View [dbo].[ControlPlan_Dtl] as
select * from PAP_ControlPlan_Dtl a where Revisionno=(
select Max(Revisionno) from PAP_ControlPlan_Dtl where ControlType=a.ControlType
and Controlno=a.Controlno and ControlDate=a.ControlDate
and Branchid=a.Branchid and Compid=a.Compid)
GO

/****** Object:  View [dbo].[Controlplan_Hdr]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[Controlplan_Hdr] as 
select * from PAP_Controlplan_Hdr a where Revisionno=
(Select Max(RevisionNo) from PAP_Controlplan_Hdr where ControlType=a.ControlType and ControlNo=a.ControlNo
and ControlDate=a.ControlDate and Branchid=a.Branchid and Compid=a.Compid)


GO

/****** Object:  View [dbo].[Criteria_Reason]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create view [dbo].[Criteria_Reason]
as
select 'R' Type,ReasonCode,ShortName,Reason from PPC_Reason_Master
union
select 'C' Type,CriteriaCode,left(CriteriaDescription,3)ShortName,CriteriaDescription 
from QC_Criteria_Master
GO

/****** Object:  View [dbo].[Dailyplan_Dtl]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


create view [dbo].[Dailyplan_Dtl] as 
select * from PPC_Quickproductionplan_dtl a where revisionno=(select max(Revisionno) from 
PPC_Quickproductionplan_dtl where DocType=a.Doctype and docno=a.Docno and Docdate=a.Docdate
and branchid=a.Branchid and Compid=a.Compid )
GO

/****** Object:  View [dbo].[DismantlingDc]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE VIEW [dbo].[DismantlingDc]
AS 
SELECT a.DocType, a.DocNo,a.DocDate, a.OrderType, a.OrderNo, a.OrderDate, a.VendorCode, a.UserID,
 a.BranchID, a.CompID, a.Finyear, b.ItemGroupKey, b.SentItemCode, b.RecdItemCode,
 b.ScheduleDate, b.ToBeRecdQty as Qty, b.Norms
FROM PUR_DismantlingDc_Hdr a,PUR_DismantlingDc_Dtl b
where a.DocType=b.DocType and a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.Branchid=b.Branchid and 
a.Compid=b.Compid and a.Finyear=b.Finyear
GO

/****** Object:  View [dbo].[DrawingMaster]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[DrawingMaster]
AS
SELECT DISTINCT ItemCode,DrawingNo,DrawingDate,DrawingCode,Description,Modification,RevisionNo,Revisiondate,
Type,Scale,Status,TotalCopies,DrawingPath,UserId,BranchId,CompId

FROM         dbo.BOM_Drawing_Master a
WHERE     Revisionno =
            (SELECT     MAX(Revisionno)
              FROM          BOM_Drawing_Master
              WHERE      a.ItemCode = ItemCode AND a.DrawingNo = DrawingNo AND a.DrawingDate = DrawingDate AND a.DrawingCode = DrawingCode 
                         AND a.Branchid = BranchID AND a.Compid = CompID )
GO

/****** Object:  View [dbo].[EmployeeDetails]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE View [dbo].[EmployeeDetails] as
Select distinct 
EmployeeCode,EmployeeName,DateofBirth,Sex,BloodGroupCode,DateofJoining,DesignationCode
,DepartmentCode,Address1,Address2,Address3,City,State,Pincode,PhoneNo,TeamCode,PhotoCode
,Userid,Branchid,Compid,TranStamp from PPC_TemporaryEmployee_Master Where Status = 1
union
Select distinct 
a.EmployeeCode,a.EmployeeName,a.DateofBirth,a.Sex,a.BloodGroupCode,a.DateofJoining,a.DesignationCode,
DepartmentCode,b.Address1,b.Address2,b.Address3,b.City,b.State,b.Pincode,b.PhoneNo,
a.TeamCode,a.PhotoCode,a.Userid,a.Branchid,a.Compid,a.TranStamp
from Pay_Employee_Master a Left Outer Join PAY_AddressDetails b On a.EmployeeCode = b.EmployeeCode
and a.Branchid=b.Branchid and a.Compid=b.Compid Where b.AddressType = 'P' and a.Status = 'W'
GO

/****** Object:  View [dbo].[EmployeeMachine]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE View [dbo].[EmployeeMachine] as

Select distinct a.EmployeeCode,a.MachineCode,a.Status,a.Userid,a.Branchid,a.Compid,a.TranStamp 
from PPC_TemporaryEmployeeMachine_Dtl a

Union

Select distinct b.EmployeeCode,b.MachineCode,b.Status,b.Userid,b.Branchid,b.Compid,b.TranStamp 
from PPC_EmployeeMachine_Dtl b
GO

/****** Object:  View [dbo].[eTest]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[eTest] as
SELECT Distinct a.ItemCode,b.ShortName,b.ParameterCode,b.Parameter,b.InstrumentTypeCode,b.Dimension,
b.FromValue,b.ToValue,b.Mean,b.Defect,b.Validate,b.Sequence,b.Status,b.Type,b.Userid,b.Branchid,b.Compid
 FROM INV_Material_Master a
left outer join 
(select * from Bom_Inspection_Parameter where ItemCode='2U070331Z0' and type='PS')b
on a.Branchid=b.Branchid
where a.ItemCode in('2U070331Z0','2U070701Z0','2U020391T0','2U020321Y0','2U070161Z0','2U070291Z0','2U020421T0',
'2U020401T0','2U020411T0','2U020311Y0','2U070241Z0','2U070261Z0','2U070321Z0','2U020011Y0',
'2U070341Z0','2U070351Z0','2U071511Z0','2U070231Z0','2U070251Z0','2U070271Z0','2U020151P0',
'2U020491S0','2U040011Y0','2U070011Z0','2U020121Y0','2U070191Z0','2U120011Y0','2U070121Z0',
'2U071420Z0')
GO

/****** Object:  View [dbo].[ExpressionGrouping]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE View [dbo].[ExpressionGrouping] AS

--Active Parent Childs
Select a.ExprGroupID, a.ExprShortName, a.ExprGroupName, a.Route, a.Width, a.isActive, Convert(bit, '1') as ByParentStatus, a.BranchID, a.CompID, a.ModuleCode
from COM_FormulaExpr_Group a 
where not exists 
	(Select * from COM_FormulaExpr_Group b 
	where (b.Route+b.Width) = Left(a.Route, Len(b.Route+b.Width)) and Len(b.Route) < Len(a.Route)  and b.isActive='0'
		and b.BranchID=a.BranchID and b.CompID=a.CompID and b.ModuleCode=a.ModuleCode)

union all

--Inactive Parent Childs
Select a.ExprGroupID, a.ExprShortName, a.ExprGroupName, a.Route, a.Width, a.isActive, Convert(bit, '0') as ByParentStatus, a.BranchID, a.CompID, a.ModuleCode
from COM_FormulaExpr_Group a 
where exists 
	(Select * from COM_FormulaExpr_Group b 
	where (b.Route+b.Width) = Left(a.Route, Len(b.Route+b.Width)) and Len(b.Route) < Len(a.Route) and b.isActive='0'
		and b.BranchID=a.BranchID and b.CompID=a.CompID and b.ModuleCode=a.ModuleCode)



GO

/****** Object:  View [dbo].[Fas_BankRecon_Vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[Fas_BankRecon_Vw]
AS
SELECT     a.DocNo, a.VrRoute, a.LrCode, a.DrAmt, a.CrAmt, c.ChqNo, b.DocDate, c.ChqDate, c.BankDate, c.ChqFavor, c.isLrCode, a.CompCode
FROM        FAS_VoucherTranSub a INNER JOIN
                    FAS_VoucherTranMast b ON a.DocNo = b.DocNo AND a.DocDate=b.DocDate AND a.VrRoute = b.VrRoute AND a.CompCode = b.CompCode INNER JOIN
                     FAS_VoucherTranPay c ON a.DocNo = c.DocNo AND a.DocDate=c.DocDate  AND a.VrRoute = c.VrRoute AND a.CompCode = c.CompCode
GO

/****** Object:  View [dbo].[FAS_BSPLGroupMaster_vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[FAS_BSPLGroupMaster_vw]   
AS   
SELECT a.SubGrpCode, a.SubGrpShortName, a.SubGrpLongName, a.SubGrpRoute, a.SubGrpWidth, a.isDebitBalance,    
    a.isExpand, a.Seq, a.EffStatus, CONVERT(bit, '1') AS ByParentStatus, a.CompCode  
FROM  FAS_BSPLGroupMaster a  
WHERE NOT EXISTS  
    (
        SELECT *  
        FROM FAS_BSPLGroupMaster b  
        WHERE a.SubGrpRoute Like (b.SubGrpRoute+ b.SubGrpWidth + '%') AND   
            Len(b.SubGrpRoute) < Len(a.SubGrpRoute) AND b.EffStatus = 'AH' AND b.CompCode = a.CompCode
    )
      
UNION ALL  

SELECT a.SubGrpCode, a.SubGrpShortName, a.SubGrpLongName, a.SubGrpRoute, a.SubGrpWidth, a.isDebitBalance, 
    a.isExpand, a.Seq, a.EffStatus, CONVERT(bit, '0') AS ByParentStatus, a.CompCode  
FROM FAS_BSPLGroupMaster a  
WHERE EXISTS  
    (
        SELECT *
        FROM FAS_BSPLGroupMaster b  
        WHERE a.SubGrpRoute Like (b.SubGrpRoute+ b.SubGrpWidth + '%') AND 
            Len(b.SubGrpRoute) < Len(a.SubGrpRoute) AND b.EffStatus = 'AH' AND b.CompCode = a.CompCode
    )  
GO

/****** Object:  View [dbo].[FAS_Budget_LrOnly_Vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[FAS_Budget_LrOnly_Vw]

AS

Select b.PlanFor, Substring(a.SLCode, 2, Len(a.SLCode)) as SLCode, b.PlanDate, 
	Sum(Case When b.DrCode=a.SlCode Then Amount Else 0 End) DrAmt, 
	Sum(Case When b.CrCode=a.SlCode Then Amount Else 0 End) CrAmt,
	Left(b.CompCode, 4) as CompCode
				From
					(Select Distinct DrCode as SLCode
					From FAS_Budget_Dtl
					Where DrCode Like 'L%' and CompCode=Left(CompCode, 4)
					Union 
					Select Distinct CrCode as SLCode
					From FAS_Budget_Dtl
					Where CrCode Like 'L%' and CompCode=Left(CompCode, 4) 
					) a 
Inner Join FAS_Budget_Dtl b
	On (a.SLCode=b.DrCode OR a.SLCode=b.CrCode)
Group By b.PlanFor, Substring(a.SLCode, 2, Len(a.SLCode)), b.PlanDate, Left(b.CompCode, 4)
GO

/****** Object:  View [dbo].[FAS_Budget_VoucherMaster]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[FAS_Budget_VoucherMaster]

AS

Select VoucherCode, VoucherName, Alias, VrWidth, VrRoute, MaxReached, VAbbr, Prefix, Suffix, isAutoNumber, StartNo, WidthNo, ReStartAt,
	UseEntryWiseNarr, PrintAtSave, isActive, CompCode
From FAS_VoucherMaster

Union All

Select Distinct -909000-n.VoucherCode, 'BUDGET - ' + n.VoucherName, '', Right(m.VrRoute, 2), 'VV', '0', '', '', '', '1', 1, 0, 'FinYearly', '0', '0', '1', m.CompCode
From FAS_Budget_Dtl	m
Left Outer Join FAS_VoucherMaster n 
On m.VrRoute='VV'+n.VrRoute+n.VrWidth and m.CompCode=n.CompCode
GO

/****** Object:  View [dbo].[FAS_Budget_Vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[FAS_Budget_Vw]

AS

Select b.PlanFor, a.SLCode, b.PlanDate, 
	Sum(Case When b.DrCode=a.SlCode Then Amount Else 0 End) DrAmt, 
	Sum(Case When b.CrCode=a.SlCode Then Amount Else 0 End) CrAmt,
	Left(b.CompCode, 4) as CompCode
				From
					(Select Distinct DrCode as SLCode
					From FAS_Budget_Dtl
					Where CompCode=Left(CompCode, 4)
					Union 
					Select Distinct CrCode as SLCode
					From FAS_Budget_Dtl
					Where CompCode=Left(CompCode, 4) 
					) a 
Inner Join FAS_Budget_Dtl b
	On (a.SLCode=b.DrCode OR a.SLCode=b.CrCode)
Group By b.PlanFor, a.SLCode, b.PlanDate, Left(b.CompCode, 4)
GO

/****** Object:  View [dbo].[FAS_CostCentre_Vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE View [dbo].[FAS_CostCentre_Vw] 

as

	Select a.CategoryCode, b.LongName as CategoryName, a.CostCentreCode, a.ShortName, a.LongName, 
		b.isMandatory, b.Priority, b.isActive as isCategoryActive, a.Description, a.isActive, a.CompCode
	from FAS_CostCentreMaster a Left Outer Join FAS_CostCategoryMaster b
	On a.CategoryCode=b.CategoryCode and a.CompCode=b.CompCode
GO

/****** Object:  View [dbo].[Fas_GroupMasterVW]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE View [dbo].[Fas_GroupMasterVW] 

AS
	
	Select GroupCode, GroupName, Alias, GrWidth, GrRoute, MaxReached, Master, isActive, Convert(bit, '1') as ByParentStatus, CompCode
	from Fas_GroupMaster a
	Where not exists 
		(Select * from Fas_GroupMaster b Where (b.GrRoute+b.GrWidth) = Left(a.GrRoute, Len(b.GrRoute+b.GrWidth))
		and Len(b.GrRoute) < Len(a.GrRoute) and b.CompCode = a.CompCode)

	Union All

	Select GroupCode, GroupName, Alias, GrWidth, GrRoute, MaxReached, Master, isActive, Convert(bit, '0') as ByParentStatus, CompCode
	from Fas_GroupMaster a
	Where exists 
		(Select * from Fas_GroupMaster b Where (b.GrRoute+b.GrWidth) = Left(a.GrRoute, Len(b.GrRoute+b.GrWidth))
		and Len(b.GrRoute) < Len(a.GrRoute) and b.CompCode = a.CompCode)
GO

/****** Object:  View [dbo].[FAS_LEDGERVIEW]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE VIEW [dbo].[FAS_LEDGERVIEW] 

AS 

SELECT * FROM FAS_LEDGERMASTER
GO

/****** Object:  View [dbo].[FAS_MEM_Voucher_Vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[FAS_MEM_Voucher_Vw]

AS

	SELECT a.DocType, a.DocNo, a.DocDate,  b.VTSNo, a.VrRoute, a.Amount as VrAmt, a.Ref,  a.FileRef, a.PayMode, a.Narration, b.LrCode, b.DrAmt, b.CrAmt, a.isCancelled, a.CompCode
	FROM FAS_MEM_VoucherTranMast a Left Outer JOIN FAS_MEM_VoucherTranSub b 
	On a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.VrRoute=b.VrRoute and a.CompCode=b.CompCode
GO

/****** Object:  View [dbo].[FAS_MEM_VrDetails_Vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[FAS_MEM_VrDetails_Vw]

AS

	SELECT a.DocDate, a.DocNo, b.VTSNo, isNull(b.LrCode, 0) as LrCode, isNull(d.LrRoute, '') AS Ledger_Route, 
		Case When a.isCancelled='1' Then 'CANCELLED' Else d.LedgerName End as LedgerName, isNull(d.OpBalance, 0) as OpBalance, c.VoucherName, 
			a.VrRoute AS Voucher_Route, a.Amount as VrAmt, isNull(b.DrAmt, 0) as DrAmt, isNull(b.CrAmt, 0) as CrAmt, a.Narration, a.FileRef, a.Paymode, a.CompCode

	FROM FAS_MEM_VoucherTranMast a  LEFT OUTER JOIN FAS_MEM_VoucherTranSub b 	

		ON a.DocNo = b.DocNo AND a.DocDate=b.DocDate AND a.VrRoute = b.VrRoute AND a.CompCode = b.CompCode 

	INNER JOIN FAS_VoucherMaster c  

		ON a.VrRoute = c.VrRoute + c.VrWidth AND  a.CompCode = c.CompCode 

	LEFT OUTER JOIN FAS_LedgerMaster d 

		ON b.LrCode = d.LedgerCode AND a.CompCode = d.CompCode

--	WHERE a.isCancelled = '0'
GO

/****** Object:  View [dbo].[FAS_ScheduleMaster_vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[FAS_ScheduleMaster_vw]
AS
/*Active Parent Childs*/ SELECT a.SchCode, a.SchShortName, a.SchLongName, a.SchRoute, a.SchWidth, a.isBalanceSheetEntry, a.isDebitBalance, isExpand, isShowInSchReport, a.Seq, 
                      a.EffStatus, CONVERT(bit, '1') AS ByParentStatus, a.CompCode
FROM         FAS_ScheduleMaster a
WHERE     NOT EXISTS
                          (SELECT     *
                            FROM          FAS_ScheduleMaster b
                            WHERE      (b.SchRoute + b.SchWidth) = LEFT(a.SchRoute, Len(b.SchRoute + b.SchWidth)) AND Len(b.SchRoute) < Len(a.SchRoute) AND 
                                                   b.EffStatus = 'AH' AND b.CompCode = a.CompCode)
UNION ALL
/*Inactive Parent Childs*/ SELECT a.SchCode, a.SchShortName, a.SchLongName, a.SchRoute, a.SchWidth, a.isBalanceSheetEntry, a.isDebitBalance, isExpand, isShowInSchReport, 
                      a.Seq, a.EffStatus, CONVERT(bit, '0') AS ByParentStatus, a.CompCode
FROM         FAS_ScheduleMaster a
WHERE     EXISTS
                          (SELECT     *
                            FROM          FAS_ScheduleMaster b
                            WHERE      (b.SchRoute + b.SchWidth) = LEFT(a.SchRoute, Len(b.SchRoute + b.SchWidth)) AND Len(b.SchRoute) < Len(a.SchRoute) AND 
                                                   b.EffStatus = 'AH' AND b.CompCode = a.CompCode)
GO

/****** Object:  View [dbo].[FAS_Voucher_Vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE VIEW [dbo].[FAS_Voucher_Vw]

AS

	SELECT a.DocType, a.DocNo, a.DocDate,  b.VTSNo, a.VrRoute, a.Amount as VrAmt, a.Ref,  a.FileRef, a.PayMode, a.Narration, b.LrCode, b.DrAmt, b.CrAmt, a.isCancelled, a.CompCode
	FROM FAS_VoucherTranMast a Left Outer JOIN FAS_VoucherTranSub b 
	On a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.VrRoute=b.VrRoute and a.CompCode=b.CompCode
GO

/****** Object:  View [dbo].[FAS_VoucherCurrency]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[FAS_VoucherCurrency]                    
As    
/*    
  DrAmt, CrAmt, VrAmt are TransactionAmt in Voucher Currency    
  DrAmtBase, CrAmtBase, VrAmtBase are TransactionAmt in Base Currency    
  DrAmtLrCur, CrAmtLrCur  are TransactionAmt in Ledger Currency    
*/    
Select a.DocType, a.DocNo, a.DocDate, a.VrRoute,     
    isNull(b.LrCode, 0) as LrCode, isNull(c.LedgerName, '') as LedgerName, isNull(c.LrRoute, '') as LrRoute,    
    isNull(b.VTSNo, 0) as VTSNo,     
    isNull(b.DrAmt, 0) as DrAmt, isNull(b.CrAmt, 0) as CrAmt, a.Amount as VrAmt,    
    isNull(b.DrAmt, 0) as DrAmtBase, isNull(b.CrAmt, 0) as CrAmtBase, a.Amount as VrAmtBase,
    isNull(b.DrAmt, 0) as DrAmtLrCur, isNull(b.CrAmt, 0) as CrAmtLrCur,
    /*
    isNull(b.DrAmt*a.ExchRate, 0) as DrAmtBase, isNull(b.CrAmt*a.ExchRate, 0) as CrAmtBase, a.VrAmt*a.ExchRate as VrAmtBase,    
    Case When a.CurrencyCode<>c.CurrencyCode Then isNull(b.DrAmt*a.ExchRate, 0) Else isNull(b.DrAmt, 0) End as DrAmtLrCur,    
    Case When a.CurrencyCode<>c.CurrencyCode Then isNull(b.CrAmt*a.ExchRate, 0) Else isNull(b.CrAmt, 0) End as CrAmtLrCur,    
    */
    0 as CurrencyCode, 1 as ExchRate, /* a.CurrencyCode, a.ExchRate, */
    a.isCancelled, a.CompCode, a.Narration, a.Ref, a.FileRef, a.Paymode                                 
From FAS_VoucherTranMast a                                        
Left Outer Join FAS_VoucherTranSub b                                        
    On a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.VrRoute=b.VrRoute and a.CompCode=b.CompCode  /* a.TVID=b.TVID and a.CompCode=b.CompCode */
Left Outer Join FAS_LedgerMaster_AuthTbl c    
    On b.LrCode=c.LedgerCode and b.CompCode=c.CompCode 
GO

/****** Object:  View [dbo].[FAS_VrDetails_Vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE VIEW [dbo].[FAS_VrDetails_Vw]

AS

	SELECT a.DocDate, a.DocNo, Case When a.isCancelled='1' Then 1 Else b.VTSNo End as VTSNo, isNull(b.LrCode, 0) LrCode, isNull(d.LrRoute, '') AS Ledger_Route, 
			Case When a.isCancelled='1' Then 'CANCELLED' Else d.LedgerName End as LedgerName, d.OpBalance, c.VoucherName, 
			a.VrRoute AS Voucher_Route, a.Amount as VrAmt, isNull(b.DrAmt, 0) as DrAmt, isNull(b.CrAmt, 0) CrAmt, a.Narration, a.ref,  a.FileRef, a.Paymode, a.isCancelled, a.CompCode

	FROM FAS_VoucherTranMast a  LEFT OUTER JOIN FAS_VoucherTranSub b 	

		ON a.DocNo = b.DocNo AND a.DocDate=b.DocDate AND a.VrRoute = b.VrRoute AND a.CompCode = b.CompCode 

	INNER JOIN FAS_VoucherMaster c  

		ON a.VrRoute = c.VrRoute + c.VrWidth AND  a.CompCode = c.CompCode 

	LEFT OUTER JOIN FAS_LedgerMaster d 

		ON b.LrCode = d.LedgerCode AND a.CompCode = d.CompCode

--	WHERE a.isCancelled = '0'
GO

/****** Object:  View [dbo].[GroupKey]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE view [dbo].[GroupKey] as
select Distinct SentItemCode fromItemCode,RecdItemCode toItemCode from PUR_SUBCONT_ITEMMASTER
union all
select ReceivedItemCode fromItemCode,SentItemCode toItemCode from SAL_JobWorkEnquiry_Dtl
union all
select SentItemCode fromItemCode,RecdItemCode toItemCode from PUR_DismantlingOrder_Dtl
union all
select ReceivedItemCode fromItemCode,SentItemCode toItemCode from SAL_JobWorkOrder_Dtl
union all
select ReceivedItemCode fromItemCode,SentItemCode toItemCode from SAL_JobWork_ItemMaster
union all
select ReceivedItemCode fromItemCode,SentItemCode toItemCode from SAL_JobWorkQuotation_Dtl
union all
select SentItemCode fromItemCode,RecdItemCode toItemCode from PUR_LaborOrder_Dtl
union all
select RecdItemCode fromItemCode,  SentItemCode toItemCode from PUR_SubCont_ItemMaster
GO

/****** Object:  View [dbo].[HRDTrainingAnalysishdr]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[HRDTrainingAnalysishdr]
AS
SELECT DISTINCT DocType,DocNo,DocDate,GroupCode,DepartmentCode,DesignationCode,TeamCode,Revisionno,
			UserID,BranchID,CompID,Finyear,TranStamp
FROM  HRD_TrainingAnalysis_hdr a
WHERE (Revisionno = (Select  max(Revisionno) as Revisionno from HRD_TrainingAnalysis_hdr 
			Where a.DocType = DocType and a.DocNo = DocNo and a.DocDate = DocDate
			and a.GroupCode = GroupCode and a.DepartmentCode = DepartmentCode and a.DesignationCode = DesignationCode
			and a.TeamCode = TeamCode and a.BranchID = BranchID and a.CompID = CompID))
GO

/****** Object:  View [dbo].[HrRate]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

    CREATE VIEW [dbo].[HrRate] as
		Select a.itemcode, a.machinecode, a.cycletime, a.operationcost,
			b.machineHrrate, isnull(round((b.machineHrrate*a.cycletime)/3600.0,2),0) As OCost
		From BOM_ItemwiseMachinewiseCycleTime a 
		Left Outer Join MAINT_Equipment_Master b
			On a.MachineCode=b.EquipmentCode and a.branchID=b.branchID and a.compID=b.compID
GO

/****** Object:  View [dbo].[Inspection_Parameter_Perc]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[Inspection_Parameter_Perc]
as
select distinct ShortName,ParameterCode,Parameter,ItemCode,InstrumentType,DimenSion,case when isnumeric(FromValue)=0 then case when dimension='Angle' then dbo.cDegree(FromValue) else FromValue end else case when dimension='Angle' then dbo.cDegree(FromValue) else FromValue end +' / '+case when dimension='Angle' then dbo.cDegree(ToValue) else ToValue end  end Specification,
Mean,Defect,ValiDate,Sequence,(select InspectionTypeLongName  from Qc_InspectionType_Master d  where  x.InspectionTypeCode=d.InspectionTypeCode and x.Branchid=d.Branchid and x.Compid=d.Compid)InspectionTypeLongName,Branchid,Compid from (
select ShortName,ParameterCode,Parameter,ItemCode,(select InstrumentType from BOM_InstrumentType_Master where InstrumentTypeCode=a.InstruMentTypeCode  and Branchid=a.Branchid and Compid=a.Compid)InstrumentType,DimenSion,convert(varchar,round(convert(float,Mean)+(((convert(float,FromValue)-convert(float,Mean))*FromPerc/100)),3)) as FromValue,convert(varchar,round(convert(float,Mean)+(((convert(float,ToValue)-convert(float,Mean))*ToPerc/100)),3))  as ToValue,
Mean,Defect,case when ValiDate=1 then 'YES' else 'NO'end ValiDate,Sequence,a.InspectionTypeCode,a.Branchid,a.Compid
from Qc_Inspection_Master a where a.Inspect=1and isnumeric(a.FromValue)=1
union
select ShortName,ParameterCode,Parameter,ItemCode,(select InstrumentType from BOM_InstrumentType_Master where InstrumentTypeCode=a.InstruMentTypeCode  and Branchid=a.Branchid and Compid=a.Compid)InstrumentType,DimenSion,FromValue,ToValue,
Mean,Defect,case when ValiDate=1 then 'YES' else 'NO'end ValiDate,Sequence,a.InspectionTypeCode,a.Branchid,a.Compid
from Qc_Inspection_Master a where a.Inspect=1and isnumeric(a.FromValue)=0 )x
GO

/****** Object:  View [dbo].[INV_Material_Master_All]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE View [dbo].[INV_Material_Master_All] 
AS
Select Distinct a.ItemID, a.ItemCode, a.ItemDescription, a.ItemNature, a.ProductCategoryCode, a.ProductSubCategoryCode, a.PurchaseNature,
	a.Origin, a.Boughtout, a.ItemSpecification, a.PostingAccount, a.PostingAccountEditable, a.CostCenter, a.UOM, a.PurchaseUOM,
	a.PurConversionFactor, a.IssueUOM, a.IssConversionFactor, a.MinimumStock, a.MaximumStock, a.ReOrderLevel, a.ReOrderQty,
	a.SafetyStock, a.MinLeadTime, a.MaxLeadTime, a.SelfLifePeriod, a.Rate, a.Modvat, a.TariffNo, a.HSNId, a.isAsset,
	a.TolerancePlusApplicable, a.PurchaseTolerancePlus, a.ToleranceMinusApplicable, a.PurchaseToleranceMinus,
	a.PackingType, a.StdPackingQty, a.DrawingNo, a.DrawingDescription, a.RevisionNo, a.RevisionDate, a.DrawingPath,
	Case IsNull(a.IsHeatNoApp, 0) When 0 Then 'No' Else 'Yes' End as HeatNoApplicable,
	Case IsNull(a.IsShelfLifePeriodApp, 0) When 0 Then 'No' Else 'Yes' End as ShelfLifePeriodApplicable,
	a.TrayTypeCode, a.Status, a.UserID, a.BranchID, a.CompID, a.TranStamp
From INV_Material_Master a
UNION
Select Distinct b.ItemID, (a.ItemCode+a.AssetCode) As ItemCode, b.ItemDescription, b.ItemNature, b.ProductCategoryCode, b.ProductSubCategoryCode,
	b.PurchaseNature, b.Origin, b.Boughtout, b.ItemSpecification, b.PostingAccount, b.PostingAccountEditable, b.CostCenter, b.UOM,
	b.PurchaseUOM, b.PurConversionFactor,b.IssueUOM, b.IssConversionFactor, b.MinimumStock, b.MaximumStock, b.ReOrderLevel, b.ReOrderQty,
	b.SafetyStock, b.MinLeadTime, b.MaxLeadTime, b.SelfLifePeriod, b.Rate, b.Modvat, b.TariffNo, b.HSNId, b.isAsset,
	b.TolerancePlusApplicable, b.PurchaseTolerancePlus, b.ToleranceMinusApplicable, b.PurchaseToleranceMinus,
	b.PackingType, b.StdPackingQty, b.DrawingNo, b.DrawingDescription, b.RevisionNo, b.RevisionDate,b.DrawingPath,
	Case IsNull(b.IsHeatNoApp, 0) When 0 Then 'No' Else 'Yes' End as HeatNoApplicable,
	Case IsNull(b.IsShelfLifePeriodApp, 0) When 0 Then 'No' Else 'Yes' End as ShelfLifePeriodApplicable,
	b.TrayTypeCode, b.Status,b.UserID,b.BranchID,b.CompID, b.TranStamp
From FAM_Asset_Master a
Inner Join INV_Material_Master b
	on a.ItemCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID
GO

/****** Object:  View [dbo].[INV_Stock_Master]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[INV_Stock_Master] as   
  
Select a.ItemCode, c1.WarehouseCode, c1.RackBinCode,a.Branchid,a.CompID,   
isNull(Sum(Case When (b.TransactionItem=a.ItemCode and b.FromWarehouseCode=c1.WarehouseCode) Then -b.TransactionQty   
  When (b.TransactionItem=a.ItemCode and b.ToWarehouseCode=c1.WarehouseCode) Then b.TransactionQty End), 0) as CBQty   
from INV_MATERIAL_MASTER_all a   
Full Outer Join INV_MATERIALTRAN_DTL b On a.Branchid = b.Branchid and a.CompID=b.CompID    
  
Full Outer Join   
(Select c.WareHouseCode, isNull(d.RackBinCode, '') as RackBinCode, c.BranchID, c.CompID   
 from INV_WAREHOUSE_MASTER c Left Outer Join INV_RACKBIN_MASTER d  
 on d.WarehouseCode = c.WarehouseCode and d.Branchid = c.Branchid and d.Compid = c.Compid  
) c1  
  On a.Branchid = c1.Branchid and a.CompID=c1.CompID  
  
Where b.TransactionItem=a.ItemCode   
and ((b.FromWarehouseCode=c1.WarehouseCode and b.FromRackBinCode=c1.RackBinCode) or (b.ToWarehouseCode=c1.WarehouseCode and b.ToRackBinCode=c1.RackBinCode))  
and convert(varchar(10),b.TransactionDate,120) + ' ' + convert(varchar(8),b.TransactionTime,108) <=GetDate()  
group by a.ItemCode, c1.WarehouseCode,c1.RackBinCode,a.Branchid,a.CompID  
  
GO

/****** Object:  View [dbo].[INV_Stock_Master_Tmp]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[INV_Stock_Master_Tmp] as 

Select a.ItemCode, c1.WarehouseCode, c1.RackBinCode,a.Branchid,a.CompID, 
isNull(Sum(Case When (b.TransactionItem=a.ItemCode and b.FromWarehouseCode=c1.WarehouseCode) Then -b.TransactionQty 
		When (b.TransactionItem=a.ItemCode and b.ToWarehouseCode=c1.WarehouseCode) Then b.TransactionQty End), 0) as CBQty 
from INV_MATERIAL_MASTER_all a 
Full Outer Join INV_MATERIALTRAN_DTL b On a.Branchid = b.Branchid and a.CompID=b.CompID  

Full Outer Join 
(Select c.WareHouseCode, isNull(d.RackBinCode, '') as RackBinCode, c.BranchID, c.CompID 
	from INV_WAREHOUSE_MASTER c Left Outer Join INV_RACKBIN_MASTER d
	on d.WarehouseCode = c.WarehouseCode and d.Branchid = c.Branchid and d.Compid = c.Compid
) c1
  On a.Branchid = c1.Branchid and a.CompID=c1.CompID

Where b.TransactionItem=a.ItemCode 
and ((b.FromWarehouseCode=c1.WarehouseCode and b.FromRackBinCode=c1.RackBinCode) or (b.ToWarehouseCode=c1.WarehouseCode and b.ToRackBinCode=c1.RackBinCode))
and convert(varchar(10),b.TransactionDate,120) + ' ' + convert(varchar(8),b.TransactionTime,108) <='2017-07-10 09:50:00'		--GetDate()
group by a.ItemCode, c1.WarehouseCode,c1.RackBinCode,a.Branchid,a.CompID
GO

/****** Object:  View [dbo].[INV_Warehouse_Master_All]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE View [dbo].[INV_Warehouse_Master_All] As

Select * from INV_Warehouse_Master

Union

Select * from INV_WarehouseTemp_Master
GO

/****** Object:  View [dbo].[JobworkEnquiry]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[JobworkEnquiry]  
AS  
	SELECT a.EnquiryType, a.EnquiryNo, a.EnquiryDate, a.ReferenceNo,a.RevisionNo, a.ReferenceDate, a.CustomerCode, a.CustomerName, a.Remarks, b.ItemGroupKey,   
		b.ReceivedItemCode, b.RecdDescription, b.SentItemCode, b.SentDescription, b.SendQty, b.RatePer, b.Rate, b.AssRate, b.QuotationType,   
		b.QuotationNo, b.QuotationDate, a.Branchid, a.Compid, a.Finyear, b.RelationType, b.Status, b.isActive, b.StatusDate, b.CreatedDate  
	FROM dbo.SAL_JobworkEnquiry_Hdr a 
	LEFT OUTER JOIN dbo.SAL_JobworkEnquiry_Dtl b
		ON a.EnquiryType = b.EnquiryType AND a.EnquiryNo = b.EnquiryNo AND a.EnquiryDate = b.EnquiryDate AND   
		a.Branchid = b.BranchID AND a.Compid = b.CompID AND a.Finyear = b.Finyear
GO

/****** Object:  View [dbo].[JobworkOrder]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[JobworkOrder]
AS
SELECT     a.OrderType, a.OrderNo, a.OrderDate, a.OrderSource, a.CustomerCode, a.ReferenceNo, a.ReferenceDate, a.ValidFrom, a.ValidTo, a.Remarks, 
                      a.FormCode, a.PackingCode, a.DeliveryCode, a.PaymentCode, a.InsuranceCode, a.ModeCode, a.Freightcode, a.FormulaCode, a.TotalAmount, 
                      a.OrderAuthorization, a.OrderAuthorizationDate, a.OrderAuthorizationBy, b.ItemGroupKey, b.ReceivedItemCode, b.SentItemCode, a.OrderingType, 
                      b.OrderQty, b.RatePer, b.Rate, b.AssRate, a.Status, b.Status AS ItemStatus, a.Branchid, a.Compid, a.Finyear, a.transtamp
FROM         dbo.SAL_JobworkOrder_Hdr a LEFT OUTER JOIN
                      dbo.SAL_JobworkOrder_Dtl b ON a.OrderType = b.OrderType AND a.OrderNo = b.OrderNo AND a.OrderDate = b.OrderDate AND 
                      a.Branchid = b.Branchid AND a.Compid = b.Compid AND a.Finyear = b.Finyear
GO

/****** Object:  View [dbo].[JobworkOrderAmendment]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE view [dbo].[JobworkOrderAmendment] as 
select a.AmendmentType, a.AmendmentNo, a.AmendmentDate, a.CustomerCode, a.OrderType, 
a.OrderNo, a.OrderDate, a.FormCode, a.PackingCode, a.DeliveryCode, a.PaymentCode, 
a.InsuranceCode, a.ModeCode, a.Freightcode, a.FormulaCode, a.TotalAmount, 
a.AmendmentAuthorization, a.AmendmentAuthorizationDate, a.AmendmentAuthorizationBy,
b.ItemGroupKey, b.ReceivedItemCode, b.SentItemCode, b.Qty, b.NewQty, b.RatePer, b.Rate, b.NewRate, b.AssRate, 
b.NewAssRate, b.ValidFrom, b.Status, a.Branchid, a.Compid, a.Finyear 
from SAL_JOBWORKAMENDMENT_Hdr a 
left outer join SAL_JOBWORKAMENDMENT_DTL b
on a.AmendmentType = b.AmendmentType and a.AmendmentNo = b.AmendmentNo 
and a.AmendmentDate = b.AmendmentDate and a.Branchid = b.Branchid and a.Compid = b.Compid
and a.Finyear = b.Finyear



GO

/****** Object:  View [dbo].[JobworkQuotation]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[JobworkQuotation]
AS
SELECT     a.QuotationType, a.QuotationNo, a.QuotationDate, a.CustomerCode, a.CustomerName, a.ValidFrom, a.ValidTo, a.FormulaCode, a.Remarks, a.FormCode,
                       a.PackingCode, a.DeliveryCode, a.PaymentCode, a.InsuranceCode, a.ModeCode, a.Freightcode, a.TotalAmount, b.ItemGroupKey, b.ReceivedItemCode, 
                      b.SentItemCode, b.SendQty, b.Rate, b.AssRate, b.OrderType, b.OrderNo, b.OrderDate, a.Branchid, a.Compid, a.Finyear, b.RatePer, 
											IsNull(b.Revisionno,0) As Revisionno, IsNull(b.SentDesc,'') As SentDesc, IsNull(b.ReceivedDesc,'') As ReceivedDesc
FROM         dbo.SAL_JobworkQuotation_Hdr a INNER JOIN
                      dbo.SAL_JOBWORKQUOTATION_DTL b ON a.QuotationType = b.QuotationType AND a.QuotationNo = b.QuotationNo AND 
                      a.QuotationDate = b.QuotationDate AND a.Branchid = b.Branchid AND a.Compid = b.Compid AND a.Finyear = b.Finyear
GO

/****** Object:  View [dbo].[LabourOrder]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


create view [dbo].[LabourOrder] as 
select a.OrderType, a.OrderNo, a.OrderDate, a.VendorCode, a.Source, a.ValidFrom, a.ValidTo,
 a.CreditDays, a.CreditFrom, a.AdvanceOn, a.OrderingType, a.FormulaCode, a.PostingAc, 
a.FormCode, a.PackingCode, a.DeliveryCode, a.PaymentCode, a.InsuranceCode, a.ModeCode,
 a.Freightcode, a.OrderValue, a.Narration, a.SplInstruction, a.AuthBy, a.Status, 
a.StatusDate, a.EffStatus , a.EffDate, a.CreatedDate,
b.Seq, b.ItemGroupKey, b.SentItemCode, b.RecdItemCode, b.ToBeRecdQty, b.RatePer, b.Rate,
 b.AssRate, b.RateEffDate, b.QuotationType, b.QuotationNo, b.QuotationDate, b.isActive,
 b.Status ItemStatus, b.StatusDate ItemStatusDate , b.EffStatus ItemEffStatus, b.EffDate ItemEffStatusDate, b.CreatedDate ItemCreatedDate, a.BranchID, a.CompID, 
a.Finyear from PUR_LaborOrder_Hdr a  Left Outer Join PUR_LaborOrder_Dtl b
		On b.OrderType=a.OrderType and b.OrderNo=a.OrderNo and b.OrderDate=a.OrderDate and b.BranchID=a.BranchID and b.CompID=a.CompID


GO

/****** Object:  View [dbo].[MATERIAL_MASTER]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE View [dbo].[MATERIAL_MASTER]         
As         
Select a.ItemID, a.itemcode, a.itemdescription, a.ItemSpecification, a.UOM, b.sdesc As UOMDesc, b.NoofDecimals As UOMDecimal,
	c.sdesc As PurchaseUOMDesc, c.NoofDecimals As PurchaseUOMDecimal, a.PurConversionFactor,
	d.sdesc As IssueUOMDesc, d.NoofDecimals As IssueUOMDecimal, a.IssConversionFactor As IssueConv,
	a.itemnature, a.Boughtout, a.MinimumStock, a.MaximumStock, a.ReOrderLevel, a.SafetyStock, a.Rate, a.isAsset, a.Status,
	(Case IsNull(IsHeatNoApp, 0) When 0 Then 'No' Else 'Yes' End) As HeatNoApplicable,
	(Case IsNull(IsShelfLifePeriodApp, 0) When 0 Then 'No' Else 'Yes' End) As ShelfLifePeriodApplicable,
	IsNull(a.SelfLifePeriod,0) As SelfLifePeriod, a.DrawingNo,a.TrayTypeCode,a.TariffNo,a.HSNId,
	a.Revisionno, a.RevisionDate, a.BranchID, a.CompID
From Inv_Material_Master a
Left Outer Join Inv_UOM_Master b
	On b.code=a.uom and b.BranchID=a.BranchID and b.CompID=a.CompID
Left Outer Join INV_UOM_Master c 
	On c.code=a.purchaseuom and c.BranchID=a.BranchID and c.CompID=a.CompID
Left Outer Join INV_UOM_Master d 
	On d.code=a.issueuom and d.BranchID=a.BranchID and d.CompID=a.CompID
GO

/****** Object:  View [dbo].[MaterialReturnNote]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE view [dbo].[MaterialReturnNote] as 
select a.MaterialReturnNoteType,a.MaterialReturnNoteNo,a.MaterialReturnNoteDate,
a.StoreRequestType,a.StoreRequestNo,a.StoreRequestDate,b.ItemCode,b.WorkInProgressQty,
b.ReturnQty,b.ApprovedQty,b.StockReturnQty,b.LineRejectionQty,b.ScrapQty,b.Remarks, b.RejectionType, b.ReceiptYN, b.QCFlg,
a.Branchid,a.Compid,a.Finyear from PPC_MaterialReturnNote_HDR a,PPC_MaterialReturnNote_DTL b
where a.MaterialReturnNoteType=b.MaterialReturnNoteType 
and a.MaterialReturnNoteNo=b.MaterialReturnNoteNo 
and a.MaterialReturnNoteDate=b.MaterialReturnNoteDate
and a.Branchid=b.Branchid and a.Compid=b.Compid and a.Finyear=b.Finyear

GO

/****** Object:  View [dbo].[MPSMaster]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

	CREATE view [dbo].[MPSMaster] as   
		(select a.ItemCode, a.StageCode, a.OperationCode,  b.OperationName, b.OperationSpecification,  
		case when a.OperationSequence = 0 then   
		(select max(b.OperationSequence)+1 from BOM_MPS_Master b where b.ItemCode = a.ItemCode  
		and b.Branchid = a.Branchid and b.Compid = a.Compid and b.RevisionNo = a.RevisionNo and Status=1)  
		else a.OperationSequence end OperationSequence, a.OperationSequence Seq, a.RevisionNo,a.Status,a.Branchid, a.Compid  
	from BOM_MPS_Master a  
	left outer join BOM_Operation_Master b 
		on a.OperationCode = b.OperationCode and a.Branchid = b.Branchid and a.Compid = b.Compid   
	left outer join INV_Material_Master d 
		on d.ITemCode=a.StageCode and a.BRanchid=d.BRanchid and a.Compid=d.Compid  
	where a.RevisionNo = (select Max(c.RevisionNo) from BOM_MPS_Master c  
					where c.ItemCode = a.ItemCode and c.Branchid = a.Branchid and c.Compid = a.Compid) 
		and d.Status=1  
	group by a.ItemCode,a.StageCode,a.OperationCode,  b.OperationName, a.OperationSequence, b.OperationSpecification,  
	a.RevisionNo,a.Status,a.Branchid, a.Compid)
GO

/****** Object:  View [dbo].[NCDC]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO







CREATE view [dbo].[NCDC] as 
select a.NCDCType,a.NCDCNo,a.NCDCDate,a.VendorCode,a.TransporterName,a.Returnable,
b.ItemCode,b.SendQty,b.ReceivedQty,b.ScheduleDate,b.Remarks,
a.Branchid,a.Compid,a.Finyear from INV_NCDC_HDR a, INV_NCDC_DTL b
where a.NCDCType=b.NCDCType and a.NCDCNo=b.NCDCNo and a.NCDCDate=b.NCDCDate 
and a.Branchid=b.Branchid and a.Compid=b.Compid and a.Finyear=b.Finyear














GO

/****** Object:  View [dbo].[OtherDetails_Drawing]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[OtherDetails_Drawing] as
select distinct b.PartyName as CustomerName,a.CustomerCode,d.ChildCode,a.PartNo,d.Priority,
c.StageCode,a.ItemCode,f.DrawingNo RefNo,f.RevisionNo,f.DrawingCode,f.Scale,
'' as OperationName,
e.ItemDescription,f.RevisionDate,a.Branchid,a.Compid from Sal_Customer_ItemMaster a
left outer join FAS_Party_Master b
on a.CustomerCode=b.PartyCode and a.Branchid=b.Branchid and a.Compid=b.Compid
left outer join BOM_Mps_Master c on a.ItemCode=c.ItemCode and a.Branchid=c.Branchid and a.Compid=c.Compid
and c.OperationSequence=(Select Distinct OperationSequence from BOM_Mps_Master where ItemCode=a.ItemCode and StageCode=c.StageCode)
left outer join BOM_Structure_Dtl d on c.StageCode=d.ParentCode and d.Branchid=c.Branchid and d.Compid=c.Compid
left outer join INV_Material_Master e on e.ItemCode=c.StageCode and e.Branchid=c.Branchid and e.Compid=c.Compid
left outer join (select a.DrawingNo,a.RevisionNo,a.RevisionDate,a.DrawingCode,a.Scale,
a.ItemCode,a.Branchid,a.Compid
from BOM_Drawing_Master a where a.RevisionNo = (select max(b.RevisionNo) from Bom_Drawing_Master b
where a.ItemCode= b.ItemCode and b.Branchid=a.Branchid and b.Compid=a.Compid)) f on f.ItemCode=e.ItemCode and f.Branchid=a.Branchid and f.Compid=a.Compid
where d.Priority=1
GO

/****** Object:  View [dbo].[OtherDetails_Drawing_RM]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[OtherDetails_Drawing_RM]
as
select distinct b.PartyName as CustomerName,a.CustomerCode,d.ChildCode,a.PartNo,d.Priority,
c.StageCode,a.ItemCode,f.DrawingNo RefNo,f.RevisionNo,f.DrawingCode,f.Scale,
(select OperationName from BOM_Operation_Master where OperationCode=c.OperationCode)OperationName,
e.ItemDescription,f.RevisionDate,a.Branchid,a.Compid from Sal_Customer_ItemMaster a
left outer join FAS_Party_Master b
on a.CustomerCode=b.PartyCode and a.Branchid=b.Branchid and a.Compid=b.Compid
left outer join BOM_Mps_Master c on a.ItemCode=c.ItemCode and a.Branchid=c.Branchid and a.Compid=c.Compid
and c.OperationSequence=(Select Distinct OperationSequence from BOM_Mps_Master where ItemCode=a.ItemCode and StageCode=c.StageCode)
left outer join BOM_Structure_Dtl d on (Select Distinct StageCode from BOM_Mps_Master where 
ItemCode=a.ItemCode and OPerationSequence=1)=d.ParentCode and d.Branchid=c.Branchid and d.Compid=c.Compid
left outer join INV_Material_Master e on e.ItemCode=d.ChildCode and e.Branchid=c.Branchid and e.Compid=c.Compid
left outer join (select a.DrawingNo,a.RevisionNo,a.RevisionDate,a.DrawingCode,a.Scale,
a.ItemCode,a.Branchid,a.Compid
from BOM_Drawing_Master a where a.RevisionNo = (select max(b.RevisionNo) from Bom_Drawing_Master b
 where a.ItemCode= b.ItemCode and b.Branchid=a.Branchid and b.Compid=a.Compid)) f on f.ItemCode=e.ItemCode and f.Branchid=a.Branchid and f.Compid=a.Compid
where d.Priority=1
GO

/****** Object:  View [dbo].[OtherDetails_Drawing1]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE view [dbo].[OtherDetails_Drawing1]  
as  
select distinct b.PartyName as CustomerName,a.CustomerCode,d.ChildCode,a.PartNo,d.Priority,  
c.StageCode,a.ItemCode,f.DrawingNo RefNo,f.RevisionNo,f.DrawingCode,f.Scale,f.Type,  
(select OperationName from BOM_Operation_Master where OperationCode=c.OperationCode)OperationName,  
e.ItemDescription,f.RevisionDate,a.Branchid,a.Compid from Sal_Customer_ItemMaster a  
left outer join FAS_Party_Master b  
on a.CustomerCode=b.PartyCode and a.Branchid=b.Branchid and a.Compid=b.Compid  
left outer join MpsMaster c on a.ItemCode=c.ItemCode and a.Branchid=c.Branchid and a.Compid=c.Compid  
and c.OperationSequence=(Select Distinct OperationSequence from MpsMaster where ItemCode=a.ItemCode and StageCode=c.StageCode)  
left outer join BOM_Structure_Dtl d on (Select Distinct StageCode from MpsMaster where   
ItemCode=a.ItemCode and OPerationSequence=1)=d.ParentCode and d.Branchid=c.Branchid and d.Compid=c.Compid  
left outer join INV_Material_Master e on e.ItemCode=c.StageCode and e.Branchid=c.Branchid and e.Compid=c.Compid  
left outer join (select x.DrawingNo,x.RevisionNo,x.RevisionDate,x.DrawingCode,x.Scale,x.Type,  
x.ItemCode,x.Branchid,x.Compid  
from BOM_Drawing_Master x where x.RevisionNo = (select max(b.RevisionNo) from Bom_Drawing_Master b  
 where x.ItemCode= b.ItemCode and b.Type=x.Type and x.Branchid=b.Branchid and x.Compid=b.Compid )) f   
on f.ItemCode=e.ItemCode and f.Branchid=e.Branchid and f.Compid=e.Compid  
where d.Priority=1 and e.status=1  
  
GO

/****** Object:  View [dbo].[Pap_Vw_DocRevision]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create view [dbo].[Pap_Vw_DocRevision] as
select * from PAP_Document_Revision a where Revisionno=
(Select Max(Revisionno) from PAP_Document_Revision where DocType=a.DocType and DocNo=a.DocNo and DocDate=a.DocDate
and Branchid=a.BRanchid and Compid=a.Compid)
GO

/****** Object:  View [dbo].[Parameter]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE view [dbo].[Parameter]
as
select distinct inspectionTypeCode, ShortName,ParameterCode,Parameter,ItemCode,InstrumentType,DimenSion,case when isnumeric(FromValue)=0 then  case when dimension='Angle' then dbo.cDegree(FromValue) else FromValue end else case when dimension='Angle' then dbo.cDegree(FromValue) else FromValue end +' / '+case when dimension='Angle' then dbo.cDegree(ToValue) else ToValue end  end Specification,
case x.Defect when 'IdentChars' then '<IC>' when 'Significant' then '<SC>' when 'Major' then '' when 'Critical' then '<CC>' else '' end SplChar,Mean,Defect,ValiDate,Sequence,Type,Branchid,Compid from (
select inspectionTypeCode,ShortName,ParameterCode,Parameter,ItemCode,(select InstrumentType from BOM_InstrumentType_Master where InstrumentTypeCode=a.InstruMentTypeCode)InstrumentType,DimenSion,FromValue,ToValue,
Mean,Defect,case when ValiDate=1 then 'YES' else 'NO'end ValiDate,Sequence,Type,a.Branchid,a.Compid
from Qc_Inspection_Master a where a.Status=1 and  Inspect=1 )x
GO

/****** Object:  View [dbo].[Parameter1]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE view [dbo].[Parameter1]
as
select distinct ShortName,ParameterCode,Parameter,ItemCode,InstrumentType,DimenSion,case when Type='SP' then case when FromValue=ToValue then FromValue else FromValue+' To ' +toValue end else
case when isnumeric(FromValue)=0 then case when dimension='Angle' then dbo.cDegree(FromValue) else FromValue end else case when dimension='Angle' then dbo.cDegree(FromValue) else FromValue end +' / '+case when dimension='Angle' then dbo.cDegree(ToValue) else ToValue end  end end Specification,
Mean,Defect,ValiDate,Type,Sequence,Branchid,Compid,
case x.Defect when 'IdentChars' then '<IC>' when 'Significant' then '<SC>' when 'Major' then '' when 'Minor' then ''when 'Critical' then '<CC>' else '' end SplChar  from (
select ShortName,ParameterCode,Parameter,ItemCode,(select InstrumentType from BOM_InstrumentType_Master where InstrumentTypeCode=a.InstruMentTypeCode)InstrumentType,DimenSion,FromValue,ToValue,
Mean,Defect,Type,case when ValiDate=1 then 'YES' else 'NO'end ValiDate,Sequence,a.Branchid,a.Compid
from BOM_Inspection_Parameter a where a.Status=1 and a.Exption=1)x
GO

/****** Object:  View [dbo].[PartMast]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[PartMast]

AS

---Select Top 1 * 
---from
	---(

	Select top 100 Percent a.ItemCode, a.StageCode,  b.ItemDescription, 

		isNull(
				(
					Select Top 1 b1.CustomerCode 
					From SAL_CUSTOMER_ITEMMASTER b1 
					Where b1.ItemCode = a.ItemCode and b1.BranchID = a.BranchID and b1.CompID = a.CompID
				),
				(
					Select Top 1 b1.CustomerCode 
					From SAL_JOBWORK_ITEMMASTER b1 
					Where b1.SentItemCode = a.ItemCode and b1.BranchID = a.BranchID and b1.CompID = a.CompID
				)
			) as CustomerCode,

		isNull(
				(
					Select Top 1 b1.PartNo 
					From SAL_CUSTOMER_ITEMMASTER b1 
					Where b1.ItemCode = a.ItemCode and b1.BranchID = a.BranchID and b1.CompID = a.CompID
				),
				(
					Select Top 1 b1.PartNo 
					From SAL_JOBWORK_ITEMMASTER b1 
					Where b1.SentItemCode = a.ItemCode and b1.isActive='1' and b1.BranchID = a.BranchID and b1.CompID = a.CompID
				)
			) as PartNo,  a.BranchID,  a.CompID
	From Bom_Mps_Master a 
	Inner Join Inv_Material_Master b
		On a.StageCode = b.ItemCode and a.BranchID = b.BranchID and a.CompID = b.CompID

/*
	
	Union All
	
	Select top 100 Percent a.ItemCode, a.StageCode,  b.ItemDescription,  a.BranchID,  a .CompID, 
		(Select Top 1 PartNo from SAL_JOBWORK_ITEMMASTER b Where b.SentItemCode = a.ItemCode 
			and BranchID = a.BranchID and CompID = a.CompID) as PartNo
	from Bom_Mps_Master a Inner Join Inv_Material_Master b
	On a.StageCode = b. ItemCode and a.BranchID = b.BranchID and a.CompID = b.CompID
*/
	---) x

GO

/****** Object:  View [dbo].[PartyMaster]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE View [dbo].[PartyMaster] 

AS

	Select a.PartyCode, a.PartyName, a.AcType, a.IsSubcontractor, a.EffStatus, a.EffDate, a.CreditDays,
		a.PackingCode,
			isNull((select b1.Description from TERMSMASTER b1 
			where b1.TypeFlg='Packing & Forwarding' and b1.Code=a.PackingCode and b1.Branchid=a.Branchid and b1.Compid=a.Compid), '') PackingDesc,
		a.InsuranceCode,
			isNull((select b4.Description from TERMSMASTER b4 
			where b4.TypeFlg='Insurance' and b4.Code=a.InsuranceCode and b4.Branchid=a.Branchid and b4.Compid=a.Compid), '') InsuranceDesc,
		a.DeliveryCode,
			isNull((select b2.Description from TERMSMASTER b2 
			where b2.TypeFlg='Delivery Instruction' and b2.Code=a.DeliveryCode and b2.Branchid=a.Branchid and b2.Compid=a.Compid), '') DeliveryDesc,
		a.PaymentCode,
			isNull((select b3.Description from TERMSMASTER b3 
			where b3.TypeFlg='Payment Terms' and b3.Code=a.PaymentCode and b3.Branchid=a.Branchid and b3.Compid=a.Compid), '') PaymentDesc,
		a.Freightcode,
			isNull((select b6.Description from TERMSMASTER b6 
			where b6.TypeFlg='Freight' and b6.Code=a.FreightCode and b6.Branchid=a.Branchid and b6.Compid=a.Compid), '') FreightDesc,
		a.ModeCode,
			isNull((select b5.Description from TERMSMASTER b5 
			where b5.TypeFlg='Mode' and b5.Code=a.ModeCode and b5.Branchid=a.Branchid and b5.Compid=a.Compid), '') ModeDesc,
		a.FormCode,
			isNull((select b.Description from TERMSMASTER b 
			where b.TypeFlg='Form' and b.Code=a.FormCode and b.Branchid=a.Branchid and b.Compid=a.Compid), '') FormDesc, 
		a.CurrencyCode,
			isNull((Select b9.CurrencyName from COM_Currency_Master b9 where b9.CurrencyCode=a.CurrencyCode), '') CurrencyName,
		a.CustomerFormulaCode,
			isNull((Select b7.FrmLName from Com_FormulaMast_Hdr b7
			where b7.FrmID=a.CustomerFormulaCode and b7.ModuleCode = 4  and b7.BranchID=a.BranchID and b7.CompID=a.CompID), '') CustomerFormulaName,
		a.VendorFormulaCode,
			isNull((Select b8.FrmLName from Com_FormulaMast_Hdr b8
			where b8.FrmID=a.VendorFormulaCode and b8.ModuleCode = 9 and b8.BranchID=a.BranchID and b8.CompID=a.CompID), '') VendorFormulaName,

		a.BranchID, a.CompID
	from FAS_Party_Master a
GO

/****** Object:  View [dbo].[PAY_PFESILimit_vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

Create View [dbo].[PAY_PFESILimit_vw]   
AS  
Select * from PAY_PFESILimit a  
Where a.TranStamp=(Select Max(TranStamp)   
                    From PAY_PFESILimit b Where b.EmployeeCode=a.EmployeeCode and b.BranchID=a.BranchID and b.CompID=a.CompID)



GO

/****** Object:  View [dbo].[PAY_ProcessResult]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


Create View [dbo].[PAY_ProcessResult] as 
select a.ProcessMonth, a.EmployeeCode, a.ExprSeq, a.ExprGroupId, b.ExprGroupName, 
a.ResultValue, a.GrandTotalFlag, a.ProcessType, a.Branchid, a.Compid from 
pay_formulacalc_hdr a left outer join PAY_FormulaExpr_Group b 
on a.ExprGroupId = b.ExprGroupId and a.Branchid = b.Branchid and a.Compid = b.Compid
GO

/****** Object:  View [dbo].[PAYExpressionGrouping]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO





CREATE View [dbo].[PAYExpressionGrouping] AS

--Active Parent Childs
Select a.ExprGroupID, a.ExprGroupName, a.Route, a.Width, a.isActive, Convert(bit, '1') as ByParentStatus, a.BranchID, a.CompID
from PAY_FormulaExpr_Group a 
where exists 
(Select * from PAY_FormulaExpr_Group b 
where b.Route+b.Width like Left(a.Route, 2) and Len(b.Route+b.Width) <= Len(a.Route) and b.isActive='1'
and b.BranchID=a.BranchID and b.CompID=a.CompID)

union all
--Inactive Parent Childs
Select a.ExprGroupID, a.ExprGroupName, a.Route, a.Width, a.isActive, Convert(bit, '0') as ByParentStatus, a.BranchID, a.CompID
from PAY_FormulaExpr_Group a 
where exists 
(Select * from PAY_FormulaExpr_Group b 
where b.Route+b.Width like Left(a.Route, 2) and Len(b.Route+b.Width) <= Len(a.Route) and b.isActive='0'
and b.BranchID=a.BranchID and b.CompID=a.CompID)

union all
--Root Elements
Select a.ExprGroupID, a.ExprGroupName, a.Route, a.Width, a.isActive, Convert(bit, '1') as ByParentStatus, a.BranchID, a.CompID
from PAY_FormulaExpr_Group a where a.Route=''





GO

/****** Object:  View [dbo].[PFMEA]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[PFMEA] as
select Distinct a.* from BOM_PFMEA a where RevisionNo=
(Select Max(RevisionNo) from BOM_PFMEA where PFMEAType=a.PFMEAType and PFMEANo=a.PFMEANo and 
PFMEADate=a.PFMEADate and Branchid=a.Branchid and Compid=a.Compid)
GO

/****** Object:  View [dbo].[planqty1]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


create view [dbo].[planqty1] as
select a.itemcode,sum(a.planqty)planqty from SAL_SalesPlan a  
where a.planperiod between '2007-08-01' and '2007-08-31'  and planfor='2'
group by a.itemcode
GO

/****** Object:  View [dbo].[PPC_Vw_PRocessWorkOrder]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[PPC_Vw_PRocessWorkOrder] as
select Distinct ProcessWorkOrderType,ProcessWorkOrderNo,ProcessWorkOrderDate,ParentCode,WorkOrderQty,FromStageCode,ToStageCode,CompletedQty,RejectionQty,ScrapQty,
isnull((select OperationSequence from MPS_Process where ItemCode=d.ParentCode 
and StageCode=d.FromStageCode and Compid=d.Compid and Branchid=d.BRanchid),0)FromSeq,
isnull((select OperationSequence from MPS_Process where ItemCode=d.ParentCode 
and StageCode=d.ToStageCode and Compid=d.Compid and Branchid=d.BRanchid),0)ToSeq,d.BRanchid,d.Compid,d.Finyear
from PPC_PRocessWorkOrder d where CompletedStatus=0 and WoShortClose=0

GO

/****** Object:  View [dbo].[PPC_Vw_QuickProductionplan_Dtl]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE VIEW [dbo].[PPC_Vw_QuickProductionplan_Dtl] as  
  
Select * from PPC_QuickProductionplan_Dtl a  
Where a.RevisionNo=(Select Max(RevisionNo) From PPC_QuickProductionplan_Dtl a1          
    Where a.DocNo = a1.DocNo and a.DocType = a1.DocType and a.DocDate = a1.DocDate and a.StageCode = a1.Stagecode and a.MachineCode = a1.MachineCode and a.Shift = a1.Shift and a.Branchid=a1.BranchID   
    and a.Compid=a1.CompID and a.Finyear = a1.Finyear)     
GO

/****** Object:  View [dbo].[Prcess_Flow_Diagram]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE view [dbo].[Prcess_Flow_Diagram]
as
select distinct b.PartyName as CustomerName,a.CustomerCode,a.Model,a.IssuedNo,a.PartNo,e.ItemDescription PartName,a.ItemCode,
e.ItemDescription,f.ProcessType,f.ProcessNo,f.ProcessDate,f.EmployeeCode,
isnull((Select Max(Revisionno) from PAP_Document_Revision where DocType = f.ProcessType And DocNo = f.ProcessNo And DocDate = f.ProcessDate and Branchid=a.Branchid and Compid=a.Compid),0)as RevisionNo,
isnull((Select RevisionDate from PAP_Document_Revision where DocType = f.ProcessType And DocNo = f.ProcessNo And DocDate = f.ProcessDate and Branchid=f.Branchid and Compid=f.Compid and Revisionno=isnull((Select Max(Revisionno) from PAP_Document_Revision where DocType = f.ProcessType And DocNo = f.ProcessNo And DocDate = f.ProcessDate and Branchid=f.Branchid and Compid=f.Compid),0)),f.ProcessDate) as RevisionDate,
(Select EmployeeName from Pay_Employee_Master where EmployeeCode=f.EmployeeCode and Branchid=f.Branchid and Compid=f.Compid)EmployeeName,b.SuplierCode as VendorCode,
(Select EmployeeName from Pay_Employee_Master where EmployeeCode=f.Approvedby and Branchid=f.Branchid and Compid=f.Compid)Approvedby,a.Branchid,a.Compid from Sal_Customer_ItemMaster a
left outer join FAS_Party_Master b
on a.CustomerCode=b.PartyCode and a.Branchid=b.Branchid and a.Compid=b.Compid
left outer join BOM_Mps_Master c on a.ItemCode=c.ItemCode and a.Branchid=c.Branchid and a.Compid=c.Compid
and c.OperationSequence=(Select Distinct OperationSequence from BOM_Mps_Master where ItemCode=a.ItemCode and StageCode=c.StageCode)
left outer join BOM_Structure_Dtl d on (Select Distinct StageCode from BOM_Mps_Master where 
ItemCode=a.ItemCode and OPerationSequence=1)=d.ParentCode and d.Branchid=c.Branchid and d.Compid=c.Compid
left outer join INV_Material_Master e on e.ItemCode=c.StageCode and e.Branchid=c.Branchid and e.Compid=c.Compid
left outer join (select a.ProcessType,a.ProcessNo,a.ProcessDate,a.RevisionNo,a.RevisionDate,a.EmployeeCode,a.Approvedby,a.ItemCode,a.Branchid,a.Compid
from ProcessFlow a where a.RevisionNo = (select max(b.RevisionNo) from ProcessFlow b
where a.PRocessType=b.ProcessType and a.ProcessNo=b.ProcessNo and a.ProcessDate=b.ProcessDate and 
a.ItemCode= b.ItemCode and b.Branchid=a.Branchid and b.Compid=a.Compid)) f
 on f.ItemCode=e.ItemCode and f.Branchid=a.Branchid and f.Compid=a.Compid
where d.Priority=1 and f.ProcessType is not null
GO

/****** Object:  View [dbo].[ProcessFlow]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[ProcessFlow]
as
select Distinct a.ProcessType,a.ProcessNo,a.ProcessDate,a.RevisionNo,a.RevisionDate,a.EmployeeCode,a.Approvedby,
a.ItemCode,b.StageCode,b.DiagramCode,b.Isprocess,b.ProcessCode,b.ProcessFlow,b.Operator,b.IncomingSource,b.ProductChrstics,b.ProcessChrstics,
b.ProcessSequence,a.UserId,a.Branchid,a.Compid,a.Finyear
from Bom_Process_Flow_Hdr a left outer join Bom_Process_Flow_Dtl b
on a.ProcessType=b.ProcessType and a.ProcessNo=b.ProcessNo and a.ProcessDate=b.ProcessDate and 
a.RevisionNo=b.RevisionNo and a.Branchid=b.Branchid and a.Compid=b.Compid and a.Finyear=b.Finyear
where a.RevisionNo=(Select Max(RevisionNo) from Bom_Process_Flow_Hdr where ProcessType=a.ProcessType and ProcessNo=a.ProcessNo
and ProcessDate=a.ProcessDate and  Branchid=a.Branchid and Compid=a.Compid )
GO

/****** Object:  View [dbo].[ProcessReport]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE view [dbo].[ProcessReport] as

	select x.ProcessMonth, x.EmployeeCode, x.ExprGroupId,  x.ExprSeq, 
		case 	when left(x.ExprSeq, 3) = 100 then 'Insurance' + right(x.ExprSeq,1) 
			when left(x.ExprSeq, 3) = 300 then 'Advance' + right(x.ExprSeq,1) 
			when left(x.ExprSeq, 3) = 400 then 'Professional Tax' + right(x.ExprSeq,1) 

			else isnull(x.ExprGroupName, (select y.ExprGroupName from PAY_FormulaExpr_Group y 
						where y.ExprGroupID = x.ExprSeq and y.Branchid = x.Branchid and y.Compid = x.Compid)) 
		end ExprName,
		sum(x.Resultvalue) Resultvalue, x.GrandTotalFlag, x.ProcessType, x.Branchid, x.Compid 
	from 
		(select a.ProcessMonth, a.EmployeeCode, a.ExprSeq, a.ExprGroupId, b.ExprGroupName, 
			a.ResultValue, a.GrandTotalFlag, a.ProcessType, a.Branchid, a.Compid 
		from pay_formulacalc_hdr a left outer join PAY_FormulaExpr_Group b 
		on a.ExprGroupId = b.ExprGroupId and a.Branchid = b.Branchid and a.Compid = b.Compid
		) x
	group by x.ProcessMonth, x.EmployeeCode, x.ExprGroupId, x.ExprSeq, x.ExprGroupName, x.GrandTotalFlag, x.ProcessType, x.Branchid, x.Compid
GO

/****** Object:  View [dbo].[PUR_LaborOrder_vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO





CREATE View [dbo].[PUR_LaborOrder_vw] 

AS

	Select b.OrderType, b.OrderNo, b.OrderDate, a.VendorCode, a.Source, a.ValidFrom, a.ValidTo, a.CreditDays,	a.OrderingType, a.Status, a.StatusDate, 
		b.ItemGroupKey, b.SentItemCode, b.RecdItemCode, b.ToBeRecdQty,  b.RatePer, b.Rate, b.isActive, b.BranchID, b.CompID, b.Finyear
	from PUR_LaborOrder_Dtl b Left Outer Join PUR_LaborOrder_Hdr a
		On b.OrderType=a.OrderType and b.OrderNo=a.OrderNo and b.OrderDate=a.OrderDate and b.BranchID=a.BranchID and b.CompID=a.CompID



GO

/****** Object:  View [dbo].[PUR_LaborOrderMisc_Vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[PUR_LaborOrderMisc_Vw]
AS
	Select   a.OrderType, a.OrderNo, a.OrderDate, a.OrderingType, a.VendorCode, a.ValidFrom, a.ValidTo, a.CreditDays, a.Status, a.StatusDate, b.ItemCode, 
                      b.Description, b.Qty, b.Rate, b.RatePer, b.isActive, a.BranchID, a.CompID, a.Finyear, a.TranStamp
	from  dbo.PUR_LaborOrderMisc_Hdr a
	 Left Outer Join  dbo.PUR_LaborOrderMisc_Dtl b 
		On a.OrderDate = b.OrderDate  and a.OrderNo = b.OrderNo  and a.OrderType = b.OrderType and
                      		a.BranchID = b.BranchID and a.CompID = b.CompID
GO

/****** Object:  View [dbo].[PUR_LO_Terms_vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO





CREATE View [dbo].[PUR_LO_Terms_vw] 

AS

	Select a.OrderType, a.OrderNo, a.OrderDate, a.VendorCode,
		a.PackingCode,
			isNull((select b1.Description from TERMSMASTER b1 
			where b1.TypeFlg='Packing & Forwarding' and b1.Code=a.PackingCode and b1.Branchid=a.Branchid and b1.Compid=a.Compid), '') PackingDesc,
		a.InsuranceCode,
			isNull((select b4.Description from TERMSMASTER b4 
			where b4.TypeFlg='Insurance' and b4.Code=a.InsuranceCode and b4.Branchid=a.Branchid and b4.Compid=a.Compid), '') InsuranceDesc,
		a.DeliveryCode,
			isNull((select b2.Description from TERMSMASTER b2 
			where b2.TypeFlg='Delivery Instruction' and b2.Code=a.DeliveryCode and b2.Branchid=a.Branchid and b2.Compid=a.Compid), '') DeliveryDesc,
		a.PaymentCode,
			isNull((select b3.Description from TERMSMASTER b3 
			where b3.TypeFlg='Payment Terms' and b3.Code=a.PaymentCode and b3.Branchid=a.Branchid and b3.Compid=a.Compid), '') PaymentDesc,
		a.Freightcode,
			isNull((select b6.Description from TERMSMASTER b6 
			where b6.TypeFlg='Freight' and b6.Code=a.FreightCode and b6.Branchid=a.Branchid and b6.Compid=a.Compid), '') FreightDesc,
		a.ModeCode,
			isNull((select b5.Description from TERMSMASTER b5 
			where b5.TypeFlg='Mode' and b5.Code=a.ModeCode and b5.Branchid=a.Branchid and b5.Compid=a.Compid), '') ModeDesc,
		a.FormCode,
			isNull((select b.Description from TERMSMASTER b 
			where b.TypeFlg='Form' and b.Code=a.FormCode and b.Branchid=a.Branchid and b.Compid=a.Compid), '') FormDesc, 
		a.FormulaCode,
			isNull((Select b7.FrmLName from Com_FormulaMast_Hdr b7
			where b7.FrmID=a.FormulaCode and b7.BranchID=a.BranchID and b7.CompID=a.CompID), '') FormulaName,
		a.BranchID, a.CompID, a.Finyear
	from PUR_LaborOrder_Hdr a



GO

/****** Object:  View [dbo].[PUR_PO_Terms_vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO





CREATE View [dbo].[PUR_PO_Terms_vw] 

AS

	Select a.OrderType, a.OrderNo, a.OrderDate, a.VendorCode,
		a.PackingCode,
			isNull((select b1.Description from TERMSMASTER b1 
			where b1.TypeFlg='Packing & Forwarding' and b1.Code=a.PackingCode and b1.Branchid=a.Branchid and b1.Compid=a.Compid), '') PackingDesc,
		a.InsuranceCode,
			isNull((select b4.Description from TERMSMASTER b4 
			where b4.TypeFlg='Insurance' and b4.Code=a.InsuranceCode and b4.Branchid=a.Branchid and b4.Compid=a.Compid), '') InsuranceDesc,
		a.DeliveryCode,
			isNull((select b2.Description from TERMSMASTER b2 
			where b2.TypeFlg='Delivery Instruction' and b2.Code=a.DeliveryCode and b2.Branchid=a.Branchid and b2.Compid=a.Compid), '') DeliveryDesc,
		a.PaymentCode,
			isNull((select b3.Description from TERMSMASTER b3 
			where b3.TypeFlg='Payment Terms' and b3.Code=a.PaymentCode and b3.Branchid=a.Branchid and b3.Compid=a.Compid), '') PaymentDesc,
		a.Freightcode,
			isNull((select b6.Description from TERMSMASTER b6 
			where b6.TypeFlg='Freight' and b6.Code=a.FreightCode and b6.Branchid=a.Branchid and b6.Compid=a.Compid), '') FreightDesc,
		a.ModeCode,
			isNull((select b5.Description from TERMSMASTER b5 
			where b5.TypeFlg='Mode' and b5.Code=a.ModeCode and b5.Branchid=a.Branchid and b5.Compid=a.Compid), '') ModeDesc,
		a.FormCode,
			isNull((select b.Description from TERMSMASTER b 
			where b.TypeFlg='Form' and b.Code=a.FormCode and b.Branchid=a.Branchid and b.Compid=a.Compid), '') FormDesc, 
		a.FormulaCode,
			isNull((Select b7.FrmLName from Com_FormulaMast_Hdr b7
			where b7.FrmID=a.FormulaCode and b7.BranchID=a.BranchID and b7.CompID=a.CompID), '') FormulaName,
		a.BranchID, a.CompID, a.Finyear
	from PUR_PurchaseOrder_Hdr a



GO

/****** Object:  View [dbo].[PUR_PurchaseOrder_vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE View [dbo].[PUR_PurchaseOrder_vw] 

AS

Select distinct a.VendorCode,a.OrderType, a.OrderNo, a.OrderDate,b.Itemcode,
b.Rate,b.AssRate,b.CompId ,b.branchID
from Pur_PurchaseOrder_Hdr a
left outer Join PUR_PurchaseOrder_Dtl b
On b.OrderType=a.OrderType and b.OrderNo=a.OrderNo 
and b.OrderDate=a.OrderDate and b.BranchID=a.BranchID and b.CompID=a.CompID
GO

/****** Object:  View [dbo].[PurchaseCalculationDetails]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[PurchaseCalculationDetails] as 
select distinct a.DocType, a.DocNo, a.DocDate, a.ExprSeq, b.ExprGroupId, b.ExprName, a.Resultvalue, 
a.GrandTotalFlag, a.ModuleCode, a.BranchID, a.CompID, a.Finyear 
from COM_FormulaCalc_Hdr a Left outer join (Select c.FrmID, d.ExprGroupId, c.ExprSeq, 
case when c.ExprRoute < 0 then 'Sub Total' else d.ExprGroupName End ExprName, 
c.ModuleCode, c.BranchID, c.CompID from COM_FORMULAMAST_DTL c Left Outer Join COM_FormulaExpr_Group d
On c.ModuleCode=d.ModuleCode and c.ExprRoute=d.Route+d.Width and c.BranchID=d.BranchID and c.CompID=d.CompID) b
On a.ModuleCode=b.ModuleCode and a.FrmId=b.FrmId and a.ExprSeq=b.ExprSeq and a.BranchID=b.BranchID and a.CompID=b.CompID
Where a.ModuleCode = 9
GO

/****** Object:  View [dbo].[PurchaseEnquiry]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE VIEW [dbo].[PurchaseEnquiry] 
AS 
SELECT a.EnquiryType, a.EnquiryNo, a.EnquiryDate, a.VendorType, a.VendorCode, a.VendorName, a.Remarks, 
	b.ItemCode, b.Qty, b.TargetDate, b.Specification, b.QuotationType, b.QuotationNo, b.QuotationDate, a.Revisionno, a.Revisiondate, b.Status,
	a.Branchid, a.Compid, a.Finyear 
FROM PUR_ENQUIRY_HDR a Right Outer Join PUR_ENQUIRY_DTL b 
ON a.EnquiryNo=b.EnquiryNo AND a.EnquiryDate=b.EnquiryDate AND a.EnquiryType=b.EnquiryType
AND a.Branchid=b.Branchid AND a.Compid=b.Compid AND a.Finyear=b.Finyear 
Where a.Revisionno = (Select max(Revisionno) from PUR_ENQUIRY_HDR c where a.EnquiryNo=c.EnquiryNo AND a.EnquiryDate=c.EnquiryDate
 AND a.EnquiryType=c.EnquiryType AND a.Branchid=c.Branchid AND a.Compid=c.Compid AND a.Finyear=c.Finyear)
GO

/****** Object:  View [dbo].[PurchaseInvoice]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE view [dbo].[PurchaseInvoice] as 

Select 'P' as Type, a.DocType, a.DocNo, a.DocDate, a.VendorCode, c.RefType, c.RefNo, c.RefDate,  c.ItemCode as RefItemCode, c.Qty as RefQty,
		a.BillRef, a.BillDate, a.BillFormula, a.InvoiceAmount, b.Seq, b.ItemCode, b.Qty, 
		b.RatePer, b.Rate, b.AssRate, b.GoodsValue, a.BranchID, a.CompID, a.Finyear
From PUR_PurchaseInvoice_Hdr a 
Left Outer Join PUR_PurchaseInvoice_Dtl b
	On a.DocType = b.DocType and a.DocNo = b.DocNo and a.DocDate = b.DocDate and a.Branchid = b.Branchid and a.Compid = b.Compid
Left Outer Join Pur_PurchaseInvoice_OrdRef c 
	On a.DocType = c.DocType and a.DocNo = c.DocNo and a.DocDate = c.DocDate 
		and b.ItemCode = c.ItemCode and a.BranchID = c.BranchID and a.CompID = c.CompID 

Union

Select distinct 'L' as Type, a.DocType, a.DocNo, a.DocDate, a.VendorCode, c.RefType, c.RefNo, c.RefDate,  c.ItemCode as RefItemCode, c.Qty as RefQty,
		a.BillRef, a.BillDate, a.BillFormula, a.InvoiceAmount, b.Seq, b.ItemCode, b.Qty, 
		b.RatePer, b.Rate, b.AssRate, b.GoodsValue, a.BranchID, a.CompID, a.Finyear 
From PUR_LaborInvoice_Hdr a 
Left Outer Join PUR_LaborInvoice_Dtl b
	On a.DocType = b.DocType and a.DocNo = b.DocNo and a.DocDate = b.DocDate and a.Branchid = b.Branchid and a.Compid = b.Compid
Left Outer Join Pur_LaborInvoice_OrdRef c 
	On a.DocType = c.DocType and a.DocNo =c.DocNo and a.DocDate = c.DocDate 
		and b.ItemCode = c.ItemCode and a.BranchID = c.BranchID and a.CompID = c.CompID
GO

/****** Object:  View [dbo].[PurchaseQuotation]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[PurchaseQuotation] as 
select a.QuotationType, a.QuotationNo, a.QuotationDate, a.VendorCode, a.VendorName, a.RefNo, a.RefDate, 
a.ValidFrom, a.ValidTo, a.FormulaCode, a.Remarks, a.FormCode, a.PackingCode, a.DeliveryCode, a.PaymentCode, 
a.InsuranceCode, a.ModeCode, a.Freightcode, a.TotalAmount, b.ItemCode, b.Qty, b.RatePer, b.Rate, b.AssRate, 
b.SelectFlag, b.OrderType, b.OrderNo, b.OrderDate, a.BranchID, a.CompID, a.Finyear
from PUR_Quotation_Hdr a left outer join PUR_Quotation_Dtl b on a.QuotationType = b.QuotationType
and a.QuotationNo = b.QuotationNo and a.QuotationDate = b.QuotationDate and a.Branchid = b.Branchid
and a.Compid = b.Compid
GO

/****** Object:  View [dbo].[PurchaseQuotation_Hdr]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO






CREATE View [dbo].[PurchaseQuotation_Hdr] 
As
Select a.QuotationType, a.QuotationNo, a.QuotationDate, a.VendorCode, 
	(Select c.PartyName from FAS_Party_Master c where c.PartyCode=a.VendorCode and c.BranchID=a.BranchID and c.CompID=a.CompID) as VendorName,
	a.ValidFrom, a.ValidTo, 
	a.FormCode,
		isNull((select b.Description from TERMSMASTER b 
		where b.TypeFlg='Form' and b.Code=a.FormCode and b.Branchid=a.Branchid and b.Compid=a.Compid), '') FormDesc, 
	a.PackingCode,
		isNull((select b1.Description from TERMSMASTER b1 
		where b1.TypeFlg='Packing & Forwarding' and b1.Code=a.PackingCode and b1.Branchid=a.Branchid and b1.Compid=a.Compid), '') PackingDesc,
	a.DeliveryCode,
		isNull((select b2.Description from TERMSMASTER b2 
		where b2.TypeFlg='Delivery Instruction' and b2.Code=a.DeliveryCode and b2.Branchid=a.Branchid and b2.Compid=a.Compid), '') DeliveryDesc,
	a.PaymentCode,
		isNull((select b3.Description from TERMSMASTER b3 
		where b3.TypeFlg='Payment Terms' and b3.Code=a.PaymentCode and b3.Branchid=a.Branchid and b3.Compid=a.Compid), '') PaymentDesc,
	a.InsuranceCode,
		isNull((select b4.Description from TERMSMASTER b4 
		where b4.TypeFlg='Insurance' and b4.Code=a.InsuranceCode and b4.Branchid=a.Branchid and b4.Compid=a.Compid), '') InsuranceDesc,
	a.ModeCode,
		isNull((select b5.Description from TERMSMASTER b5 
		where b5.TypeFlg='Mode' and b5.Code=a.ModeCode and b5.Branchid=a.Branchid and b5.Compid=a.Compid), '') ModeDesc,
	a.Freightcode,
		isNull((select b6.Description from TERMSMASTER b6 
		where b6.TypeFlg='Freight' and b6.Code=a.FreightCode and b6.Branchid=a.Branchid and b6.Compid=a.Compid), '') FreightDesc,
	a.FormulaCode,
		isNull((Select b7.FrmLName from Com_FormulaMast_Hdr b7
		where b7.FrmID=a.FormulaCode and b7.BranchID=a.BranchID and b7.CompID=a.CompID), '') FormulaName,
	a.BranchID, a.CompID, a.Finyear, 
	(Select Count(*) from Pur_Quotation_Dtl d where d.OrderNo=0 and d.BranchID=a.BranchID and d.CompID=a.CompID and d.Finyear=a.Finyear) as PendingItem
from Pur_Quotation_Hdr a







GO

/****** Object:  View [dbo].[QC_CARHDR]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE VIEW [dbo].[QC_CARHDR]
AS
SELECT DISTINCT 
                      	DocType,DocNo,DocDate,RefType,RefNo,RefDate,ItemCode,MachineCode,RaisedBy,CARBase,
		isNew,mParamCode,Parameter,FromValueToValue,DefectValue,DeviationStatus,DueDate,ApprovedBy,
		ApprovedDate,CauseEffectTargetDate,WhyWhyTargetDate,SimulationProblem,SimulationReason,
		CARAdequacy,CARAdequacyDate,ReasonForNonAcceptance,ReasonForNonAcceptanceDate,
		HorizondalDeployment,HorizondalDeploymentReason,CARImplementation,CARImplementationDate,
		CARVerified,CARVerifiedDate,CARClosed,CARClosedDate,Revisionno,UserID,BranchID,CompID,
		FinYear,TranStamp
FROM         dbo.Qc_Car_Hdr a
WHERE     (Revisionno =
                          (SELECT     MAX(Revisionno)
                            FROM          Qc_Car_hdr
                            WHERE      a.Doctype = Doctype AND a.Docno = Docno AND a.DocDate = DocDate AND a.Itemcode = Itemcode AND a.Reftype = Reftype AND 
                                                   a.Refno = Refno AND a.Branchid = BranchID AND a.Compid = CompID ))
GO

/****** Object:  View [dbo].[Qc_Inspection_Dtl_Vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE view [dbo].[Qc_Inspection_Dtl_Vw] as
select * from Qc_Inspection_Dtl a where 
TobeEnterTime=(Select Max(TobeEnterTime) from Qc_Inspection_Dtl b where b.InspectionType=a.InspectionType and b.InspectionNo=a.InspectionNo
and b.InspectionDate=a.InspectionDate and b.Branchid=a.Branchid and b.Compid=a.Compid)
GO

/****** Object:  View [dbo].[QC_Kaizen1]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE VIEW [dbo].[QC_Kaizen1]  AS 

	
	SELECT distinct a1.DocType,a1.DocNo,a1.DocDate,a1.BaseType,a1.RefType,a1.RefNo,a1.RefDate,a1.Itemcode,
		a1.RevNo,a1.RevDate,a1.EffStatus,a1.EffDate,a1.TranStamp,a1.UserID,a1.Branchid,a1.CompId,
		a1.Finyear FROM qc_kaizen a1 
		WHERE a1.RevNo = (SELECT MAX(a.RevNo) AS RevNo FROM qc_kaizen a 
		WHERE a1.DocType = a.DocType AND a1.DocNo = a.DocNo 
		AND a1.DocDate = a.DocDate and a1.Itemcode =a.Itemcode 
		AND a1.Branchid = a.Branchid AND a1.Compid = a.Compid and a1.Finyear = a.Finyear)
GO

/****** Object:  View [dbo].[REC_ReceiptEntry_vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE View [dbo].[REC_ReceiptEntry_vw] As
Select a.ReceiptEntryType, a.ReceiptEntryNo, a.ReceiptEntryDate, a.VendorCode, a.InwardType, 
	a.DCNo, a.DCDate, a.InvoiceNo, a.InvoiceDate, a.TransporterName, a.VehicleNo, a.Narration,
	b.ItemCode, b.ScheduleDate, b.ChallanQty, b.ReceiptQty, b.ReceiptDate, b.AcceptedQty, b.RejectionQty, b.ScrapQty,
	b.AcceptedDate, b.Remarks, b.RefType, b.RefNo, b.RefDate, b.ItemGroupKey, a.Branchid, a.Compid, a.Finyear
from REC_ReceiptEntry_Hdr a Left Outer Join REC_ReceiptEntry_Dtl b
On a.ReceiptEntryNo=b.ReceiptEntryNo and a.ReceiptEntryDate=b.ReceiptEntryDate and a.ReceiptEntryType=b.ReceiptEntryType
and a.Branchid=b.Branchid and a.Compid=b.Compid
GO

/****** Object:  View [dbo].[ReceiptEntry]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



 CREATE View [dbo].[ReceiptEntry]   
  
AS   
  
Select a.ReceiptEntryType, a.ReceiptEntryNo, a.ReceiptEntryDate, a.VendorCode, a.InwardType, a.DCNo,  
 a.DCDate, a.InvoiceNo, a.InvoiceDate, a.TransporterName, a.VehicleNo, a.Narration, b.ItemCode,  
 b.ScheduleDate, b.ChallanQty, b.ReceiptQty, b.ReceiptDate, b.AcceptedQty, b.RejectionQty,  
 b.ScrapQty, b.AcceptedDate, b.Remarks, b.RefType, b.RefNo, b.RefDate,b.ItemGroupKey, b.isDrawingRecd, IsNull(b.HeatNo,'') as HeatNo, b.ItemSeq, 
 a.Userid, a.Branchid, a.Compid, a.Finyear  
from REC_ReceiptEntry_Hdr a Right Outer Join REC_ReceiptEntry_Dtl b  
On a.ReceiptEntryNo=b.ReceiptEntryNo and a.ReceiptEntryDate=b.ReceiptEntryDate and a.ReceiptEntryType=b.ReceiptEntryType  
 and a.Branchid=b.Branchid and a.Compid=b.Compid  
  
GO

/****** Object:  View [dbo].[Receive_DismantlingDc]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[Receive_DismantlingDc]  
                as  
                select RefType,RefNo,RefDate,RecdItemCode,ItemGroupKey,ScheduleDate,sum(AccptedQty) AccQty,  
                sum(RejectedQty) RejQty,sum(ScrapQty) ScrapQty,Branchid,Compid, Finyear from PUR_DismantlingDc_Receive_Dtl  
                group by RefType,RefNo,RefDate,RecdItemCode,ItemGroupKey,ScheduleDate, Branchid, Compid, Finyear
GO

/****** Object:  View [dbo].[Receive_ReWorkDc]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[Receive_ReWorkDc]
as
select RefType,RefNo,RefDate,ItemCode,ScheduleDate,sum(AccptedQty) AccQty,
sum(RejectedQty) RejQty,sum(ScrapQty) ScrapQty,Branchid,Compid,Finyear from PUR_ReworkDc_Receive_Dtl
group by RefType,RefNo,RefDate,ItemCode,ScheduleDate,Branchid,Compid,Finyear
GO

/****** Object:  View [dbo].[RejectionDC]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE View [dbo].[RejectionDC] as 
select a.DocType, a.DocNo, a.DocDate, a.RRType, a.RRNo, a.RRDate, a.VendorCode, 
b.ItemCode, b.ScheduleDate, b.RejectionDCQty, b.Remarks, a.Branchid, a.Compid, 
a.Finyear from REC_RejectionDC_Hdr a left outer join REC_RejectionDC_Dtl b
on a.DocType = b.DocType and a.DocNo = b.DocNo and a.DocDate = b.DocDate
and a.Branchid = b.Branchid and a.Compid = b.Compid and a.Finyear = b.Finyear



GO

/****** Object:  View [dbo].[RESX_FileRepository_vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE View [dbo].[RESX_FileRepository_vw]        
as        
Select * From   UnicornERPResx.dbo.[RESX_FileRepository]     
Where  BranchID='00' and CompID='02' 
GO

/****** Object:  View [dbo].[ReworkDc]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[ReworkDc] as 
Select  distinct a.DocType,a.DocNo,a.DocDate,a.VendorCode,a.OrderType,a.OrderNo,a.OrderDate,b.ItemCode,b.ItemGroupKey,sum(b.Qty)Qty,
a.UserID,a.BranchID,a.CompID,a.Finyear
from PUR_ReworkDc_Hdr a
left outer join PUR_ReworkDc_Dtl b
on a.DocType=b.DocType and a.DocNo=b.DocNo and a.DocDate=b.DocDate
and a.Branchid=b.Branchid and a.Compid=b.Compid and a.Finyear=b.Finyear
group by a.DocType,a.DocNo,a.DocDate,a.VendorCode,a.OrderType,a.OrderNo,a.OrderDate,b.ItemCode,b.ItemGroupKey,
a.UserID,a.BranchID,a.CompID,a.Finyear
GO

/****** Object:  View [dbo].[RR_CompletedQty]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create view [dbo].[RR_CompletedQty] as
select RefType as DcsrType,RefNo as DcsrNo,ScheduleDate,ItemCode,Sum(AcceptedQty+RejectionQty+ScrapQty)As CompletedQty,
Branchid,Compid,Finyear from Rec_ReceiptEntry_Dtl
group by RefType,RefNo,ScheduleDate,ItemCode,Branchid,Compid,Finyear
GO

/****** Object:  View [dbo].[sal]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[sal] as
select a.OrderType,a.OrderNo,a.OrderDate,a.OrderingType,case b.ScheduleType when 'Y' then 'Open' else 'Specific' end as ScheduleType 
from Sal_Order_Hdr a 
left outer join SAL_ORDER_DTL b
on a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate and a.Branchid=b.Branchid
and a.Compid=b.Compid and a.Finyear=b.Finyear
GO

/****** Object:  View [dbo].[sal_JobWork]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[sal_JobWork] as
select a.OrderType,a.OrderNo,a.OrderDate,a.OrderingType,case b.ScheduleType when 'Y' then 'Open' else 'Specific' end as ScheduleType 
from Sal_JobWorkOrder_Hdr a 
left outer join SAL_JobWorkORDER_DTL b
on a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate and a.Branchid=b.Branchid
and a.Compid=b.Compid and a.Finyear=b.Finyear
GO

/****** Object:  View [dbo].[Sal_VW_Enquiry_Dtl]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create View [dbo].[Sal_VW_Enquiry_Dtl] as
select * from SAL_Enquiry_Dtl a where Revisionno=
(select max(Revisionno) from SAL_Enquiry_Dtl where EnquiryType=a.EnquiryType and EnquiryNo=a.EnquiryNo
and EnquiryDate=a.EnquiryDate and Branchid=a.Branchid and Compid=a.Compid)
GO

/****** Object:  View [dbo].[Sal_VW_Enquiry_Hdr]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create View [dbo].[Sal_VW_Enquiry_Hdr] as
select * from SAL_Enquiry_Hdr a where Revisionno=
(select max(Revisionno) from SAL_Enquiry_Hdr where EnquiryType=a.EnquiryType and EnquiryNo=a.EnquiryNo
and EnquiryDate=a.EnquiryDate and Branchid=a.Branchid and Compid=a.Compid)
GO

/****** Object:  View [dbo].[Sal_VW_JobWorkEnquiry_Dtl]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create View [dbo].[Sal_VW_JobWorkEnquiry_Dtl] as
select * from Sal_JobWorkEnquiry_Dtl a where Revisionno=
(select max(Revisionno) from Sal_JobWorkEnquiry_Dtl where EnquiryType=a.EnquiryType and EnquiryNo=a.EnquiryNo
and EnquiryDate=a.EnquiryDate and Branchid=a.Branchid and Compid=a.Compid)
GO

/****** Object:  View [dbo].[Sal_VW_JobWorkEnquiry_Hdr]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create View [dbo].[Sal_VW_JobWorkEnquiry_Hdr] as
select * from Sal_JobWorkEnquiry_Hdr a where Revisionno=
(select max(Revisionno) from Sal_JobWorkEnquiry_Hdr where EnquiryType=a.EnquiryType and EnquiryNo=a.EnquiryNo
and EnquiryDate=a.EnquiryDate and Branchid=a.Branchid and Compid=a.Compid)
GO

/****** Object:  View [dbo].[SAL_VW_JobWorkQuotation_Dtl]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create view [dbo].[SAL_VW_JobWorkQuotation_Dtl] as
select * from SAL_JobWorkQuotation_Dtl a where 
Revisionno=(Select max(Revisionno) from SAL_JobWorkQuotation_Dtl where 
QuotationType=a.QuotationType and Quotationno=a.Quotationno 
and QuotationDate=a.QuotationDate)
GO

/****** Object:  View [dbo].[SAL_VW_JobWorkQuotation_Hdr]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create view [dbo].[SAL_VW_JobWorkQuotation_Hdr] as
select * from SAL_JobWorkQuotation_Hdr a where 
Revisionno=(Select max(Revisionno) from SAL_JobWorkQuotation_Hdr where 
QuotationType=a.QuotationType and Quotationno=a.Quotationno 
and QuotationDate=a.QuotationDate)
GO

/****** Object:  View [dbo].[SAL_VW_Quotation_Dtl]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create view [dbo].[SAL_VW_Quotation_Dtl] as
select * from SAL_Quotation_Dtl a where 
Revisionno=(Select max(Revisionno) from SAL_Quotation_Dtl where 
QuotationType=a.QuotationType and Quotationno=a.Quotationno 
and QuotationDate=a.QuotationDate)
GO

/****** Object:  View [dbo].[SAL_VW_Quotation_Hdr]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[SAL_VW_Quotation_Hdr] as
select * from SAL_Quotation_Hdr a where 
Revisionno=(Select max(Revisionno) from SAL_Quotation_Hdr where 
QuotationType=a.QuotationType and Quotationno=a.Quotationno 
and QuotationDate=a.QuotationDate)
GO

/****** Object:  View [dbo].[SalBankInformation]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


/****** Object:  View dbo.SalBankInformation    Script Date: 14/04/2010 1:27:06 PM ******/
CREATE VIEW [dbo].[SalBankInformation]
		AS
		Select * from Sal_BankInformation_Dtl a 
		Where a.TranStamp = (Select max(b.TranStamp) from Sal_BankInformation_Dtl b
			Where b.Itemcode = a.Itemcode and b.WHCode = a.WHCode and b.Branchid = a.Branchid and b.Compid = a.Compid) 

GO

/****** Object:  View [dbo].[Sales_CancelInvoice]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[Sales_CancelInvoice] as 

Select  a.InvoiceType,a.InvoiceNo,a.InvoiceDate,a.InvoiceTime,a.InvoiceSource,a.CustomerCode,
a.AInformation1,a.AInformation2,a.Narration,a.ModeCode,a.VehicleNo,a.DCNo,a.PLASlNo,a.PLADate,
a.NoofPackets,a.InsurancePolicyNo,a.InsurancePolicyDate,a.InsuranceFor,a.InsuranceSlNo,
a.Type,a.FIRNo,a.FIRDate,a.RemovalTime,a.TotalAmount,a.FormulaCode,b.ItemGroupKey, b.ItemCode,b.InvoiceQty,b.Rate,
b.AssessableRate,b.OrderType,b.OrderNo,b.OrderDate,a.RRNo,a.RRDate,a.RRRemarks,b.ReceivedQty,
b.AcceptedQty,a.Branchid,a.Compid,a.Finyear from SAL_Invoice_Cancel_Hdr a inner join SAL_Invoice_Cancel_Dtl b 
on a.InvoiceType=b.InvoiceType and a.InvoiceNo=b.InvoiceNo and a.InvoiceDate=b.InvoiceDate 
and a.Branchid=b.Branchid and a.Compid=b.Compid
GO

/****** Object:  View [dbo].[SalesCalculationDetails]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[SalesCalculationDetails] 

as 

Select distinct a.DocType, a.DocNo, a.DocDate, a.FrmId, b.ExprGroupId, a.ExprSeq, b.ExprName, a.Resultvalue, 
	a.GrandTotalFlag, a.PostingAc, a.ModuleCode, a.BranchID, a.CompID, a.Finyear 
from COM_FormulaCalc_Hdr a Left outer join 
		(Select c.FrmID, d.ExprGroupId, c.ExprSeq, case when c.ExprRoute < 0 then 'Sub Total' else d.ExprGroupName End ExprName, 
			c.ModuleCode, c.BranchID, c.CompID 
		from COM_FORMULAMAST_DTL c Left Outer Join COM_FormulaExpr_Group d
			On c.ModuleCode=d.ModuleCode and c.ExprRoute=d.Route+d.Width and c.BranchID=d.BranchID and c.CompID=d.CompID
		) b
On a.ModuleCode=b.ModuleCode and a.FrmId=b.FrmId and a.ExprSeq=b.ExprSeq and a.BranchID=b.BranchID and a.CompID=b.CompID
where a.DocType in (Select DocumentType from COM_DOCUMENTTYPE_MASTER where PertainingTo='I' and Branchid = a.Branchid and Compid = a.Compid)
GO

/****** Object:  View [dbo].[SalesDC]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[SalesDC] as 

	select a.DCType,a.DCNo,a.DCDate,a.DCSource,a.CustomerCode,a.Remarks,b.ItemCode,b.DCQty,
	b.Rate,b.AssessableRate,b.ItemGroupKey,b.OrderType,b.OrderNo,b.OrderDate,b.InvoiceType,b.InvoiceNo,b.InvoiceDate,a.Branchid,a.Compid,a.Finyear 

	from SAL_DC_HDR a left outer join SAL_DC_DTL b 

	on a.DCType=b.DCType and a.DCNo=b.DCNo and a.DCDate=b.DCDate and a.Branchid=b.Branchid and a.Compid=b.Compid 
		and a.Finyear=b.Finyear
GO

/****** Object:  View [dbo].[SalesEnquiry]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE VIEW [dbo].[SalesEnquiry]
AS
SELECT     a.EnquiryType, a.EnquiryNo, a.EnquiryDate, a.ReferenceNo, a.ReferenceDate, a.CustomerType, a.CustomerCode, a.CustomerName, a.Remarks, 
                      b.ItemType, b.ItemCode, b.ItemDescription, b.UOM, b.Qty, b.Rate, b.AssessableRate, b.TargetDate, b.Specification, b.QuotationType, b.QuotationNo, 
                      b.QuotationDate, a.Branchid, a.Compid, a.Finyear, a.RevisionNo, a.RevisionDate
FROM         dbo.Sal_Enquiry_Hdr a INNER JOIN
                      dbo.Sal_Enquiry_Dtl b ON a.EnquiryType = b.EnquiryType AND a.EnquiryNo = b.EnquiryNo AND a.EnquiryDate = b.EnquiryDate AND a.RevisionNo = b.RevisionNo and
                      a.Branchid = b.Branchid AND a.Compid = b.Compid AND a.Finyear = b.Finyear Where a.RevisionNo = (Select max(RevisionNo) from Sal_Enquiry_Hdr c Where c.EnquiryType = a.EnquiryType 
		and c.EnquiryNo =  a.EnquiryNo and c.EnquiryDate= a.EnquiryDate and c.Branchid = a.Branchid 
		AND c.Compid = a.Compid AND c.Finyear = a.Finyear)
GO

/****** Object:  View [dbo].[SalesInvoice]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE view [dbo].[SalesInvoice] as   
Select  a.InvoiceType,a.InvoiceNo,a.InvoiceDate,a.InvoiceTime,a.InvoiceSource,a.CustomerCode,  
a.AInformation1,a.AInformation2,a.Narration,a.ModeCode,a.VehicleNo,a.DCNo,a.PLASlNo,a.PLADate,  
b.NoofPkts as NoofPackets,a.InsurancePolicyNo,a.InsurancePolicyDate,a.InsuranceFor,a.InsuranceSlNo,  
a.Type,a.FIRNo,a.FIRDate,a.RemovalTime,a.TotalAmount,a.FormulaCode,b.ItemGroupKey, b.ItemCode,b.InvoiceQty,b.Rate,  
b.AssessableRate,b.OrderType,b.OrderNo,b.OrderDate,a.RRNo,a.RRDate,a.RRRemarks,b.ReceivedQty,  
b.AcceptedQty,a.Branchid,a.Compid,a.Finyear from SAL_INVOICE_HDR a inner join SAL_INVOICE_DTL b   
on a.InvoiceType=b.InvoiceType and a.InvoiceNo=b.InvoiceNo and a.InvoiceDate=b.InvoiceDate   
and a.Branchid=b.Branchid and a.Compid=b.Compid
GO

/****** Object:  View [dbo].[SalesOrder]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE view [dbo].[SalesOrder] as 
select a.OrderType,a.OrderNo,a.OrderDate,a.OrderSource,a.CustomerCode,a.ReferenceNo, a.ReferenceDate,a.ValidFrom,a.ValidTo,a.Remarks,a.FormCode,a.PackingCode,a.DeliveryCode,
a.PaymentCode,a.InsuranceCode,a.ModeCode,a.Freightcode,a.FormulaCode,a.TotalAmount, a.RevisionLevel, a.GeneralDetails, a.Status , a.CreatedDate, 
a.OrderAuthorization,a.OrderAuthorizationDate,a.OrderAuthorizationBy,
b.ItemCode, b.ItemDescription, b.PartNo, b.RevisionNo, b.UOM, b.ScheduleType, b.OrderQty,b.Rate,b.AssessableRate,b.Status ItemStatus,a.Branchid,a.Compid,a.Finyear 
from SAL_ORDER_HDR a Inner Join SAL_ORDER_DTL b 
On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate and a.Branchid=b.Branchid and a.Compid=b.Compid  and a.finyear = b.finyear
Where a.CreatedDate = (Select max(CreatedDate)  from SAL_ORDER_HDR Where a.OrderType=OrderType and a.OrderNo=OrderNo and a.OrderDate=OrderDate 
and a.Branchid=Branchid and a.Compid=Compid and a.finyear = finyear)
GO

/****** Object:  View [dbo].[SalesOrderAmendment]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE view [dbo].[SalesOrderAmendment] as 
select a.AmendmentType, a.AmendmentNo, a.AmendmentDate, a.CustomerCode, a.OrderType, 
a.OrderNo, a.OrderDate, a.FormCode, a.PackingCode, a.DeliveryCode, a.PaymentCode, 
a.InsuranceCode, a.ModeCode, a.Freightcode, a.FormulaCode, a.TotalAmount, 
a.AmendmentAuthorization, a.AmendmentAuthorizationDate, a.AmendmentAuthorizationBy,
b.ItemCode, b.ItemDescription, b.PartNo, b.RevisionNo, b.UOM, b.Qty, b.NewQty, b.Rate, b.NewRate, b.AssessableRate, b.NewAssessableRate, 
b.ValidFrom, b.Status, b.UpdateOrder, a.Branchid, a.Compid, a.Finyear 
from SAL_AMENDMENT_HDR a 
left outer join SAL_AMENDMENT_DTL b
on a.AmendmentType = b.AmendmentType and a.AmendmentNo = b.AmendmentNo 
and a.AmendmentDate = b.AmendmentDate and a.Branchid = b.Branchid and a.Compid = b.Compid
GO

/****** Object:  View [dbo].[SalesQuotation]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[SalesQuotation] as   
select a.QuotationType,a.QuotationNo,a.QuotationDate,a.CustomerCode,a.CustomerName,a.Revisionno,
	a.ValidFrom,a.ValidTo,a.FormulaCode,a.Remarks,a.FormCode,a.PackingCode,a.DeliveryCode,  
	a.PaymentCode,a.InsuranceCode,a.ModeCode,a.Freightcode,a.TotalAmount,b.ItemCode,  
	b.ItemDescription,b.UOM,b.Qty,b.Rate,b.AssessableRate,b.OrderType,b.OrderNo,b.OrderDate,a.Branchid,a.Compid,a.Finyear   
from SAL_QUOTATION_HDR a,SAL_QUOTATION_DTL b 

where a.QuotationType=b.QuotationType and a.QuotationNo=b.QuotationNo and a.QuotationDate=b.QuotationDate  
	and a.Branchid=b.Branchid and a.Compid=b.Compid and a.Finyear=b.Finyear  
GO

/****** Object:  View [dbo].[SchEntry]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE VIEW [dbo].[SchEntry]
AS
SELECT     a.ScheduleType, a.ScheduleNo, a.OrderType, a.OrderNo, a.OrderDate, a.ScheduleQty, a.VendorCode, a.ScheduleDate , b.CompID, b.Finyear, 
                      b.ItemCode , b.RefDocType, b.RefDocNo, b.RefDocDate, b.AcceptedQty
FROM         dbo.PUR_Schedule a INNER JOIN
                      dbo.PUR_Schedule_Receive_Dtl b ON a.ScheduleType = b.ScheduleType AND a.ScheduleNo = b.ScheduleNo AND a.Branchid = b.BranchID AND 
                      a.Compid = b.CompID
GO

/****** Object:  View [dbo].[schentrysam]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE VIEW [dbo].[schentrysam]
AS
SELECT     a.VendorCode, a.ScheduleType, a.ScheduleNo, a.ScheduleDate, a.ItemCode, a.ScheduleQty, a.OrderType, a.OrderNo, a.OrderDate, b.CompID, 
                      b.BranchID, b.Finyear, b.RefDocType, b.RefDocNo, b.RefDocDate
FROM         dbo.PUR_Schedule a INNER JOIN
                      dbo.PUR_Schedule_Receive_Dtl b ON a.ScheduleType = b.ScheduleType AND a.ScheduleNo = b.ScheduleNo AND a.Branchid = b.BranchID AND 
                      a.Compid = b.CompID AND a.Finyear = b.Finyear


GO

/****** Object:  View [dbo].[ServerDate]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

Create View [dbo].[ServerDate] as
select getdate()ServerDate
GO

/****** Object:  View [dbo].[SetUpParameter]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[SetUpParameter]
as
select distinct ShortName,ParameterCode,Parameter,a.ItemCode,
case when Type='SP' then 
		case when FromValue=ToValue then case when dimension='Angle' then dbo.cDegree(FromValue) else FromValue end 
else 
		case when dimension='Angle' then dbo.cDegree(FromValue) else FromValue end+' To ' +case when dimension='Angle' then dbo.cDegree(ToValue) else ToValue end 
end else '' end Specification,b.CamCode,b.ToolCode,Type,a.Sequence,a.Branchid,a.Compid from BOM_Inspection_Parameter a 
left outer join BOM_StageWise_ToolArangement b 
on a.ItemCode=b.ItemCode and a.Branchid=b.BRanchid and a.Compid=b.Compid and a.Parameter=b.Activity
where a.Status=1 and Type='SP'
GO

/****** Object:  View [dbo].[sGroup]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE View [dbo].[sGroup] 
as
select a. * from fn_ItemGroupKey('all','00','01')a
inner join
fn_Customer_ItemGroupKey('all','00','01')b
on a.ToCode=b.ToCode
GO

/****** Object:  View [dbo].[StoreRequest]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[StoreRequest] as 
select a.StoreRequestType,a.StoreRequestNo,a.StoreRequestDate,a.RequiredDate,a.SourceDept,a.WOType,a.WONo,
a.WODate,a.ParentCode,a.WOQty,a.PrintFlg,b.ItemCode,b.RequiredQty,b.IssuedQty,b.ReturnQty,b.ShortClose,
a.Branchid,a.Compid,a.Finyear 

from INV_STOREREQUEST_HDR a,INV_STOREREQUEST_DTL b
where a.StoreRequestType=b.StoreRequestType and a.StoreRequestNo=b.StoreRequestNo and a.StoreRequestDate=b.StoreRequestDate
and a.Branchid=b.Branchid and a.Compid=b.Compid and a.Finyear=b.Finyear
GO

/****** Object:  View [dbo].[TableList]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create view [dbo].[TableList] as 
select distinct a.Id,a.Type, a.Name as TableName,b.colid, b.Name as FieldName,c.Name as Datatype, b.Length, b.Prec, b.Scale, b.isnullable as Nullable,
isnull(e.text,'') as DefaultValue, case when f.indid > 0 then 1 else 0 end PrimeryKey,  b.colorder

from sysobjects a

    left outer join syscolumns b on a.Id=b.id

    left outer join systypes c on c.xtype=b.xtype

    left outer join sysobjects d on d.info = b.colid and d.parent_obj=b.Id

    left outer join syscomments e on e.Id=d.id  

    left outer join sysindexkeys f on f.id = b.id and b.colid=f.colid and f.indid=1

where a.type in('U','P','IF','TF','V') and ((a.Type in ('IF', 'TF') and b.Number = 1) OR (a.Type Not In ('IF', 'TF') and b.Number like '%'))
GO

/****** Object:  View [dbo].[tmp_0607_CB_Vw]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

Create View [dbo].[tmp_0607_CB_Vw]
As
Select a.LrCode, a.RefNo, Sum(a.Amount) as BalAmt from fas_billdetails a
where a.compcode='020020062007' and docdate<='2007-03-31'
group by a.LrCode, a.RefNo
having Sum(a.Amount)<>0
GO

/****** Object:  View [dbo].[tmp_BRS2014]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

Create View [dbo].[tmp_BRS2014] AS
Select * from fas_vouchertranpay where transtamp>='2015-09-15' and bankdate is null and docdate<'2015-04-01'
GO

/****** Object:  View [dbo].[Tmp_FAS_VrTranSub]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

Create View [dbo].[Tmp_FAS_VrTranSub] AS
Select Distinct VrRoute, DocNo, DocDate, CompCode from FAS_VoucherTranSub
where ([DrAmt] > 0) and ([CrAmt] > 0)
GO

/****** Object:  View [dbo].[TmpPWOAccept]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE View [dbo].[TmpPWOAccept] as 

Select a.ProcessWorkOrderType, a.ProcessWorkOrderNo, a.ProcessWorkOrderDate, a.ToStageCode, a.WorkorderQty, 
a.CompletedQty as WOAcceptedQty, Sum(b.AcceptQty) IgrAcceptQty, a.BranchID, a.CompID, a.Finyear
From PPC_PROCESSWORKORDER a
Left Outer Join PPC_InwardGoodsReceipt b
On b.WorkOrderType = a.ProcessWorkOrderType and b.WorkOrderNo = a.ProcessWorkOrderNo and b.WorkOrderDate = a.ProcessWorkOrderDate
and b.ItemCode = a.ToStageCode and b.BranchID = a.BranchID and b.CompID = a.CompID
Group by a.ProcessWorkOrderType, a.ProcessWorkOrderNo, a.ProcessWorkOrderDate, a.ToStageCode, a.WorkorderQty, a.CompletedQty,
a.BranchID, a.CompID, a.Finyear Having Sum(b.AcceptQty) < a.CompletedQty
GO

/****** Object:  View [dbo].[TmpPWOReject]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE View [dbo].[TmpPWOReject] as 

Select a.ProcessWorkOrderType, a.ProcessWorkOrderNo, a.ProcessWorkOrderDate, a.ToStageCode, a.WorkorderQty, 
a.RejectionQty as WORejectQty, Sum(b.RejectQty) IgrRejectQty, a.BranchID, a.CompID, a.Finyear
From PPC_PROCESSWORKORDER a
Left Outer Join PPC_InwardGoodsReceipt b
On b.WorkOrderType = a.ProcessWorkOrderType and b.WorkOrderNo = a.ProcessWorkOrderNo and b.WorkOrderDate = a.ProcessWorkOrderDate
and b.ItemCode = a.ToStageCode and b.BranchID = a.BranchID and b.CompID = a.CompID
Group by a.ProcessWorkOrderType, a.ProcessWorkOrderNo, a.ProcessWorkOrderDate, a.ToStageCode, a.WorkorderQty, 
a.RejectionQty,a.BranchID, a.CompID, a.Finyear Having Sum(b.RejectQty) < a.RejectionQty
GO

/****** Object:  View [dbo].[TmpPWOScrap]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE View [dbo].[TmpPWOScrap] as 

Select a.ProcessWorkOrderType, a.ProcessWorkOrderNo, a.ProcessWorkOrderDate, a.ToStageCode, a.WorkorderQty, 
a.ScrapQty as WOScrapQty, Sum(b.ScrapQty) IgrScrapQty, a.BranchID, a.CompID, a.Finyear
From PPC_PROCESSWORKORDER a
Left Outer Join PPC_InwardGoodsReceipt b
On b.WorkOrderType = a.ProcessWorkOrderType and b.WorkOrderNo = a.ProcessWorkOrderNo and b.WorkOrderDate = a.ProcessWorkOrderDate
and b.ItemCode = a.ToStageCode and b.BranchID = a.BranchID and b.CompID = a.CompID
Group by a.ProcessWorkOrderType, a.ProcessWorkOrderNo, a.ProcessWorkOrderDate, a.ToStageCode, a.WorkorderQty, 
a.ScrapQty, a.BranchID, a.CompID, a.Finyear Having Sum(b.ScrapQty) < a.ScrapQty
GO

/****** Object:  View [dbo].[Tmps]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create view [dbo].[Tmps] as 

select Distinct ProcessSequence,ProcessCode,ProcessFlow,a.StageCode,ShortName, 
Case when c.Type='PT' then Parameter else ''end Product, Case when c.Type='PS' or c.Type='SP' then Parameter else '' end Process, 
case when isnumeric(FromValue)=0 then FromValue else case when Dimension='Angle' then dbo.CDegree(fromValue)else fromValue end +'/'+ case when Dimension='Angle' then dbo.CDegree(ToValue)else toValue end end Specification,
'' as tmp , (select InstrumentType from Bom_InstrumentType_Master where InstrumentTypeCode=c.InstrumentTypeCode and Branchid=c.Branchid and Compid=c.Compid)InstrumentType,
''as size,'' as Freq, case c.Defect when 'IdentChars' then '<IC>' when 'Significant' then '<SC>' when 'Major' then '' when 'Critical' then '<CC>' else '' end SplChar ,
'' as Reactionpln ,'' as Sequence from ProcessFlow a left outer join Bom_Inspection_Parameter c on c.ItemCode=case when ProcessCode in('P01','P07') then a.StageCode else '' end 
and a.Branchid=c.Branchid and a.Compid=c.Compid where a.Branchid='00' 
and a.Compid='01'and a.ItemCode='20520100N0' and c.Status=1
GO

/****** Object:  View [dbo].[ToolDrawing_Details]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE view [dbo].[ToolDrawing_Details] as    
    Select Distinct a.Itemcode, b.ItemDescription as ParentDesc, IsNull(c.PartNo, '') as PartNo, IsNull(d.PartyName, '') as CustomerName,
        IsNull(c.CustomerCode, '') as CustomerCode, IsNull(a.DrawingNo, '') as DrawingNo, IsNull(a.RevisionNo, '') as  RevisionNo, 
        IsNull(a.DrawingCode, '') as DrawingCode, a.Scale,'' as OperationName, IsNull(e.ChildCode, '') as ChildCode, IsNull(A.RevisionDate, '') as RevisionDate,
        a.Branchid,a.Compid 
    From Bom_Drawing_Master a  
    Left Outer Join INV_Material_Master b on a.Itemcode = b.Itemcode and a.Branchid=b.Branchid and a.Compid=b.Compid   
    left outer join Sal_Customer_ItemMaster c on c.ItemCode=A.ItemCode and c.Branchid=A.Branchid and c.Compid=A.Compid  
    left outer join FAS_Party_Master d on c.CustomerCode=d.PartyCode and c.Branchid=d.Branchid and c.Compid=d.Compid  
    left outer join BOM_Structure_Dtl e on a.Itemcode=e.ParentCode and a.Branchid=e.Branchid and a.Compid=e.Compid and Priority=1  
    Where a.RevisionNo = (select max(b.RevisionNo) from Bom_Drawing_Master b where a.ItemCode= b.ItemCode and b.Branchid=a.Branchid and b.Compid=a.Compid)  
    --and c.CustomerCode is not null  
/*  
  
select Distinct d.PartyName as CustomerName,c.CustomerCode,e.ChildCode,c.PartNo,e.Priority,  
b.StageCode as ItemCode,a.ToolCode AS StageCode,f.DrawingNo RefNo,f.RevisionNo,f.DrawingCode,f.Scale,  
'' as OperationName,  
g.ItemDescription,f.RevisionDate,a.Branchid,a.Compid from BOM_StagewisetoolMaster_dtl a   
left outer join BOM_Mps_Master b on a.ItemCode=b.StageCode and a.Branchid=b.Branchid and a.Compid=b.Compid  
left outer join Sal_Customer_ItemMaster c on c.ItemCode=b.ItemCode and c.Branchid=b.Branchid and c.Compid=b.Compid  
left outer join (select * from BOM_Drawing_Master a where a.RevisionNo = (select max(b.RevisionNo) from Bom_Drawing_Master b  
where a.ItemCode= b.ItemCode and b.Branchid=a.Branchid and b.Compid=a.Compid)) f   
on f.ItemCode=a.ToolCode and f.Branchid=a.Branchid and f.Compid=a.Compid  
left outer join FAS_Party_Master d  
on c.CustomerCode=d.PartyCode and c.Branchid=d.Branchid and c.Compid=d.Compid  
left outer join BOM_Structure_Dtl e on a.ToolCode=e.ParentCode and a.Branchid=e.Branchid and a.Compid=e.Compid and Priority=1  
left outer join INV_Material_Master g on g.ItemCode=a.ToolCode and g.Branchid=a.Branchid and g.Compid=a.Compid  
--inner join Qc_Inspection_Master h on a.ToolCode=h.ItemCode and h.Branchid=a.Branchid and h.Compid=a.Compid   
where c.CustomerCode is not null  
*/
GO

/****** Object:  View [dbo].[totalplan]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE view [dbo].[totalplan] as 

select max(revisionno)revisionno,itemcode,planqty ,planperiod ,planfor from SAL_SalesPlan group by itemcode,planqty,planperiod,planfor

union


select max(revisionno)revisionno,itemcode,planqty,planperiod,planfor  from PPC_ProductionPlan group by itemcode,planqty,planperiod,planfor


GO

/****** Object:  View [dbo].[TrainingRequest]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO




create view [dbo].[TrainingRequest] as 
select a.RequestNo, a.RequestDate, a.TrainingCode, a.StartDate, a.EndDate, 
b.EmployeeCode, b.Feedback , a.Branchid, a.Compid from 
HRD_TrainingRequest_Hdr a left outer join HRD_TrainingRequest_Dtl b
on a.RequestNo = b.RequestNo and a.RequestDate = b.RequestDate and a.Branchid = b.Branchid
and a.Compid = b.Compid




GO

/****** Object:  View [dbo].[vw_DD_KPI_ShiftStart]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE View [dbo].[vw_DD_KPI_ShiftStart]
As
    Select Convert(Datetime,'07:30',120) as ShiftStart
GO

/****** Object:  View [dbo].[vw_INV_Material_WIP_Tran_Dtl]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[vw_INV_Material_WIP_Tran_Dtl] as
      Select x.ICode, x.DOCType, x.DOCNo, x.DOCDate, x.TransactionDate, x.TransactionTime, x.TransactionItem, x.TransactionQty, 
		    x.Norms, x.FromWarehouseCode, x.FromRackBinCode, x.ToWarehouseCode, x.ToRackBinCode, x.ItemGroupKey, x.FromItem, 
		    x.ToItem, x.Justification, x.LIFO, x.FIFO, x.WA, x.UserID, x.BranchID, x.CompID, x.Finyear, x.mwTranSeq
      From (
        Select a.ICode, a.DOCType, a.DOCNo, a.DOCDate, a.TransactionDate, a.TransactionTime, a.TransactionItem, a.TransactionQty, 
		    a.Norms, a.FromWarehouseCode, a.FromRackBinCode, a.ToWarehouseCode, a.ToRackBinCode, a.ItemGroupKey, a.FromItem, 
		    a.ToItem, a.Justification, a.LIFO, a.FIFO, a.WA, a.UserID, a.BranchID, a.CompID, a.Finyear, 1 As mwTranSeq    
	    From vw_INV_MaterialTran_Dtl a                       				
                                  
	    UNION ALL
          
	    Select a.ICode, a.DOCType, a.DOCNo, a.DOCDate, a.TransactionDate, a.TransactionTime, a.TransactionItem, a.TransactionQty, 
		    a.Norms, a.FromWarehouseCode, a.FromRackBinCode, a.ToWarehouseCode, a.ToRackBinCode, a.ItemGroupKey, a.FromItem, 
		    a.ToItem, a.Justification, a.LIFO, a.FIFO, a.WA, a.UserID, a.BranchID, a.CompID, a.Finyear, 2 As mwTranSeq           
	    From vw_INV_WIPTRAN_DTL a                        			
		    Where Not Exists( Select * From vw_INV_MaterialTran_Dtl t1     
				    Where t1.DocNo=a.DocNo and t1.DocDate=a.DocDate and t1.DocType=a.DocType and t1.TransactionDate=a.TransactionDate           
					    and t1.TransactionTime=a.TransactionTime and t1.TransactionItem=a.TransactionItem and t1.FromwarehouseCode=a.FromWarehouseCode          
					    and t1.FromRackBinCode=a.FromRackBinCode and t1.TowarehouseCode=a.TowarehouseCode and t1.ToRackBinCode=a.ToRackBinCode           
					    and t1.ItemGroupKey=a.ItemGroupKey and t1.FromItem=a.FromItem and t1.ToItem=a.ToItem           
					    and t1.BranchID=a.BranchID and t1.CompID=a.CompID and t1.Finyear=a.Finyear          
			    )
       ) x
GO

/****** Object:  View [dbo].[vw_INV_MaterialTran_Dtl]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

 CREATE VIEW [dbo].[vw_INV_MaterialTran_Dtl] as     
  Select a.TranDocId, a.ICode, a.DOCType, a.DOCNo, a.DOCDate,   
   Convert(datetime,Convert(varchar(10),a.TransactionDate,121) + ' ' + Convert(varchar(8),a.TransactionTime,108),121) As TransactionDate,  
   a.TransactionTime, a.TransactionItem, a.TransactionQty, a.Norms, a.FromWarehouseCode, a.FromRackBinCode,   
   a.ToWarehouseCode, a.ToRackBinCode, a.ItemGroupKey, a.FromItem, a.ToItem, a.Justification,   
   a.LIFO, a.FIFO, a.WA, a.UserID, a.BranchID, a.CompID, a.Finyear    
  From INV_MaterialTran_Dtl a
GO

/****** Object:  View [dbo].[vw_INV_WIPTran_Dtl]    Script Date: 17/06/2026 11:45:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

 CREATE VIEW [dbo].[vw_INV_WIPTran_Dtl] as     
  Select a.TranDocId, a.ICode, a.DOCType, a.DOCNo, a.DOCDate,   
   Convert(datetime,Convert(varchar(10),a.TransactionDate,121) + ' ' + Convert(varchar(8),a.TransactionTime,108),121) As TransactionDate,  
   a.TransactionTime, a.TransactionItem, a.TransactionQty, a.Norms, a.FromWarehouseCode, a.FromRackBinCode,   
   a.ToWarehouseCode, a.ToRackBinCode, a.ItemGroupKey, a.FromItem, a.ToItem, a.Justification,   
   a.LIFO, a.FIFO, a.WA, a.UserID, a.BranchID, a.CompID, a.Finyear  
  From INV_WIPTran_Dtl a
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'FAS_ScheduleMaster_vw'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=N'1' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'FAS_ScheduleMaster_vw'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1[50] 2[25] 3) )"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1 [56] 4 [18] 2))"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "a"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 114
               Right = 191
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "b"
            Begin Extent = 
               Top = 6
               Left = 229
               Bottom = 114
               Right = 400
            End
            DisplayFlags = 280
            TopColumn = 22
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      RowHeights = 220
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'JobworkEnquiry'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=N'1' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'JobworkEnquiry'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1[50] 2[25] 3) )"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1 [56] 4 [18] 2))"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "a"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 114
               Right = 233
            End
            DisplayFlags = 280
            TopColumn = 20
         End
         Begin Table = "b"
            Begin Extent = 
               Top = 6
               Left = 271
               Bottom = 114
               Right = 442
            End
            DisplayFlags = 280
            TopColumn = 6
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      RowHeights = 220
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'JobworkOrder'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=N'1' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'JobworkOrder'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1[50] 2[25] 3) )"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1 [56] 4 [18] 2))"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "a"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 114
               Right = 191
            End
            DisplayFlags = 280
            TopColumn = 17
         End
         Begin Table = "b"
            Begin Extent = 
               Top = 6
               Left = 229
               Bottom = 114
               Right = 400
            End
            DisplayFlags = 280
            TopColumn = 13
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      RowHeights = 220
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'JobworkQuotation'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=N'1' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'JobworkQuotation'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1[50] 2[25] 3) )"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1 [56] 4 [18] 2))"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "a"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 114
               Right = 220
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      RowHeights = 220
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'QC_CARHDR'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=N'1' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'QC_CARHDR'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1[50] 2[25] 3) )"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1 [56] 4 [18] 2))"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "a"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 114
               Right = 191
            End
            DisplayFlags = 280
            TopColumn = 1
         End
         Begin Table = "b"
            Begin Extent = 
               Top = 6
               Left = 229
               Bottom = 114
               Right = 385
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      RowHeights = 220
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'SalesEnquiry'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=N'1' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'SalesEnquiry'
GO


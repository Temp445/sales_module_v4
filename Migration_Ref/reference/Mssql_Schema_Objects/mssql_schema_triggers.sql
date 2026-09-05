/* All Triggers List */

/****** Object:  Trigger Script Date: 17/06/2026 11:50:45 ******/
CREATE TRIGGER [dbo].[TGR_To_Insert_FAS_LedgerMaster_AuthTbl] ON dbo.FAS_LedgerMaster  
For INSERT, UPDATE  
  
AS  
  
BEGIN  
  
 SET NOCOUNT ON;  
  
  If Exists (SELECT * FROM INSERTED)  
    Begin  
     --Delete Existing Data  
     Delete From FAS_LedgerMaster_AuthTbl  
     From FAS_LedgerMaster_AuthTbl a  
     Inner Join INSERTED b  
       On a.LedgerCode=b.LedgerCode and a.CompCode=b.CompCode       
        
     --Insert  
     Insert into FAS_LedgerMaster_AuthTbl  
      Select LedgerCode, LedgerName, Alias, LedgerName, LrRoute, MaxReached, isBillwise, CreditDays, OPbalance,  
          isCostCentreApplicable, 0, Convert(varchar(10), TranStamp, 120),   
          Case When isActive=Convert(bit, '1') Then 'AA' Else 'AH' End, '2004-04-01', '0', CompCode, UserID, TranStamp  
        From INSERTED   
    End  
  
END  
GO

/****** Object:  Trigger Script Date: 17/06/2026 11:50:45 ******/
CREATE TRIGGER [dbo].[TGR_To_Insert_FAS_PartyMaster_AuthTbl_Rep]           
   ON  [dbo].[FAS_Party_Master]            
FOR INSERT, DELETE, UPDATE          
AS           
BEGIN            
 SET NOCOUNT ON;          
  /* Insert Current Inserted/Deleted Data Into Temp Table */    
  
 Declare @CurrDate As datetime;  
 Declare @RevNo As int;  
 Declare @Tmp_RevNo As int;  
 Declare @RevDate As datetime;  
 Declare @Tmp_RevDate As datetime;  
  
 SET @CurrDate=getdate();  
 Select @RevNo=0, @Tmp_RevNo=0, @RevDate='1900-01-01', @Tmp_RevDate='1900-01-01';  
  
 Select * INTO #TmpTbl_PartyMaster     
 From      
 (      
   Select * From INSERTED        
   Union       
   Select * From DELETED        
 ) a      
  
  /* Delete Data In Table */    
 /*  
    Delete From FAS_PartyMaster_AuthTbl_Rep          
    From FAS_PartyMaster_AuthTbl a          
    Inner Join #TmpTbl_PartyMaster b            
   On a.PartyCode=b.PartyCode and a.BranchID=b.BranchID and a.CompID=b.CompID  
 */       
  
 Begin  
  Select @RevNo=Max(a.RevNo), @RevDate=Max(a.RevDate)   
  From FAS_PartyMaster_AuthTbl_Rep a  
  Where Exists( Select * From #TmpTbl_PartyMaster b                       
         Where a.PartyCode=b.PartyCode and a.BranchID=b.BranchID and a.CompID=b.CompID   
          and a.TransTamp=b.TranStamp              
        )      
  
  Select @Tmp_RevNo=IsNull(@RevNo,0)+1, @Tmp_RevDate=@CurrDate;  
 End  
  
  /* Insert Current Inserted/Deleted Data Into Table */    
    ;With PartyMaster_CTE      
    As      
    (Select a.PartyCode, a.VendorCode, a.SuplierCode, a.PartyName, a.AcType, a.FormCode, a.PackingCode, a.DeliveryCode,  
    a.PaymentCode, a.InsuranceCode, a.ModeCode, a.Freightcode, a.IsSubcontractor, a.AliasName, a.EffStatus,      
    a.EffDate, a.CreditDays, a.CreditDaysPeriod, a.CurrencyCode, a.CustomerFormulaCode, a.VendorFormulaCode,  
    a.TinNo, a.PANNo, a.UserID, a.BranchID, a.CompID, a.TranStamp  
   From FAS_Party_Master a         
    Where Exists( Select * From #TmpTbl_PartyMaster b                       
      Where a.PartyCode=b.PartyCode and a.BranchID=b.BranchID and a.CompID=b.CompID             
      )      
    ) Insert Into FAS_PartyMaster_AuthTbl_Rep          
        Select a.PartyCode, a.VendorCode, a.SuplierCode, a.PartyName, a.AcType, a.FormCode, a.PackingCode, a.DeliveryCode,  
     a.PaymentCode, a.InsuranceCode, a.ModeCode, a.Freightcode, a.IsSubcontractor, a.AliasName,    
     a.CreditDays, a.CreditDaysPeriod, a.CurrencyCode, a.CustomerFormulaCode, a.VendorFormulaCode,   
     a.TinNo, a.PANNo, @Tmp_RevNo As RevNo, @Tmp_RevDate As RevDate,   
     (Case When a.EffStatus='0' Then 'H' Else 'A' End) As EffStatus, a.EffDate, a.UserID, a.BranchID, a.CompID,   
     (Case When a.TranStamp Is Null Then @CurrDate Else a.TranStamp End) As TranStamp  
        From PartyMaster_CTE a          
        Inner Join( Select a.PartyCode, a.PartyName, Max(a.TranStamp) As MaxTranStamp,   
          a.BranchID, a.CompID    
         From PartyMaster_CTE a           
         Group by a.PartyCode, a.PartyName, a.BranchID, a.CompID      
        ) b        
    On a.PartyCode=b.PartyCode and a.BranchID=b.BranchID and a.CompID=b.CompID       
     and a.TranStamp=b.MaxTranStamp    
    
 /* Drop Temp Table */    
  DROP TABLE #TmpTbl_PartyMaster;            
END     
GO

/****** Object:  Trigger Script Date: 17/06/2026 11:50:45 ******/
CREATE TRIGGER [dbo].[TGR_To_Insert_QMS_V1_StandardClauseMaster]           
   ON  dbo.QMS_V1_StandardClauseMaster_Rep            
FOR INSERT, DELETE, UPDATE          
AS           
BEGIN            
    SET NOCOUNT ON;          
      
  /* Insert Current Inserted/Deleted Data Into Temp Table */    
    Select * INTO #TmpTbl_StdClauseMast     
    From      
    (      
        Select * From INSERTED        
        Union       
        Select * From DELETED        
    ) a      
  
  /* Delete Data In Table */    
    Delete From QMS_V1_StandardClauseMaster          
    From QMS_V1_StandardClauseMaster a          
    Inner Join #TmpTbl_StdClauseMast b            
   On a.StdClauseID=b.StdClauseID and a.BranchID=b.BranchID and a.CompID=b.CompID       
  
  /* Insert Current Inserted/Deleted Data Into Table */    
    ;With StdClauseListMast_CTE      
    As      
    (Select a.StdClauseID, a.StdClauseName, a.EffStatus, a.UserID, a.BranchID, a.CompID, a.TranStamp  
   From QMS_V1_StandardClauseMaster_Rep a         
    Where Exists( Select * From #TmpTbl_StdClauseMast b                       
      Where a.StdClauseID=b.StdClauseID and a.BranchID=b.BranchID and a.CompID=b.CompID             
      )      
    ) Insert Into QMS_V1_StandardClauseMaster          
        Select a.StdClauseID, a.StdClauseName, a.EffStatus, a.UserID, a.BranchID, a.CompID, b.MaxTranStamp As TranStamp   
        From StdClauseListMast_CTE a          
        Inner Join( Select a.StdClauseID, a.StdClauseName, Max(a.TranStamp) As MaxTranStamp,   
          a.BranchID, a.CompID    
         From StdClauseListMast_CTE a           
         Group by a.StdClauseID, a.StdClauseName, a.BranchID, a.CompID      
        ) b        
    On a.StdClauseID=b.StdClauseID and a.BranchID=b.BranchID and a.CompID=b.CompID       
     and a.TranStamp=b.MaxTranStamp    
    
 /* Drop Temp Table */    
  DROP TABLE #TmpTbl_StdClauseMast;            
END  
GO

/****** Object:  Trigger Script Date: 17/06/2026 11:50:45 ******/
CREATE TRIGGER [dbo].[TGR_To_Insert_INV_MATERIAL_MASTER] ON dbo.INV_MATERIAL_MASTER_HIS  
For INSERT, UPDATE,DELETE  
AS  
BEGIN  
 SET NOCOUNT ON;  
  
 If Exists (SELECT * FROM DELETED)  
 BEGIN  
  Delete From INV_MATERIAL_MASTER  
  From INV_MATERIAL_MASTER a  
  Inner Join DELETED b  
   On a.ItemCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID and a.TranStamp=b.TranStamp  
 END  
  
 If Exists (SELECT * FROM INSERTED)  
 BEGIN  
  
  Delete From INV_MATERIAL_MASTER  
  From INV_MATERIAL_MASTER a  
  Inner Join INSERTED b  
   On a.ItemCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID  
  
  Insert into INV_MATERIAL_MASTER  
  Select a.ItemID, a.ItemCode, a.ItemDescription, a.ItemNature, a.ProductCategoryCode, a.ProductSubCategoryCode, a.PurchaseNature, a.Origin, a.Boughtout, a.ItemSpecification, a.IsHeatNoApp,  
   a.PostingAccount, a.PostingAccountEditable,a.CostCenter, a.UOM, a.PurchaseUOM, a.PurConversionFactor,a.IssueUOM, a.IssConversionFactor, a.AlternateUom, a.AlternateConvFactor,  
   a.MinimumStock, a.MaximumStock,a.ReOrderLevel, a.ReOrderQty, a.SafetyStock, a.MinLeadTime, a.MaxLeadTime, a.IsShelfLifePeriodApp, a.SelfLifePeriod,  
   a.Rate, a.Modvat, a.TariffNo, a.HSNId, a.isAsset, a.TolerancePlusApplicable, a.PurchaseTolerancePlus, a.ToleranceMinusApplicable, a.PurchaseToleranceMinus,  
   a.PackingType, a.StdPackingQty, a.DrawingNo, a.DrawingDescription, a.RevisionNo, a.RevisionDate, a.DrawingPath, a.TrayTypeCode, a.Status, a.Userid, a.BranchId, a.Compid, a.Transtamp  
  From INSERTED a  
  INNER JOIN  
  (  
   Select ItemCode, BranchID, CompID, Max(TranStamp) as MaxTranStamp  
   From INSERTED  
   Group by ItemCode, BranchID, CompID  
  ) b  
  On b.ItemCode= a.ItemCode and b.BranchID=a.BranchID and b.CompID=a.CompID and b.MaxTranStamp=a.TranStamp  
 END  
END  
GO

/****** Object:  Trigger Script Date: 17/06/2026 11:50:45 ******/
CREATE TRIGGER [dbo].[TGR_To_Insert_QMS_V1_StandardDocument_Dtl]           
   ON  dbo.QMS_V1_StandardDocument_Dtl_Rep            
FOR INSERT, DELETE, UPDATE          
AS           
BEGIN            
    SET NOCOUNT ON;          
      
  /* Insert Current Inserted/Deleted Data Into Temp Table */    
    Select * INTO #TmpTbl_StdDocDtl     
    From      
    (      
        Select * From INSERTED        
        Union       
        Select * From DELETED        
    ) a      
  
  /* Delete Data In Table */    
    Delete From QMS_V1_StandardDocument_Dtl          
    From QMS_V1_StandardDocument_Dtl a          
    Inner Join #TmpTbl_StdDocDtl b            
   On a.StdID=b.StdID and a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.SeqNo=b.SeqNo   
   and a.FromDate=b.FromDate /*and a.ToDate=b.ToDate*/ and a.BranchID=b.BranchID   
   and a.CompID=b.CompID and a.Finyear=b.Finyear       
  
  /* Insert Current Inserted/Deleted Data Into Table */    
    ;With StdDocDtl_CTE      
    As      
    (Select a.StdID, a.DocNo, a.DocDate, a.LevelSerialNo, a.MainContent, a.SubContent, a.Parent, a.Relation, a.SeqNo,  
    a.SymRef, a.RefValue, a.DocFileID, a.Remarks, a.FromDate, a.ToDate, a.EffStatus, a.UserID, a.BranchID,   
    a.CompID, a.Finyear, a.TranStamp  
   From QMS_V1_StandardDocument_Dtl_Rep a         
    Where Exists( Select * From #TmpTbl_StdDocDtl b                       
      Where a.StdID=b.StdID and a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.SeqNo=b.SeqNo   
       and a.FromDate=b.FromDate /*and a.ToDate=b.ToDate*/ and a.BranchID=b.BranchID   
       and a.CompID=b.CompID and a.Finyear=b.Finyear               
      )      
    ) Insert Into QMS_V1_StandardDocument_Dtl       
        Select a.StdID, a.DocNo, a.DocDate, a.LevelSerialNo, a.MainContent, a.SubContent, a.Parent, a.Relation, a.SeqNo,  
     a.SymRef,a.RefValue, a.DocFileID, a.Remarks, a.FromDate, a.ToDate, a.EffStatus, a.UserID, a.BranchID,   
     a.CompID, a.Finyear, b.MaxTranStamp As TranStamp   
        From StdDocDtl_CTE a          
        Inner Join( Select a.StdID, a.DocNo, a.DocDate, a.SeqNo, a.FromDate,   
           /*a.ToDate,*/ Max(a.TranStamp) As MaxTranStamp, a.BranchID, a.CompID,   
           a.Finyear    
         From StdDocDtl_CTE a           
         Group by a.StdID, a.DocNo, a.DocDate, a.SeqNo, a.FromDate, /*a.ToDate,*/ a.BranchID,   
          a.CompID, a.Finyear      
        ) b        
    On a.StdID=b.StdID and a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.SeqNo=b.SeqNo   
     and a.FromDate=b.FromDate /*and a.ToDate=b.ToDate*/ and a.BranchID=b.BranchID   
     and a.CompID=b.CompID and a.Finyear=b.Finyear and a.TranStamp=b.MaxTranStamp    
    
 /* Drop Temp Table */    
  DROP TABLE #TmpTbl_StdDocDtl;            
END  
GO

/****** Object:  Trigger Script Date: 17/06/2026 11:50:45 ******/
CREATE TRIGGER [dbo].[TGR_To_Insert_COM_Trace_MaterialTranLotQty] ON dbo.INV_MATERIALTRAN_DTL                                
FOR INSERT, UPDATE, DELETE                                
AS      
BEGIN    -- # 1.1                                                        
    SET NOCOUNT ON;         
     
    Declare @StgCnt As int;    
    SET @StgCnt=0;    
                                                      
    /* PRINT 'START' */                      
    IF Exists ( Select * From DELETED a Where Exists(Select * From BOM_TraceableItems t Where t.ItemCode=a.TransactionItem) )                                
        BEGIN        
            /* Delete Data Into COM_Trace_MaterialTran_TmpQty */                             
            Delete COM_Trace_MaterialTran_TmpQty                                
            From COM_Trace_MaterialTran_TmpQty a                                
            Where Exists( Select * From DELETED s                                
                            Where s.ICode=a.TranID and s.BranchID=a.BranchID and s.CompID=a.CompID       
                                and s.Finyear=a.Finyear                                
                        )                                
        END                               
                            
    /* PRINT 'START 1' */                                
    IF Exists( Select * From Inserted a Where Exists(Select * From BOM_TraceableItems t Where t.ItemCode=a.TransactionItem) )              
        BEGIN /* PRINT 'START 1.1' */                         
            /* PRINT 'START Declaration' */                                             
            Declare @ChildCode Varchar(30), @QtyAvailable Float;                                      
            Declare @StockItem Varchar(30), @StockNorms Float;                                    
            Declare @SetItemID int, @SetItem Varchar(30), @SetTranID bigint, @PrevTranQty Float, @PrevNorms Float;                                
                                     
            Declare @cICode bigint, @cDocType Varchar(3), @cDocNo int, @cDocDate datetime, @cTransactionDate datetime,       
                @cTransactionTime datetime, @cTransactionItem Varchar(30), @cTransactionQty Float, @cNorms Float,                                               
                @cFromWarehouseCode Varchar(6), @cFromRackBinCode Varchar(6), @cToWarehouseCode Varchar(6),       
                @cToRackBinCode Varchar(6), @cFromItem Varchar(30), @cToItem Varchar(30), @cUserID Varchar(10),       
                @cBranchID Varchar(3), @cCompID Varchar(3), @cFinyear Varchar(20), @cTranstamp datetime                                               
                                              
            Declare @LotID bigint, @ReqQty decimal(15, 5);                                                
            Declare @SubLotID bigint;                                
                                
            Declare @AsOnDate datetime, @RptDate datetime, @fAsOnDate datetime, @fRptDate datetime, @PrevFinyear Datetime;                                
            Select @AsOnDate = getdate(), @RptDate=@AsOnDate;              
            --//For Financial year stock import into Stock Verification warehouse.                        
            Select @fAsOnDate=Convert(Varchar(4), year(@cTransactionDate)) + Convert(Varchar(15),'-03-31 23:59:59');        
      
            Select @fRptDate = @fAsOnDate;      
                     
            /* PRINT 'ASSIGN START' */                                    
            Select @cICode=ICode, @cDocType=DocType, @cDocNo=DocNo, @cDocDate=DocDate,       
                @cTransactionDate=(Convert(varchar(10),a.TransactionDate,121)+' '+Convert(varchar(8),a.TransactionTime,108)),                                
                @cTransactionItem=TransactionItem, @cTransactionQty=TransactionQty, @cNorms=Norms, @cFromWarehouseCode=FromWarehouseCode,                                
                @cFromRackBinCode=FromRackBinCode, @cToWarehouseCode=ToWarehouseCode, @cToRackBinCode=ToRackBinCode,           
                @cFromItem=FromItem, @cToItem=ToItem, @cUserID=a.UserID, @cBranchID=a.BranchID, @cCompID=a.CompID,       
               @cFinyear=a.Finyear                                      
            From Inserted a                                     
            Where Exists (Select * From BOM_TraceableItems t Where t.ItemCode=a.TransactionItem)                            
                                
            Select @StockNorms=1, @StockItem=@cTransactionItem;                                      
                                
            /* Inserting Data into Trace Table*/                                
            IF ((@cDocType = 'SV' and @cFromWarehouseCode='SV')                         
                OR  (Exists (Select * From COM_DocumentType_Master s                                 
                                Where s.PertainingTo in ('M') and s.ModuleName='INV' and s.DocumentType=@cDocType                                
                                    and s.BranchID=@cBranchID and s.CompID=@cCompID              
                            )              
                    )              
                )                                             
                BEGIN                                            
                    IF Not Exists (Select * from COM_Trace_LotMaster_TmpQty                                            
                                    Where DocType=@cDocType and DocNo=@cDocNo and DocDate=@cDocDate              
                                        and ItemCode=@cTransactionItem and BranchID=@cBranchID and CompID=@cCompID)                                                                         
                        BEGIN                                
                            /* Lot ID Generation */                                              
                            Select @LotID=IsNull(Max(LotID),0)+1                                              
                            From COM_Trace_LotMaster_TmpQty                             
                            Where BranchID=@cBranchID and CompID=@cCompID                                                 
                                           
                            /* Insert Data Into COM_Trace_LotMaster_TmpQty --> In Future need to remove the before Union Part */         
                            Insert into COM_Trace_LotMaster_TmpQty                  
       Select @cDocType, @cDocNo, @cDocDate, @cTransactionItem, @LotID As LotID, 1 As SubLotID, --'' As HeatNo,                             
        dbo.fn_COM_Get_LatestHeatNo_StockVerification(@cTransactionItem, @cTransactionDate, @cBranchID, @cCompID, @cFinyear) as HeatNo,    
        '' As ManufactureDate, '' As ShelfLifeDays, @cUserID, @cBranchID, @cCompID, @cTransactionDate                              
       UNION                 
       /* the below From part newly added */                             
       Select @cDocType, @cDocNo, @cDocDate, @cTransactionItem, @LotID As LotID, IsNull(a.SubLotID, 1) As SubLotID, --'' As HeatNo,                                             
        a.HeatNo, a.ManufactureDate, a.ShelfLifeDays, @cUserID, @cBranchID, @cCompID, @cTransactionDate                                 
       From Rec_ReceiptEntry_Dtl_Trace a                              
       Where ReceiptEntryType=@cDocType and ReceiptEntryNo=@cDocNo and ReceiptEntryDate=@cDocDate and a.ItemCode=@cTransactionItem                              
        and a.BranchId=@cBranchID and a.CompId=@cCompID                                     
                        END                                             
                    ELSE                                            
                        BEGIN                                            
                            SELECT @LotID=LotID, @SubLotID=IsNull(SubLotID, 1)              
                            From COM_Trace_LotMaster_TmpQty                                            
                            Where DocType=@cDocType and DocNo=@cDocNo and DocDate=@cDocDate                                            
                                and ItemCode=@cTransactionItem and BranchID=@cBranchID and CompID=@cCompID                                            
                        END      
                                      
                    Declare @WhsCode Varchar(10);      
                        Select @WhsCode = Case When @cFromWarehouseCode = 'SV' Then @cToWarehouseCode Else @cFromWarehouseCode End     
                
                    /* Insert Data Into COM_Trace_MaterialTran_TmpQty */                                   
                    Insert into COM_Trace_MaterialTran_TmpQty       
                        --Select @cICode, IsNull(@LotID, b.LotID),  IsNull(@SubLotID, b.SubLotID), @cTransactionItem, @cTransactionQty As TransactionQty, @cTransactionItem,       
                        Select @cICode, IsNull(@LotID, -1),  IsNull(@SubLotID, 1), @cTransactionItem, @cTransactionQty As TransactionQty, @cTransactionItem,       
                            @cTransactionQty As ChildQty, -4 As CriteriaCode, @cFromWarehouseCode As FromWarehouseCode,       
                            @cFromRackBinCode As FromRackBinCode, @cToWarehouseCode As ToWarehouseCode, @cToRackBinCode As ToRackBinCode,                                    
                            '0' As ReverseEntryFlg, -2 UserID, @cBranchID As BranchID, @cCompID As CompID, @cFinyear As Finyear,       
                            @cTransactionDate         
                        --From Inserted a                                  
                        -- Cross Apply fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotSubLotQty(a.TransactionItem, @WhsCode, @cTransactionQty, @fAsOnDate, @fAsOnDate, @fRptDate, @cBranchID, @cCompID, @cFinyear)b                               
      
                    /* PRINT 'ICode: ' + Convert(varchar, @cICode) + ' - Completed - SV Entry.'; */                                         
                END    
                
            ELSE If @cDocType='RR'                                                  
                BEGIN  /* PRINT 'RR START' */                                                    
                    IF Exists (Select * From Rec_ReceiptEntry_Hdr                                 
                                Where ReceiptEntryType=@cDocType and ReceiptEntryNo=@cDocNo and ReceiptEntryDate=@cDocDate                                 
                                    and BranchID=@cBranchID and CompID=@cCompID and InwardType in ('P','C','U'))   --Purchase, CashPurchase, CustomerSuppliedProduct                                
                        BEGIN                                          
                            IF Not Exists (Select * From COM_Trace_LotMaster_TmpQty                                            
                                                Where DocType=@cDocType and DocNo=@cDocNo and DocDate=@cDocDate                                            
                                                    and ItemCode=@cTransactionItem and BranchID=@cBranchID and CompID=@cCompID      
                                            )                                            
                                BEGIN   /* Receipt Entry */                          
                                    /* Lot ID Generation */                                              
                                    Select @LotID=IsNull(Max(LotID),0)+1                                              
                                    From COM_Trace_LotMaster_TmpQty                                                
                                    Where BranchID=@cBranchID and CompID=@cCompID                             
                                              
                                    /* PRINT 'Lot ID Generation'  + Convert(varchar, @LotID) */        
                               
                                    /* Insert Data Into COM_Trace_LotMaster_TmpQty */                                    
                                    Insert into COM_Trace_LotMaster_TmpQty       
                                        Select a.ReceiptEntryType As DocType, a.ReceiptEntryNo As DocNo, a.ReceiptEntryDate As DocDate, @cTransactionItem,                                 
                                           @LotID As LotID, IsNull(c.SubLotID, 1) As SubLotID, IsNull(c.HeatNo,'') as HeatNo,    
             IsNull(c.ManufactureDate,'') as ManufactureDate, IsNull(c.ShelfLifeDays,0) as ShelfLifeDays,     
             @cUserID, @cBranchID, @cCompID, @cTransactionDate                                              
                                        From Rec_ReceiptEntry_Hdr a                                              
                                        Left Outer Join Rec_ReceiptEntry_Dtl b                                              
                                            On a.ReceiptEntryNo=b.ReceiptEntryNo and a.ReceiptEntryDate=b.ReceiptEntryDate and a.ReceiptEntryType=b.ReceiptEntryType                                 
                                                and a.BranchID=b.BranchID and a.CompID=b.CompID                               
                                        Left Outer Join Rec_ReceiptEntry_Dtl_Trace c                                              
                                            On c.ReceiptEntryNo=b.ReceiptEntryNo and c.ReceiptEntryDate=b.ReceiptEntryDate and c.ReceiptEntryType=b.ReceiptEntryType                                 
                                                and c.ItemCode=b.ItemCode and b.ItemSeq=c.ItemSeq and c.BranchID=b.BranchID and c.CompID=b.CompID                                             
                                        Where a.BranchID=@cBranchID and a.CompID=@cCompID and a.Finyear=@cFinyear and b.ItemCode=@cTransactionItem                                              
                                            and a.ReceiptEntryType=@cDocType and a.ReceiptEntryNo=@cDocNo and a.ReceiptEntryDate=@cDocDate                            
        
                                    /* Insert Data Into COM_Trace_LotMaster_TmpQty */                         
                                    Insert into COM_Trace_MaterialTran_TmpQty      
                                        Select @cICode, @LotID, IsNull(c.SubLotID, 1) as SubLotID, @cTransactionItem, IsNull(c.Qty, @cTransactionQty) As TransactionQty,     
           @cTransactionItem,  IsNull(c.Qty, @cTransactionQty) As ChildQty,                           
                                            /* @cTransactionQty As TransactionQty,   @cTransactionItem,   @cTransactionQty As ChildQty, */                          
                                            -4 As CriteriaCode, @cFromWarehouseCode As FromWarehouseCode, @cFromRackBinCode As FromRackBinCode,       
                                            @cToWarehouseCode As ToWarehouseCode, @cToRackBinCode As ToRackBinCode, '0' As ReverseEntryFlg, -2 UserID,       
                                            @cBranchID As BranchID, @cCompID As CompID, @cFinyear As Finyear, @cTransactionDate                                               
                                        From Rec_ReceiptEntry_Dtl b    
          Left Outer Join Rec_ReceiptEntry_Dtl_Trace c         
           On c.ReceiptEntryNo=b.ReceiptEntryNo and c.ReceiptEntryDate=b.ReceiptEntryDate and c.ReceiptEntryType=b.ReceiptEntryType                                 
                                                and c.ItemCode=b.ItemCode and b.ItemSeq=c.ItemSeq and c.BranchID=b.BranchID and c.CompID=b.CompID                                     
                                        Where b.ReceiptEntryNo=@cDocNo and b.ReceiptEntryDate=@cDocDate and b.ReceiptEntryType=@cDocType                               
    and b.ItemCode=@cTransactionItem and b.BranchID=@cBranchID and b.CompID=@cCompID      
                                     
                                    /* PRINT 'COM_Trace_MaterialTran_TmpQty - Receipt Entry' */                         
         END      
                            ELSE                                            
                                BEGIN                                            
                                    Select @LotID=LotID, @SubLotID=SubLotID                                 
                                    From COM_Trace_LotMaster_TmpQty                                            
                                    Where DocType=@cDocType and DocNo=@cDocNo and DocDate=@cDocDate                                            
                                    and ItemCode=@cTransactionItem and BranchID=@cBranchID and CompID=@cCompID                            
                         
                                    /*PRINT 'Lot ID Generation'    + Convert(varchar, @LotID) */                    
                                   
                                    /* Acceptance Entry - If User enter Accepted, Rejection, Scrap qty then material Tran dtl table have 3 rows       
                                    for three different whs Since this 3 entries need to be saved in  table COM_Trace_MaterialTran_TmpQty        
                                    individually As well As SublotID wise. So we saved the entries As below */                      
                             
                                    /* Insert Data Into COM_Trace_MaterialTran_TmpQty - Accepted Qty */                        
                                    Insert into COM_Trace_MaterialTran_TmpQty      
                                        Select @cICode, @LotID, IsNull(c.SubLotID,1) as SubLotID, @cTransactionItem,     
           Case When IsNull(c.AcceptedQty,0)>0 Then c.AcceptedQty Else r.AcceptedQty End As TransactionQty, @cTransactionItem,      
                                            Case When IsNull(c.AcceptedQty,0)>0 Then c.AcceptedQty Else r.AcceptedQty End As ChildQty,          
           --c.AcceptedQty As TransactionQty, @cTransactionItem,  c.AcceptedQty As ChildQty,                           
                                            1 As CriteriaCode, @cFromWarehouseCode As FromWarehouseCode, @cFromRackBinCode As FromRackBinCode, @cToWarehouseCode As ToWarehouseCode, @cToRackBinCode As ToRackBinCode,                                         
 
               
                                            '0' As ReverseEntryFlg, -2 UserID, @cBranchID As BranchID, @cCompID As CompID, @cFinyear As Finyear, @cTransactionDate                                               
                                        From Rec_ReceiptEntry_Dtl r      
                                        Left Outer Join Rec_ReceiptEntry_Dtl_Trace c    
                                            On c.ReceiptEntryNo=r.ReceiptEntryNo and c.ReceiptEntryDate=r.ReceiptEntryDate and c.ReceiptEntryType=r.ReceiptEntryType and c.ItemCode=r.ItemCode    
                                            and c.BranchID=r.BranchID and c.CompID=r.CompID      
                                        Where r.ReceiptEntryNo=@cDocNo and r.ReceiptEntryDate=@cDocDate and r.ReceiptEntryType=@cDocType                               
                                            and r.ItemCode=@cTransactionItem and r.BranchID=@cBranchID and r.CompID=@cCompID and r.AcceptedQty>0    --c.AcceptedQty > 0                          
                                            and @cToWarehouseCode not in (Select WarehouseCode From Inv_Warehouse_Master Where WarehouseNature in ('Rejection', 'Scrap'))                    
                                   
                                    /* PRINT 'COM_Trace_MaterialTran_TmpQty - Accepted Entry - Accepted Qty' */                        
                        
             /* Insert Data Into COM_Trace_MaterialTran_TmpQty - Rejected Qty */                        
                                    Insert into COM_Trace_MaterialTran_TmpQty       
                                        Select @cICode, @LotID, IsNull(c.SubLotID,1) as SubLotID, @cTransactionItem,     
           Case When IsNull(c.RejectedQty,0)>0 Then c.RejectedQty Else r.RejectionQty End As TransactionQty, @cTransactionItem,     
                                            Case When IsNull(c.RejectedQty,0)>0 Then c.RejectedQty Else r.RejectionQty End As ChildQty,    
                                            --c.RejectedQty As TransactionQty, @cTransactionItem,  c.RejectedQty As ChildQty,                                                  
                                            2 As CriteriaCode, @cFromWarehouseCode As FromWarehouseCode, @cFromRackBinCode As FromRackBinCode, @cToWarehouseCode As ToWarehouseCode,              
                                            @cToRackBinCode As ToRackBinCode, '0' As ReverseEntryFlg, -2 UserID, @cBranchID As BranchID, @cCompID As CompID, @cFinyear As Finyear,               
                                            @cTransactionDate                                                  
                                        From Rec_ReceiptEntry_Dtl r    
                                        Left Outer Join Rec_ReceiptEntry_Dtl_Trace c    
                                            On c.ReceiptEntryNo=r.ReceiptEntryNo and c.ReceiptEntryDate=r.ReceiptEntryDate and c.ReceiptEntryType=r.ReceiptEntryType     
            and c.ItemCode=r.ItemCode and c.BranchID=r.BranchID and c.CompID=r.CompID      
                                        Where r.ReceiptEntryNo=@cDocNo and r.ReceiptEntryDate=@cDocDate and r.ReceiptEntryType=@cDocType                               
                                            and r.ItemCode=@cTransactionItem and r.BranchID=@cBranchID and r.CompID=@cCompID and r.RejectionQty > 0      --and c.RejectedQty > 0                          
                                            and @cToWarehouseCode not in (Select WarehouseCode From Inv_Warehouse_Master Where WarehouseNature in ('InHouse', 'Scrap'))                    
                    
                                    /* PRINT 'COM_Trace_MaterialTran_TmpQty - Accepted Entry - Rejected Qty' */                        
                        
                                    /* Insert Data Into COM_Trace_MaterialTran_TmpQty - Scrap Qty */                        
                                    Insert into COM_Trace_MaterialTran_TmpQty      
                                        Select @cICode, @LotID, IsNull(c.SubLotID,1) as SubLotID, @cTransactionItem,     
           Case When IsNull(c.ScrapQty,0)>0 Then c.ScrapQty Else r.ScrapQty End As TransactionQty, @cTransactionItem,     
                                            Case When IsNull(c.ScrapQty,0)>0 Then c.ScrapQty Else r.ScrapQty End As ChildQty,    
                                            --c.ScrapQty As TransactionQty, @cTransactionItem,  c.ScrapQty As ChildQty,                           
                                            4 As CriteriaCode, @cFromWarehouseCode As FromWarehouseCode, @cFromRackBinCode As FromRackBinCode, @cToWarehouseCode As ToWarehouseCode,              
                                            @cToRackBinCode As ToRackBinCode, '0' As ReverseEntryFlg, -2 UserID, @cBranchID As BranchID, @cCompID As CompID, @cFinyear As Finyear,              
                                            @cTransactionDate                                               
                                        From Rec_ReceiptEntry_Dtl r    
                                        Left Outer Join Rec_ReceiptEntry_Dtl_Trace c    
                                            On c.ReceiptEntryNo=r.ReceiptEntryNo and c.ReceiptEntryDate=r.ReceiptEntryDate and c.ReceiptEntryType=r.ReceiptEntryType     
            and c.ItemCode=r.ItemCode and c.BranchID=r.BranchID and c.CompID=r.CompID     
                                       Where r.ReceiptEntryNo=@cDocNo and r.ReceiptEntryDate=@cDocDate and r.ReceiptEntryType=@cDocType                               
                                            and r.ItemCode=@cTransactionItem and r.BranchID=@cBranchID and r.CompID=@cCompID and r.ScrapQty > 0 --and c.ScrapQty > 0                          
                                   and @cToWarehouseCode not in ( Select WarehouseCode From Inv_Warehouse_Master Where WarehouseNature in ('InHouse', 'Rejection'))                    
                    
                                    /* PRINT 'COM_Trace_MaterialTran_TmpQty - Accepted Entry - Scrap Qty' */                         
                                        
                                END      
                        END    
                    /* PRINT 'ICode: ' + Convert(varchar, @cICode) + ' - Completed - RR (Purchase, CashPurchase, Customer Supplied Product) Entry.'; */                                
                    ELSE If @cToWarehouseCode='UI' and Exists (Select * From Rec_ReceiptEntry_Hdr                                 
                                                                Where ReceiptEntryType=@cDocType and ReceiptEntryNo=@cDocNo and ReceiptEntryDate=@cDocDate                                
                                                                    and BranchID=@cBranchID and CompID=@cCompID and InwardType in ('S','L','M')              
                                                            )                       
                        BEGIN     /* Subcontract, SubcontractPartial, DismantlingDC */         --# 1.1.1    
      -- <Temporary Query 30/04/2025 by Mr.Selvaganesh>: (Temporarily using: If LotID, SubLotID, & CBQty  is null, v.LotID=0 / v.SubLotID=1 / v.CBQty=@cTransactionQty will be assigned)  
                            /* TranID Generation */                                 
                            Declare @MaxTranID bigint;                                             
                                Select @MaxTranID=IsNull(Max(TranID),0)+1                                              
                                From COM_Trace_MaterialTran_TmpQty                                                
                                Where ReverseEntryFlg='1' and BranchID=@cBranchID and CompID=@cCompID                                
                                  
                            /* Insert Data Into COM_Trace_MaterialTran_TmpQty - SentItem -> Subcontract to Dummy */                            
                            Insert into COM_Trace_MaterialTran_TmpQty      
                                Select @MaxTranID As ICode, IsNull(b.LotID,0) As LotID, IsNull(b.SubLotID,1) As SubLotID, a.TransactionItem,   
         IsNull(b.CBQty,@cTransactionQty) As TransactionQty,       
                                    a.TransactionItem As ChildCode, IsNull(b.CBQty,@cTransactionQty) As ChildQty, -4 As CriteriaCode, a.ToWarehouseCode As FromWarehouseCode,       
                                    a.ToRackBinCode As FromRackBinCode, 'DUM' As ToWarehouseCode, '' As ToRackBinCode, '1' As ReverseEntryFlg,       
                                    -2 UserID, @cBranchID As BranchID, @cCompID As CompID, @cFinyear As Finyear, @cTransactionDate                                       
                                From INV_MATERIALTRAN_DTL a                                      
                                Cross Apply fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotSubLotQty(a.TransactionItem, a.ToWarehouseCode, @cTransactionQty, @AsOnDate, @AsOnDate, @RptDate, @cBranchID, @cCompID, @cFinyear)b                               
                                Where a.BranchID=@cBranchID and a.CompID=@cCompID                                
                                    and Exists ( Select * From Rec_ReceiptEntry_Dtl s           
                                                Where s.RefNo=a.DocNo and s.RefDate=a.DocDate and s.RefType=a.DocType and s.BranchID=a.BranchID and s.CompID=a.CompID                                
                                                    and s.ReceiptEntryType=@cDocType and s.ReceiptEntryNo=@cDocNo and s.ReceiptEntryDate=@cDocDate                                 
                                                    and s.ItemCode= Case When a.ToItem ='' Then a.TransactionItem Else a.ToItem End                                
                                            )                                
                                                    
                            /* Insert Data Into COM_Trace_MaterialTran_TmpQty - ReceivedItem -> Dummy to Subcontract */                               
                            Insert into COM_Trace_MaterialTran_TmpQty        
                                Select (@MaxTranID+1) As ICode, IsNull(b.LotID,0) As LotID, IsNull(b.SubLotID,1) As SubLotID,   
         @cTransactionItem As TransactionItem, IsNull(b.CBQty,@cTransactionQty) As TransactionQty,   
         @cTransactionItem As ChildCode, IsNull(b.CBQty,@cTransactionQty) As ChildQty, -4 As CriteriaCode,                                      
                                    'DUM' As FromWarehouseCode, '' As FromRackBinCode, @cFromWarehouseCode As ToWarehouseCode,       
                                    @cFromRackBinCode As ToRackBinCode, '1' As ReverseEntryFlg, -2 UserID, @cBranchID As BranchID,       
                                    @cCompID As CompID, @cFinyear As Finyear, @cTransactionDate                                       
                                From INV_MATERIALTRAN_DTL a       
        Outer Apply    
         (    
          Select o.TranID, o.LotID, o.SubLotID, o.ItemCode, o.Qty as CBQty     
          From COM_Trace_MaterialTran_TmpQty o    
          Where o.TranID=@MaxTranID and o.BranchID=@cBranchID and o.CompID=@cCompID --and o.Transtamp=@cTransactionDate     
         ) b    
                                --Cross Apply dbo.fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotSubLotQty(a.TransactionItem, 'DUM', @cTransactionQty, @AsOnDate, @AsOnDate, @RptDate, @cBranchID, @cCompID, @cFinyear)b                                
                                Where b.LotID is Not Null and a.BranchID=@cBranchID and a.CompID=@cCompID                                
                                    and Exists ( Select * From Rec_ReceiptEntry_Dtl s                                
                                                    Where s.RefNo=a.DocNo and s.RefDate=a.DocDate and s.RefType=a.DocType and s.BranchID=a.BranchID and s.CompID=a.CompID                                
                                                        and s.ReceiptEntryType=@cDocType and s.ReceiptEntryNo=@cDocNo and s.ReceiptEntryDate=@cDocDate                                 
                                                        and s.ItemCode= Case When a.ToItem ='' Then a.TransactionItem Else a.ToItem End                                
                                                    )                                
              
                            /* Insert Data Into COM_Trace_MaterialTran_TmpQty */                         
                            Insert into COM_Trace_MaterialTran_TmpQty      
                                Select @cICode, IsNull(b.LotID,0) As LotID, IsNull(b.SubLotID,1) As SubLotID, @cTransactionItem,   
         IsNull(b.CBQty,@cTransactionQty) As TransactionQty, @cTransactionItem As ChildCode,   
         IsNull(b.CBQty,@cTransactionQty) As ChildQty, -4 As CriteriaCode, @cFromWarehouseCode As FromWarehouseCode, +      
                                    @cFromRackBinCode As FromRackBinCode, @cToWarehouseCode As ToWarehouseCode, @cToRackBinCode As ToRackBinCode,                                               
                                    '0' As ReverseEntryFlg, -2 UserID, @cBranchID As BranchID, @cCompID As CompID, @cFinyear As Finyear,   
         @cTransactionDate             
                                From INV_MaterialTran_Dtl a      
        Outer Apply    
         (    
          Select o.TranID, o.LotID, o.SubLotID, o.ItemCode, o.Qty as CBQty     
          From COM_Trace_MaterialTran_TmpQty o    
          Where o.TranID=(@MaxTranID+1) and o.BranchID=@cBranchID and o.CompID=@cCompID --and o.Transtamp=@cTransactionDate     
         ) b    
                                --Cross Apply dbo.fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotSubLotQty(@cTransactionItem, a.ToWarehouseCode, @cTransactionQty, @AsOnDate, @AsOnDate, @RptDate, @cBranchID, @cCompID, @cFinyear)b     
                                Where Exists ( Select * From Rec_ReceiptEntry_Dtl s                                
                                                Where s.RefNo=a.DocNo and s.RefDate=a.DocDate and s.RefType=a.DocType and s.BranchID=a.BranchID and s.CompID=a.CompID                                
                                                    and s.ReceiptEntryType=@cDocType and s.ReceiptEntryNo=@cDocNo and s.ReceiptEntryDate=@cDocDate                                 
                                                    and s.ItemCode= Case When a.ToItem ='' Then a.TransactionItem Else a.ToItem End                                
                                            )    
  
       /* -- <Original Query Backup>:                                  
                            Declare @MaxTranID bigint;                                             
                                Select @MaxTranID=IsNull(Max(TranID),0)+1                                              
                                From COM_Trace_MaterialTran_TmpQty                                                
                                Where ReverseEntryFlg='1' and BranchID=@cBranchID and CompID=@cCompID                                
                                  
                            /* Insert Data Into COM_Trace_MaterialTran_TmpQty - SentItem -> Subcontract to Dummy */                            
                            Insert into COM_Trace_MaterialTran_TmpQty      
                                Select @MaxTranID As ICode, b.LotID, b.SubLotID As SubLotID, a.TransactionItem, b.CBQty As TransactionQty,       
                                    a.TransactionItem As ChildCode, b.CBQty As ChildQty, -4 As CriteriaCode, a.ToWarehouseCode As FromWarehouseCode,       
                                    a.ToRackBinCode As FromRackBinCode, 'DUM' As ToWarehouseCode, '' As ToRackBinCode, '1' As ReverseEntryFlg,       
                                    -2 UserID, @cBranchID As BranchID, @cCompID As CompID, @cFinyear As Finyear, @cTransactionDate                                       
                                From INV_MATERIALTRAN_DTL a                                      
                                Cross Apply fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotSubLotQty(a.TransactionItem, a.ToWarehouseCode, @cTransactionQty, @AsOnDate, @AsOnDate, @RptDate, @cBranchID, @cCompID, @cFinyear)b                               
                                Where a.BranchID=@cBranchID and a.CompID=@cCompID                                
                                    and Exists ( Select * From Rec_ReceiptEntry_Dtl s           
                                                Where s.RefNo=a.DocNo and s.RefDate=a.DocDate and s.RefType=a.DocType and s.BranchID=a.BranchID and s.CompID=a.CompID                                
                                                    and s.ReceiptEntryType=@cDocType and s.ReceiptEntryNo=@cDocNo and s.ReceiptEntryDate=@cDocDate                                 
                                                    and s.ItemCode= Case When a.ToItem ='' Then a.TransactionItem Else a.ToItem End                                
                                            )                                
                                                    
               /* Insert Data Into COM_Trace_MaterialTran_TmpQty - ReceivedItem -> Dummy to Subcontract */                               
                            Insert into COM_Trace_MaterialTran_TmpQty        
                                Select (@MaxTranID+1) As ICode, b.LotID, b.SubLotID As SubLotID, @cTransactionItem As TransactionItem,       
                                    b.CBQty As TransactionQty, @cTransactionItem As ChildCode, b.CBQty As ChildQty, -4 As CriteriaCode,                                      
                                    'DUM' As FromWarehouseCode, '' As FromRackBinCode, @cFromWarehouseCode As ToWarehouseCode,       
                                    @cFromRackBinCode As ToRackBinCode, '1' As ReverseEntryFlg, -2 UserID, @cBranchID As BranchID,       
                                    @cCompID As CompID, @cFinyear As Finyear, @cTransactionDate                                       
                                From INV_MATERIALTRAN_DTL a       
        Outer Apply    
         (    
          Select o.TranID, o.LotID, o.SubLotID, o.ItemCode, o.Qty as CBQty     
          From COM_Trace_MaterialTran_TmpQty o    
          Where o.TranID=@MaxTranID and o.BranchID=@cBranchID and o.CompID=@cCompID --and o.Transtamp=@cTransactionDate     
         ) b    
                                --Cross Apply dbo.fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotSubLotQty(a.TransactionItem, 'DUM', @cTransactionQty, @AsOnDate, @AsOnDate, @RptDate, @cBranchID, @cCompID, @cFinyear)b                                
                                Where b.LotID is Not Null and a.BranchID=@cBranchID and a.CompID=@cCompID                                
                                    and Exists ( Select * From Rec_ReceiptEntry_Dtl s                                
                                                    Where s.RefNo=a.DocNo and s.RefDate=a.DocDate and s.RefType=a.DocType and s.BranchID=a.BranchID and s.CompID=a.CompID                                
                                                        and s.ReceiptEntryType=@cDocType and s.ReceiptEntryNo=@cDocNo and s.ReceiptEntryDate=@cDocDate                                 
                                                        and s.ItemCode= Case When a.ToItem ='' Then a.TransactionItem Else a.ToItem End                                
                                                    )                                
              
                            /* Insert Data Into COM_Trace_MaterialTran_TmpQty */                         
                            Insert into COM_Trace_MaterialTran_TmpQty      
                                Select @cICode, b.LotID, b.SubLotID As SubLotID, @cTransactionItem, b.CBQty As TransactionQty,       
                                    @cTransactionItem As ChildCode, b.CBQty As ChildQty, -4 As CriteriaCode, @cFromWarehouseCode As FromWarehouseCode, +      
                                    @cFromRackBinCode As FromRackBinCode, @cToWarehouseCode As ToWarehouseCode, @cToRackBinCode As ToRackBinCode,                                               
                                    '0' As ReverseEntryFlg, -2 UserID, @cBranchID As BranchID, @cCompID As CompID, @cFinyear As Finyear, @cTransactionDate                                       
                                From INV_MaterialTran_Dtl a      
        Outer Apply    
         (    
          Select o.TranID, o.LotID, o.SubLotID, o.ItemCode, o.Qty as CBQty     
          From COM_Trace_MaterialTran_TmpQty o    
          Where o.TranID=(@MaxTranID+1) and o.BranchID=@cBranchID and o.CompID=@cCompID --and o.Transtamp=@cTransactionDate     
         ) b    
                                --Cross Apply dbo.fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotSubLotQty(@cTransactionItem, a.ToWarehouseCode, @cTransactionQty, @AsOnDate, @AsOnDate, @RptDate, @cBranchID, @cCompID, @cFinyear)b     
                                Where Exists ( Select * From Rec_ReceiptEntry_Dtl s                    
                                                Where s.RefNo=a.DocNo and s.RefDate=a.DocDate and s.RefType=a.DocType and s.BranchID=a.BranchID and s.CompID=a.CompID                                
                                                    and s.ReceiptEntryType=@cDocType and s.ReceiptEntryNo=@cDocNo and s.ReceiptEntryDate=@cDocDate                                 
                                                    and s.ItemCode= Case When a.ToItem ='' Then a.TransactionItem Else a.ToItem End                                
                                            )    
  
      */  
                        END    
                    ELSE      
                        BEGIN        --# 1.1.1      
                            /* Insert Data Into COM_Trace_MaterialTran_TmpQty - Doctype = RR  and Other Inward Types */    
       -- <Temporary Query 30/04/2025 by Mr.Selvaganesh>: (Temporarily using: If LotID, SubLotID, & CBQty  is null, v.LotID=0 / v.SubLotID=1 / v.CBQty=@cTransactionQty will be assigned)  
                            Insert into COM_Trace_MaterialTran_TmpQty       
                                Select @cICode, IsNull(b.LotID,0) As LotID, IsNull(b.SubLotID,1) As SubLotID, @cTransactionItem,   
         IsNull(b.CBQty,@cTransactionQty) As TransactionQty, @cTransactionItem As ChildCode,         
                                    IsNull(b.CBQty,@cTransactionQty) As ChildQty, -4 As CriteriaCode,    
         @cFromWarehouseCode As FromWarehouseCode, @cFromRackBinCode As FromRackBinCode,         
                                    @cToWarehouseCode As ToWarehouseCode, @cToRackBinCode As ToRackBinCode, '0' As ReverseEntryFlg, -2 UserID,         
                                    @cBranchID As BranchID, @cCompID As CompID, @cFinyear As Finyear, @cTransactionDate                                       
                                From fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotSubLotQty(@cTransactionItem, @cFromWarehouseCode, @cTransactionQty, @AsOnDate, @AsOnDate, @RptDate, @cBranchID, @cCompID, @cFinyear) b  
  
       /* -- <Original Query Backup>:  
       Insert into COM_Trace_MaterialTran_TmpQty       
                                Select @cICode, b.LotID, b.SubLotID As SubLotID, @cTransactionItem, b.CBQty As TransactionQty, @cTransactionItem As ChildCode,         
                                    b.CBQty As ChildQty, -4 As CriteriaCode,  @cFromWarehouseCode As FromWarehouseCode, @cFromRackBinCode As FromRackBinCode,         
                                    @cToWarehouseCode As ToWarehouseCode, @cToRackBinCode As ToRackBinCode, '0' As ReverseEntryFlg, -2 UserID,         
                                    @cBranchID As BranchID, @cCompID As CompID, @cFinyear As Finyear, @cTransactionDate                                       
                                From fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotSubLotQty(@cTransactionItem, @cFromWarehouseCode, @cTransactionQty, @AsOnDate, @AsOnDate, @RptDate, @cBranchID, @cCompID, @cFinyear) b  
       */  
                        END          --# 1.1.1      
    
                    /* PRINT 'ICode: ' + Convert(varchar, @cICode) + ' - Completed - RR Other Entry.'; */       
                END    
    
            ELSE                                              
                BEGIN      --# 1.1                                                            
                    If ((Exists (Select * From COM_DocumentType_Master s                                 
                                    Where s.PertainingTo in ('PI', 'SI', 'LO','R') and s.ModuleName='QC' and s.DocumentType = @cFromWarehouseCode                                
                                        and s.BranchID=@cBranchID and s.CompID=@cCompID                                
                                )                                 
                            OR (@cDocType='LOG'))                                 
                            AND @cToWarehouseCode = 'WIP'                        
                        )   
                        BEGIN     --# 1.1.1                               
                            /* Finding Set(Previous) TranId, Item */                                
                            Select @SetTranID=a.ICode, @SetItem=a.TransactionItem, @PrevTranQty=a.TransactionQty, @PrevNorms=a.Norms                                
                            From INV_MATERIALTRAN_DTL a                                
                            Where a.BranchID=@cBranchID and a.CompID=@cCompID and a.Finyear=@cFinyear                                
                                and a.ICode=(Select Max(s.ICode) From INV_MATERIALTRAN_DTL s                                      
                                                Where s.DocType=@CDocType and s.DocNo=@cDocNo and s.DocDate=@cDocDate and s.FromItem=@cFromItem and s.ToItem=@cToItem                                 
                                                    and s.ToWarehouseCode='WAP' and s.ICode<(@cICode)                                 
                                                    and s.BranchID=a.BranchID and s.CompID=a.CompID and s.Finyear=a.Finyear                                   
                                            )                                
                                       
                            /* Insert Data Into COM_Trace_MaterialTran_TmpQty - Log Entry Prev Stage WIP Entry*/                   
                            Insert into COM_Trace_MaterialTran_TmpQty      
                            Select @SetTranID, v.LotID, v.SubLotID As SubLotID, @SetItem, v.CBQty As TransactionQty, /* Check Transaction Qty Concept */                                              
                                isNull(@ChildCode, 0), isNull(@QtyAvailable, 0) As ChildQty, -4 As CriteriaCode,                                               
                                @cToWarehouseCode As FromWarehouseCode, @cToRackBinCode As FromRackBinCode,                                 
                                @cFromWarehouseCode As ToWarehouseCode, @cFromRackBinCode As ToRackBinCode,                                               
                                0 As ReverseEntryFlg, -2 UserID, @cBranchID As BranchID, @cCompID As CompID, @cFinyear As Finyear, @cTransactionDate                                 
                            From fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotSubLotQty(@SetItem, @cToWarehouseCode, @PrevTranQty, @AsOnDate, @AsOnDate, @RptDate, @cBranchID, @cCompID, @cFinyear) v                                       
                                                          
                            /* Insert Data Into COM_Trace_MaterialTran_TmpQty - Log Entry Current Stage WAP Entry*/                                
                            Insert into COM_Trace_MaterialTran_TmpQty      
                            Select @cICode, v.LotID, v.SubLotID As SubLotID, @cTransactionItem, (v.Qty/@PrevNorms) As TransactionQty, /* Check Transaction Qty Concept */                                              
                                v.ChildCode, v.ChildQty As ChildQty, -4 As CriteriaCode,                                               
                                @cFromWarehouseCode As FromWarehouseCode, @cFromRackBinCode As FromRackBinCode, @cToWarehouseCode As ToWarehouseCode,       
                                @cToRackBinCode As ToRackBinCode, 0 As ReverseEntryFlg, -2 UserID, @cBranchID As BranchID, @cCompID As CompID,       
                                @cFinyear As Finyear, @cTransactionDate                                      
                            From COM_Trace_MaterialTran_TmpQty v                                      
                            Where v.TranID=@SetTranID and BranchID=@cBranchID and CompID=@cCompID and Finyear=@cFinyear                                            
                                            
                       /* PRINT 'ICode: ' + Convert(varchar, @cICode) + ' - Completed - LogEntry Entry.'; */      
                        END      --# 1.1.1                                         
                    ELSE If (NOT (@cFromWarehouseCode='WIP' AND @cToWarehouseCode = 'WAP'))                                      
                        BEGIN     --# 1.1.1                                             
                            SET @ReqQty=Convert(decimal(15, 5), @cTransactionQty*@StockNorms);                                              
      
                            /* Insert Data Into COM_Trace_MaterialTran_TmpQty */                                                                       
                            Insert into COM_Trace_MaterialTran_TmpQty    
         
       -- <Temporary Query 30/04/2025 by Mr.Selvaganesh>: (Temporarily using: If LotID, SubLotID, & CBQty  is null, v.LotID=0 / v.SubLotID=1 / v.CBQty=@cTransactionQty will be assigned)  
       Select @cICode, IsNull(v.LotID,0), IsNull(v.SubLotID,1) As SubLotID, @cTransactionItem,   
        IsNull(v.CBQty,@cTransactionQty)*@StockNorms As TransactionQty, /* Check Transaction Qty Concept */                                              
                                isNull(@ChildCode, 0), isNull(@QtyAvailable, 0) As ChildQty, -4 As CriteriaCode,                                               
                                @cFromWarehouseCode As FromWarehouseCode, @cFromRackBinCode As FromRackBinCode,   
        @cToWarehouseCode As ToWarehouseCode, @cToRackBinCode As ToRackBinCode,                               
                                0 As ReverseEntryFlg, -2 UserID, @cBranchID As BranchID, @cCompID As CompID,   
        @cFinyear As Finyear, @cTransactionDate                                              
                            From dbo.fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotSubLotQty(@StockItem, @cFromWarehouseCode, @ReqQty, @AsOnDate, @AsOnDate, @RptDate, @cBranchID, @cCompID, @cFinyear) v                                       
  
                           /* -- <Original Query Backup>:  
       Select @cICode, v.LotID, v.SubLotID As SubLotID, @cTransactionItem, v.CBQty*@StockNorms As TransactionQty, /* Check Transaction Qty Concept */                                              
                                isNull(@ChildCode, 0), isNull(@QtyAvailable, 0) As ChildQty, -4 As CriteriaCode,                                               
                                @cFromWarehouseCode As FromWarehouseCode, @cFromRackBinCode As FromRackBinCode, @cToWarehouseCode As ToWarehouseCode, @cToRackBinCode As ToRackBinCode,                               
                                0 As ReverseEntryFlg, -2 UserID, @cBranchID As BranchID, @cCompID As CompID, @cFinyear As Finyear, @cTransactionDate                                              
                            From dbo.fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotSubLotQty(@StockItem, @cFromWarehouseCode, @ReqQty, @AsOnDate, @AsOnDate, @RptDate, @cBranchID, @cCompID, @cFinyear) v  
       */  
                                        
                            /* PRINT 'ICode: ' + Convert(varchar, @cICode) + ' - Completed - Other Entry.'; */                                            
                        END      --# 1.1.1       
                END                --# 1.1      
        END    
END  
GO

/****** Object:  Trigger Script Date: 17/06/2026 11:50:45 ******/
 CREATE TRIGGER [dbo].[TGR_To_Insert_COM_Trace_WIPTranLotQty] ON dbo.INV_WIPTran_Dtl  
For INSERT, UPDATE, DELETE  
AS  
  
BEGIN  
  
 SET NOCOUNT ON;  
  
 IF Exists   
  (  
   Select * From Inserted a  
   Where Exists   
    (  
     Select * From BOM_TraceableItems t   
     Where t.ItemCode=a.TransactionItem  
    )  
    and not exists  
     (  
      Select * From INV_MaterialTran_Dtl s  
      Where s.DocNo=a.DocNo and s.DocDate=a.DocDate and s.DocType=a.DocType and s.TransactionDate=a.TransactionDate   
       and s.TransactionTime=a.TransactionTime and s.TransactionItem=a.TransactionItem and s.FromwarehouseCode=a.FromWarehouseCode  
       and s.FromRackBinCode=a.FromRackBinCode and s.TowarehouseCode=a.TowarehouseCode and s.ToRackBinCode=a.ToRackBinCode   
       and s.ItemGroupKey=a.ItemGroupKey and s.FromItem=a.FromItem and s.ToItem=a.ToItem   
       and s.BranchID=a.BranchID and s.CompID=a.CompID and s.Finyear=a.Finyear  
     )   
                and Not (DocType='IGR' and FromWarehouseCode='WIP')  
    --and Not (DocType='IGR' and FromWarehouseCode='WIP' and ToWarehouseCode='WAP')  
  )  
  BEGIN  
   /* Declaration */           
   Declare @ChildCode Varchar(30), @QtyAvailable Float        
   Declare @StockItem Varchar(30), @StockNorms Float      
   Declare @SetItem Varchar(30), @SetTranID bigint, @PrevTranQty Float, @PrevNorms Float  
       
   Declare @cICode bigint, @cDocType Varchar(3), @cDocNo int, @cDocDate datetime,  @cTransactionDate datetime , @cTransactionTime datetime,   
    @cTransactionItem Varchar(30), @cTransactionQty Float, @cNorms Float,                 
    @cFromWarehouseCode Varchar(6), @cFromRackBinCode Varchar(6), @cToWarehouseCode Varchar(6), @cToRackBinCode Varchar(6) ,   
    @cFromItem Varchar(30), @cToItem Varchar(30), @cUserID Varchar(10), @cBranchID Varchar(3), @cCompID Varchar(3), @cFinyear Varchar(20), @cTranstamp datetime                 
                
   Declare @LotID bigint, @SubLotID bigint, @ReqQty decimal(15, 5);                  
       
   Declare @AsOnDate datetime, @RptDate datetime  
    Select @AsOnDate = getdate(), @RptDate=@AsOnDate  
       
   /* Assign */       
   Select @cICode=ICode, @cDocType=DocType, @cDocNo=DocNo, @cDocDate=DocDate, @cTransactionDate=(TransactionDate + Convert(varchar(8), TransactionTime, 108)),  
     @cTransactionItem=TransactionItem, @cTransactionQty=TransactionQty, @cNorms=Norms,   
     @cFromWarehouseCode=FromWarehouseCode, @cFromRackBinCode=FromRackBinCode, @cToWarehouseCode=ToWarehouseCode,   
     @cToRackBinCode=ToRackBinCode,  @cFromItem=FromItem, @cToItem=ToItem, @cUserID=a.UserID,   
     @cBranchID=a.BranchID, @cCompID=a.CompID, @cFinyear=a.Finyear        
   From Inserted a                        
   Where Exists   
    (  
     Select * From BOM_TraceableItems t   
     Where t.ItemCode=a.TransactionItem  
    )  
  
   Select @StockNorms = 1, @StockItem=@cTransactionItem        
  
   IF (@cDocType = 'SV' and @cFromWarehouseCode='SV')                                 
                BEGIN                                          
                    IF Not Exists (Select * from COM_Trace_LotMaster_TmpQty                                          
                                    Where DocType=@cDocType and DocNo=@cDocNo and DocDate=@cDocDate            
                                        and ItemCode=@cTransactionItem and BranchID=@cBranchID and CompID=@cCompID)                                                                       
                        BEGIN                              
                            /* Lot ID Generation */                                            
                            Select @LotID=IsNull(Max(LotID),0)+1                                            
                            From COM_Trace_LotMaster_TmpQty                           
                            Where BranchID=@cBranchID and CompID=@cCompID                                               
                                         
                            /* Insert Data Into COM_Trace_LotMaster_TmpQty */       
                            Insert into COM_Trace_LotMaster_TmpQty                
        Select @cDocType, @cDocNo, @cDocDate, @cTransactionItem, @LotID As LotID, 1 As SubLotID, --'' As HeatNo,                           
         dbo.fn_COM_Get_LatestHeatNo_StockVerification(@cTransactionItem, @cTransactionDate, @cBranchID, @cCompID, @cFinyear) as HeatNo,  
         '' As ManufactureDate, '' As ShelfLifeDays, @cUserID, @cBranchID, @cCompID, @cTransactionDate                            
                        END                                           
                    ELSE                                          
                        BEGIN                                          
                            SELECT @LotID=LotID, @SubLotID=IsNull(SubLotID, 1)            
                            From COM_Trace_LotMaster_TmpQty                                          
                            Where DocType=@cDocType and DocNo=@cDocNo and DocDate=@cDocDate                                          
                                and ItemCode=@cTransactionItem and BranchID=@cBranchID and CompID=@cCompID                                          
                        END    
                                    
                    Declare @WhsCode Varchar(10);    
                        Select @WhsCode = Case When @cFromWarehouseCode = 'SV' Then @cToWarehouseCode Else @cFromWarehouseCode End   
              
                    /* Insert Data Into COM_Trace_MaterialTran_TmpQty */                                 
                    Insert into COM_Trace_MaterialTran_TmpQty     
                        --Select @cICode, IsNull(@LotID, b.LotID),  IsNull(@SubLotID, b.SubLotID), @cTransactionItem, @cTransactionQty As TransactionQty, @cTransactionItem,     
                        Select @cICode, IsNull(@LotID, -1),  IsNull(@SubLotID, 1), @cTransactionItem, @cTransactionQty As TransactionQty, @cTransactionItem,     
                            @cTransactionQty As ChildQty, -4 As CriteriaCode, @cFromWarehouseCode As FromWarehouseCode,     
                            @cFromRackBinCode As FromRackBinCode, @cToWarehouseCode As ToWarehouseCode, @cToRackBinCode As ToRackBinCode,                                  
                            '0' As ReverseEntryFlg, -2 UserID, @cBranchID As BranchID, @cCompID As CompID, @cFinyear As Finyear,     
                            @cTransactionDate       
                        --From Inserted a                                
                        -- Cross Apply fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotSubLotQty(a.TransactionItem, @WhsCode, @cTransactionQty, @fAsOnDate, @fAsOnDate, @fRptDate, @cBranchID, @cCompID, @cFinyear)b                             
    
                    /* PRINT 'ICode: ' + Convert(varchar, @cICode) + ' - Completed - SV Entry.'; */                                       
                END  
              
   ELSE IF Not (@cDocType='IGR' and @cFromWarehouseCode='WIP')  --IF Not (@cDocType='IGR' and @cFromWarehouseCode='WIP' and @cToWarehouseCode='WAP')  
    BEGIN  
     IF ( (Exists (Select * From COM_DocumentType_Master s   
         Where s.PertainingTo in ('PI', 'SI', 'LO','R') and s.ModuleName='QC' and s.DocumentType=@cFromWarehouseCode  
          and s.BranchID=@cBranchID and s.CompID=@cCompID  
         )  
          
        OR (@cDocType='LOG'))  
           
        AND @cToWarehouseCode = 'WIP'  
      )  
      BEGIN   
       /* Finding Set(Previous) TranId, Item */  
       Select @SetTranID=-(a.ICode), @SetItem=a.TransactionItem, @PrevTranQty=a.TransactionQty, @PrevNorms=a.Norms  
       From INV_WIPTRAN_DTL a  
       Where a.BranchID=@cBranchID and a.CompID=@cCompID and a.Finyear=@cFinyear  
        and a.ICode =  
         (  
          Select Max(s.ICode)  
          From INV_WIPTRAN_DTL s        
          Where s.DocType=@CDocType and s.DocNo=@cDocNo and s.DocDate=@cDocDate   
           and s.FromItem=@cFromItem and s.ToItem=@cToItem   
           and s.ToWarehouseCode = 'WAP' and s.ICode< (@cICode)   
           and s.BranchID=a.BranchID and s.CompID=a.CompID and s.Finyear=a.Finyear   
         )  
         
       /* Log Entry Prev Stage WIP Entry*/  
       Insert into COM_Trace_MaterialTran_TmpQty                
        Select @SetTranID, v.LotID, IsNull(v.SubLotID, 1) as SubLotID, @SetItem, v.CBQty as TransactionQty, /* Check Transaction Qty Concept */                
         isNull(@ChildCode, 0), isNull(@QtyAvailable, 0) as ChildQty, -4 as CriteriaCode,                 
         @cToWarehouseCode as FromWarehouseCode, @cToRackBinCode as FromRackBinCode,   
         'WAP' as ToWarehouseCode, @cFromRackBinCode as ToRackBinCode,                 
         0 as ReverseEntryFlg, -2 UserID, @cBranchID as BranchID, @cCompID as CompID, @cFinyear as Finyear, @cTransactionDate   
        From fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotSubLotQty(@SetItem, @cToWarehouseCode, @PrevTranQty, @AsOnDate, @AsOnDate, @RptDate, @cBranchID, @cCompID, @cFinyear) v         
        --From fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotQty(@SetItem, @cToWarehouseCode, @PrevTranQty, @AsOnDate, @AsOnDate, @RptDate, @cBranchID, @cCompID, @cFinyear) v         
  
  
       /* Log Entry Current Stage WAP Entry*/  
       Insert into COM_Trace_MaterialTran_TmpQty                
        Select -(@cICode), v.LotID, IsNull(v.SubLotID, 1) as SubLotID, @cTransactionItem, (v.Qty/@PrevNorms) as TransactionQty, /* Check Transaction Qty Concept */                
         v.ChildCode, v.ChildQty as ChildQty, -4 as CriteriaCode,                 
         @cFromWarehouseCode as FromWarehouseCode, @cFromRackBinCode as FromRackBinCode, @cToWarehouseCode as ToWarehouseCode, @cToRackBinCode as ToRackBinCode,                 
         0 as ReverseEntryFlg, -2 UserID, @cBranchID as BranchID, @cCompID as CompID, @cFinyear as Finyear, @cTransactionDate        
        From COM_Trace_MaterialTran_TmpQty v        
        Where v.TranID=@SetTranID and BranchID=@cBranchID and CompID=@cCompID and Finyear=@cFinyear              
              
       --PRINT 'ICode: ' + Convert(varchar, @cICode) + ' - Completed - LogEntry Entry.';              
     
      END    
           
     ELSE IF (NOT (@cFromWarehouseCode='WIP' AND @cToWarehouseCode = 'WAP'))        
      BEGIN               
       SET @ReqQty=Convert(decimal(15, 5), @cTransactionQty*@StockNorms);                
                
       Insert into COM_Trace_MaterialTran_TmpQty                
        Select -(@cICode), v.LotID, IsNull(v.SubLotID, 1) as SubLotID, @cTransactionItem, v.CBQty*@StockNorms as TransactionQty, /* Check Transaction Qty Concept */                
         isNull(@ChildCode, 0), isNull(@QtyAvailable, 0) as ChildQty, -4 as CriteriaCode,                 
         @cFromWarehouseCode as FromWarehouseCode, @cFromRackBinCode as FromRackBinCode, @cToWarehouseCode as ToWarehouseCode, @cToRackBinCode as ToRackBinCode,                 
         0 as ReverseEntryFlg, -2 UserID, @cBranchID as BranchID, @cCompID as CompID, @cFinyear as Finyear, @cTransactionDate                
         From fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotSubLotQty(@SetItem, @cToWarehouseCode, @PrevTranQty, @AsOnDate, @AsOnDate, @RptDate, @cBranchID, @cCompID, @cFinyear) v         
  
        --From fn_COM_Trace_AsOnDate_Stock_ItemWHS_LotQty(@StockItem, @cFromWarehouseCode, @ReqQty, @AsOnDate, @AsOnDate, @RptDate, @cBranchID, @cCompID, @cFinyear) v                
              
       --PRINT 'ICode: ' + Convert(varchar, @cICode) + ' - Completed - Other Entry.';              
      END           
    END  
  END      
END  
GO

/****** Object:  Trigger Script Date: 17/06/2026 11:50:45 ******/ 
CREATE TRIGGER [dbo].[TGR_To_Insert_BOM_Inspection_Parameter] ON [dbo].[BOM_Inspection_Parameter_His]                
For INSERT, UPDATE,DELETE                
AS                
BEGIN                
 SET NOCOUNT ON;                
                
 If Exists (SELECT * FROM DELETED)                
  BEGIN                
   Delete From BOM_Inspection_Parameter                
   From BOM_Inspection_Parameter a                
   Inner Join DELETED b                
    On a.ItemCode=b.ItemCode and a.ParamCode=b.ParamCode and a.BranchID=b.BranchID and a.CompID=b.CompID  and a.TranStamp=b.TranStamp                
  END                
                
 If Exists (SELECT * FROM INSERTED)                
  BEGIN                
        
  Delete From BOM_Inspection_Parameter                
  From BOM_Inspection_Parameter a                
  Inner Join INSERTED b                
   On a.ItemCode=b.ItemCode  and a.ParamCode=b.ParamCode and a.BranchID=b.BranchID and a.CompID=b.CompID                
                
  Insert into BOM_Inspection_Parameter                
   Select a.ParamCode, a.ItemCode, a.ShortName, a.ParameterCode, a.Parameter, a.InstrumentTypeCode, a.Dimension, a.FromValue, a.ToValue, a.Mean, a.Defect, a. Validate,a. Exption, a.Sequence,      
   a.Status, a.Type, a.Userid, a.Branchid, a.Compid, a.Transtamp                
   From INSERTED a                
   INNER JOIN         
    (                
     Select ItemCode, ParamCode, BranchID, CompID,  Max(TranStamp) as MaxTranStamp                 
     From INSERTED                
     Group by ItemCode, ParamCode, BranchID, CompID                
    ) b                
    On  b.ItemCode= a.ItemCode  and a.ParamCode=b.ParamCode and b.BranchID=a.BranchID and b.CompID=a.CompID and b.MaxTranStamp=a.TranStamp                  
  END                
END   
 GO

/****** Object:  Trigger Script Date: 17/06/2026 11:50:45 ******/ 
CREATE TRIGGER [dbo].[TGR_To_Insert_BOM_Drawing_Master] ON [dbo].[Bom_Drawing_Master_His]                
For INSERT, UPDATE,DELETE                
AS                
BEGIN                
 SET NOCOUNT ON;                
                
 If Exists (SELECT * FROM DELETED)                
  BEGIN                
   Delete From BOM_Drawing_Master                
   From BOM_Drawing_Master a                
   Inner Join DELETED b                
    On a.ItemCode=b.ItemCode and a.RevisionNo=b.RevisionNo and a.Revisiondate=b.Revisiondate    
  and a.Type=b.Type and a.BranchID=b.BranchID and a.CompID=b.CompID  and a.TranStamp=b.TranStamp                
  END                
                
 If Exists (SELECT * FROM INSERTED)                
  BEGIN                
        
  Delete From BOM_Drawing_Master                
  From BOM_Drawing_Master a                
  Inner Join INSERTED b                
   On a.ItemCode=b.ItemCode and a.Type=b.Type and a.BranchID=b.BranchID and a.CompID=b.CompID                
                
  Insert into BOM_Drawing_Master                
   Select a.ItemCode, a.DrawingNo, a.DrawingDate, a.DrawingCode, a.Description, a.Modification, a.RevisionNo, a.Revisiondate,   
        a.Type, a.Scale, a.Status, a. TotalCopies,a. DrawingPath,      
        a.Userid, a.Branchid, a.Compid, a.Transtamp                
   From INSERTED a                
   INNER JOIN         
    (                
     Select ItemCode, Type, BranchID, CompID,  Max(TranStamp) as MaxTranStamp                 
     From INSERTED                
     Group by ItemCode, Type, BranchID, CompID                
    ) b                
    On  b.ItemCode= a.ItemCode and a.Type=b.Type and b.BranchID=a.BranchID and b.CompID=a.CompID and b.MaxTranStamp=a.TranStamp                  
  END                
END 
 GO

/****** Object:  Trigger Script Date: 17/06/2026 11:50:45 ******/  
 CREATE TRIGGER [dbo].[TGR_To_Insert_BOM_MPS_MASTER] ON [dbo].[BOM_MPS_MASTER_His]                  
For INSERT, UPDATE,DELETE                  
AS                  
BEGIN                  
 SET NOCOUNT ON;                  
                  
 If Exists (SELECT * FROM DELETED)                  
  BEGIN                  
   Delete From BOM_MPS_MASTER                  
   From BOM_MPS_MASTER a                  
   Inner Join DELETED b                  
    On a.ItemCode=b.ItemCode and a.StageCode=b.StageCode and a.RevisionNo=b.RevisionNo      
  and a.BranchID=b.BranchID and a.CompID=b.CompID  and a.TranStamp=b.TranStamp                  
  END                  
                  
 If Exists (SELECT * FROM INSERTED)                  
  BEGIN                  
          
  Delete From BOM_MPS_MASTER                  
  From BOM_MPS_MASTER a                  
  Inner Join INSERTED b                  
   On a.ItemCode=b.ItemCode and a.StageCode=b.StageCode and a.RevisionNo=b.RevisionNo    
  and a.BranchID=b.BranchID and a.CompID=b.CompID                  
                  
  Insert into BOM_MPS_MASTER                  
   Select a.ItemCode, a.StageCode, a.OperationCode, a.OperationSequence, a.RevisionNo, a.Status, a.Userid, a.Branchid, a.Compid, a.Transtamp                  
   From INSERTED a                  
   INNER JOIN           
    (                  
     Select ItemCode, StageCode, RevisionNo, BranchID, CompID,  Max(TranStamp) as MaxTranStamp                   
     From INSERTED                  
     Group by ItemCode, StageCode, RevisionNo, BranchID, CompID                  
    ) b                  
    On  b.ItemCode= a.ItemCode and b.StageCode=a.StageCode and b.RevisionNo=a.RevisionNo     
   and b.BranchID=a.BranchID and b.CompID=a.CompID and b.MaxTranStamp=a.TranStamp                    
  END                  
END 
 GO

/****** Object:  Trigger Script Date: 17/06/2026 11:50:45 ******/  
CREATE TRIGGER [dbo].[TGR_To_Insert_BOM_STRUCTURE_HDR] ON [dbo].[BOM_STRUCTURE_HDR_His]                    
For INSERT, UPDATE,DELETE                    
AS                    
BEGIN                    
 SET NOCOUNT ON;                    
                    
 If Exists (SELECT * FROM DELETED)                    
  BEGIN                    
   Delete From BOM_STRUCTURE_HDR                    
   From BOM_STRUCTURE_HDR a                    
   Inner Join DELETED b                    
    On a.ParentCode=b.ParentCode and a.RevisionNo=b.RevisionNo      
  and a.BranchID=b.BranchID and a.CompID=b.CompID  and a.TranStamp=b.TranStamp                    
  END                    
                    
 If Exists (SELECT * FROM INSERTED)                    
  BEGIN                    
            
  Delete From BOM_STRUCTURE_HDR                    
  From BOM_STRUCTURE_HDR a                    
  Inner Join INSERTED b                    
   On a.ParentCode=b.ParentCode  and a.RevisionNo=b.RevisionNo      
  and a.BranchID=b.BranchID and a.CompID=b.CompID                    
                    
  Insert into BOM_STRUCTURE_HDR                    
   Select a.ParentCode, a.BaseQty, a.RevisionNo, a.EffectiveDate, a.Status, a.Userid, a.Branchid, a.Compid, a.Transtamp                    
   From INSERTED a                    
   INNER JOIN             
    (                    
     Select ParentCode, RevisionNo, BranchID, CompID,  Max(TranStamp) as MaxTranStamp                     
     From INSERTED                    
     Group by ParentCode, RevisionNo, BranchID, CompID                    
    ) b                    
    On  b.ParentCode= a.ParentCode  and b.RevisionNo=a.RevisionNo       
   and b.BranchID=a.BranchID and b.CompID=a.CompID and b.MaxTranStamp=a.TranStamp                      
  END                    
END 
 GO

/****** Object:  Trigger Script Date: 17/06/2026 11:50:45 ******/
 CREATE TRIGGER [dbo].[TGR_To_Insert_BOM_STRUCTURE_DTL] ON [dbo].[BOM_STRUCTURE_DTL_His]                    
For INSERT, UPDATE,DELETE                    
AS                    
BEGIN                    
 SET NOCOUNT ON;                    
                    
 If Exists (SELECT * FROM DELETED)                    
  BEGIN                    
   Delete From BOM_STRUCTURE_DTL                    
   From BOM_STRUCTURE_DTL a                    
   Inner Join DELETED b                    
    On a.ParentCode=b.ParentCode and a.ChildCode=b.ChildCode and a.RevisionNo=b.RevisionNo      
  and a.BranchID=b.BranchID and a.CompID=b.CompID  and a.TranStamp=b.TranStamp                    
  END                    
                    
 If Exists (SELECT * FROM INSERTED)                    
  BEGIN                    
            
  Delete From BOM_STRUCTURE_DTL                    
  From BOM_STRUCTURE_DTL a                    
  Inner Join INSERTED b                    
   On a.ParentCode=b.ParentCode and a.ChildCode=b.ChildCode and a.RevisionNo=b.RevisionNo      
  and a.BranchID=b.BranchID and a.CompID=b.CompID                    
                    
  Insert into BOM_STRUCTURE_DTL                    
   Select a.ItemGroupKey, a.ParentCode, a.ChildCode, a.Qty, a.RevisionNo, a.Explodable, a.Percentage, a.Priority, a.Status, a.Userid, a.Branchid, a.Compid, a.Transtamp                    
   From INSERTED a                    
   INNER JOIN             
    (                    
     Select ParentCode, ChildCode, RevisionNo, BranchID, CompID,  Max(TranStamp) as MaxTranStamp                     
     From INSERTED                    
     Group by ParentCode, ChildCode, RevisionNo, BranchID, CompID                    
    ) b                    
    On  b.ParentCode= a.ParentCode and b.ChildCode= a.ChildCode and b.RevisionNo=a.RevisionNo       
   and b.BranchID=a.BranchID and b.CompID=a.CompID and b.MaxTranStamp=a.TranStamp                      
  END                    
END 
GO

/****** Object:  Trigger Script Date: 17/06/2026 11:50:45 ******/
CREATE TRIGGER [dbo].[TGR_To_Insert_QMS_V1_StandardMaster_Dtl]           
   ON  dbo.QMS_V1_StandardMaster_Dtl_Rep            
FOR INSERT, DELETE, UPDATE          
AS           
BEGIN            
    SET NOCOUNT ON;          
      
  /* Insert Current Inserted/Deleted Data Into Temp Table */    
    Select * INTO #TmpTbl_StdMastDtl     
    From      
    (      
        Select * From INSERTED        
        Union       
        Select * From DELETED        
    ) a      
      
  /* Delete Already Existing Data Into QMS_V1_StandardMaster_Dtl_Rep Table */    
    Delete From QMS_V1_StandardMaster_Dtl         
    From QMS_V1_StandardMaster_Dtl a          
    Inner Join #TmpTbl_StdMastDtl b                  
   On a.StdID=b.StdID and a.StdClauseID=b.StdClauseID and a.StdLinkID=b.StdLinkID  
    and a.StdLinkClauseID=b.StdLinkClauseID and a.BranchID=b.BranchID and a.CompID=b.CompID        
              
  /* Insert Current Inserted/Deleted Data Into QMS_V1_StandardMaster_Dtl_Rep Table */    
    ;With StdMastDtl_CTE      
    As      
    (Select a.StdID, a.StdClauseID, a.IsClauseApplicable, a.StdLinkID, a.StdLinkClauseID, a.IsLinkApplicable,   
    a.EffStatus, a.UserID, a.BranchID, a.CompID, a.TranStamp  
   From QMS_V1_StandardMaster_Dtl_Rep a         
    Where Exists( Select * From #TmpTbl_StdMastDtl b                       
      Where a.StdID=b.StdID and a.StdClauseID=b.StdClauseID and a.StdLinkID=b.StdLinkID   
        and a.StdLinkClauseID=b.StdLinkClauseID and a.BranchID=b.BranchID and a.CompID=b.CompID             
      )      
    ) Insert Into QMS_V1_StandardMaster_Dtl         
        Select a.StdID, a.StdClauseID, a.IsClauseApplicable, a.StdLinkID, a.StdLinkClauseID, a.IsLinkApplicable,   
      a.EffStatus, a.UserID, a.BranchID, a.CompID, b.MaxTranStamp As TranStamp     
        From StdMastDtl_CTE a          
        Inner Join( Select a.StdID, a.StdClauseID, a.StdLinkID, a.StdLinkClauseID,  
           Max(a.TranStamp) As MaxTranStamp, a.BranchID, a.CompID     
         From StdMastDtl_CTE a           
         Group by a.StdID, a.StdClauseID, a.StdLinkID, a.StdLinkClauseID, a.BranchID, a.CompID      
        ) b        
    On a.StdID=b.StdID and a.StdClauseID=b.StdClauseID and a.StdLinkID=b.StdLinkID   
     and a.StdLinkClauseID=b.StdLinkClauseID and a.BranchID=b.BranchID and a.CompID=b.CompID   
     and a.TranStamp=b.MaxTranStamp    
    
 /* Drop Temp Table */    
  DROP TABLE #TmpTbl_StdMastDtl;            
END  
GO

/****** Object:  Trigger Script Date: 17/06/2026 11:50:45 ******/
 CREATE TRIGGER [dbo].[TGR_To_Insert_BOM_ItemwiseMachinewiseCycleTime] ON [dbo].[BOM_ItemwiseMachinewiseCycleTime_His]                      
For INSERT, UPDATE,DELETE                      
AS                      
BEGIN                      
 SET NOCOUNT ON;                      
                      
 If Exists (SELECT * FROM DELETED)                      
  BEGIN                      
   Delete From BOM_ItemwiseMachinewiseCycleTime                      
   From BOM_ItemwiseMachinewiseCycleTime a                      
   Inner Join DELETED b                      
    On a.ItemCode=b.ItemCode and a.MachineCode=b.MachineCode       
  and a.BranchID=b.BranchID and a.CompID=b.CompID  and a.TranStamp=b.TranStamp                      
  END                      
                      
 If Exists (SELECT * FROM INSERTED)                      
  BEGIN                      
              
  Delete From BOM_ItemwiseMachinewiseCycleTime                      
  From BOM_ItemwiseMachinewiseCycleTime a                      
  Inner Join INSERTED b                      
   On a.ItemCode=b.ItemCode and a.MachineCode=b.MachineCode and a.BranchID=b.BranchID and a.CompID=b.CompID                      
                      
  Insert into BOM_ItemwiseMachinewiseCycleTime                      
   Select a.ItemCode, a.MachineCode, a.CycleTime, a.OperationCost, a.SettingTime, a.Priority, a.Status, a.Userid, a.Branchid, a.Compid, a.Transtamp                      
   From INSERTED a                      
   INNER JOIN               
    (                      
     Select ItemCode, MachineCode, BranchID, CompID,  Max(TranStamp) as MaxTranStamp                       
     From INSERTED                      
     Group by ItemCode, MachineCode,  BranchID, CompID                      
    ) b                      
    On  b.ItemCode= a.ItemCode and b.MachineCode= a.MachineCode       
  and b.BranchID=a.BranchID and b.CompID=a.CompID and b.MaxTranStamp=a.TranStamp                        
  END                      
END 
GO

/****** Object:  Trigger Script Date: 17/06/2026 11:50:45 ******/
CREATE TRIGGER [dbo].[TGR_To_Insert_QMS_V1_StandardMaster_Hdr]           
   ON  dbo.QMS_V1_StandardMaster_Hdr_Rep            
FOR INSERT, DELETE, UPDATE          
AS           
BEGIN            
    SET NOCOUNT ON;          
      
  /* Insert Current Inserted/Deleted Data Into Temp Table */    
    Select * INTO #TmpTbl_StdMastHdr     
    From      
    (      
        Select * From INSERTED        
        Union       
        Select * From DELETED        
    ) a      
  
  /* Delete Data In Table */    
    Delete From QMS_V1_StandardMaster_Hdr          
    From QMS_V1_StandardMaster_Hdr a          
    Inner Join #TmpTbl_StdMastHdr b            
   On a.StdID=b.StdID and a.BranchID=b.BranchID and a.CompID=b.CompID       
  
  /* Insert Current Inserted/Deleted Data Into Table */    
    ;With StdMastHdr_CTE      
    As      
    (Select a.StdID, a.StdName, a.EffStatus, a.UserID, a.BranchID, a.CompID, a.TranStamp  
   From QMS_V1_StandardMaster_Hdr_Rep a         
    Where Exists( Select * From #TmpTbl_StdMastHdr b                       
      Where a.StdID=b.StdID and a.BranchID=b.BranchID and a.CompID=b.CompID             
      )      
    ) Insert Into QMS_V1_StandardMaster_Hdr          
        Select a.StdID, a.StdName, a.EffStatus, a.UserID, a.BranchID, a.CompID, b.MaxTranStamp As TranStamp   
        From StdMastHdr_CTE a          
        Inner Join( Select a.StdID, a.StdName, Max(a.TranStamp) As MaxTranStamp,   
          a.BranchID, a.CompID    
         From StdMastHdr_CTE a           
         Group by a.StdID, a.StdName, a.BranchID, a.CompID      
        ) b        
    On a.StdID=b.StdID and a.BranchID=b.BranchID and a.CompID=b.CompID       
     and a.TranStamp=b.MaxTranStamp    
    
 /* Drop Temp Table */    
  DROP TABLE #TmpTbl_StdMastHdr;            
END  
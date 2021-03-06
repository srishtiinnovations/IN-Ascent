set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go


--exec ProcessSheet_Report 4469
ALTER procedure [dbo].[ProcessSheet_Report]
(
   @PordNo Numeric(10))
   
as
begin
	
	Select T0.DocNum,T0.ItemCode,T2.ItemName,T0.PostDate,T0.DueDate,T0.PlannedQty ,
	T1.U_Rteid, Case When T3.U_Defrte='Y' then 'Yes' when T3.U_Defrte='N' then 'No' End as 'DefltRte',
	T1.U_Seqnce,T1.U_OprCode,((T1.U_OprCode) + '-' + (U_Oprname)) as 'Operation',
	(((T4.U_Opertime)/(T4.U_perqty))* (T0.PlannedQty)) as 'EstmtTim' from OWOR T0
	Inner Join [@PSSIT_WOR2] T1 on T1.U_Pordno =T0.DocNum
	inner Join OITM T2 on T2.ItemCode=T0.ItemCode
	Inner Join [@PSSIT_ORTE] T3 on T3.Code=T1.U_Rteid
	Inner Join [@PSSIT_RTE1] T4 on T4.U_Rteid=T3.Code 
	where T0.Itemcode=T3.U_ItemCode and 
	T1.U_OprCode=T4.U_OprCode and T0.Docnum=@PordNo

end





























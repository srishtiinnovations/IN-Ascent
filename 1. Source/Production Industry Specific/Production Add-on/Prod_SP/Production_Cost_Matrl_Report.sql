set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go







--exec Production_Cost_Matrl_Report '4469','4469','R'
ALTER procedure [dbo].[Production_Cost_Matrl_Report]
(
    @FPordNo Numeric(20),
    @TPordNo Numeric(20),
    @Status  nvarchar(10) )

as
begin

		select Tbl2.DocNum,Tbl2.ItemCode,Tbl2.PlannedQty,Tbl2.MtrlPlanVal,Tbl2.IssuedQty,Tbl2.MtrlActVal,Tbl2.MtrlVarQty,
		Tbl2.MtrlVarVal, Case When Tbl2.MtrlVarValP >=100 Then 100 else Tbl2.MtrlVarValP End as 'MtrlVarValP' from 
			 (select Tbl1.DocNum,Tbl1.ItemCode,Tbl1.PlannedQty,Tbl1.MtrlPlanVal,Tbl1.IssuedQty,Tbl1.MtrlActVal,Tbl1.MtrlVarQty,
			 Tbl1.MtrlVarVal, Case When Tbl1.MtrlPlanVal > 0 then((Tbl1.MtrlVarVal/Tbl1.MtrlPlanVal) * 100) End as 'MtrlVarValP' from 
				(select Tbl.DocNum,Tbl.ItemCode,Tbl.PlannedQty,Tbl.MtrlPlanVal,Tbl.IssuedQty,Tbl.MtrlActVal,Tbl.MtrlVarQty,
				((Tbl.MtrlActVal)-(MtrlPlanVal)) as 'MtrlVarVal' from 
					(select T0.DocNum,(T3.ItemCode + '-' + T1.ItemName) 
					as 'ItemCode',T3.PlannedQty,((T3.PlannedQty) * (T2.AvgPrice)) 
					as 'MtrlPlanVal',T3.IssuedQty,((T3.IssuedQty) * (T2.AvgPrice)) as 'MtrlActVal',
					((T3.IssuedQty)-(T3.PlannedQty)) as 'MtrlVarQty' from OWOR T0
					inner join OITM T1 on T1.ItemCode=T0.Itemcode
					inner Join OITW T2 on T2.ItemCode=T1.ItemCode
					Inner Join WOR1 T3 on T3.DocEntry=T0.DocEntry
					where T2.WhsCode=T3.wareHouse and  T0.Docnum>=@FPordNo 
					and T0.Docnum<=@TPordNo and T0.Status=@Status 
		) Tbl ) Tbl1) Tbl2

end






























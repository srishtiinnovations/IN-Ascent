set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go






--exec Production_Cost_FxdCst_Report '4469','4469','R'
ALTER procedure [dbo].[Production_Cost_FxdCst_Report]
(
    @FPordNo Numeric(20),
    @TPordNo Numeric(20),
    @Status  nvarchar(10) )

as
begin

		Select Tbl2.DocNum,Tbl2.U_FCost,Tbl2.FxdPlnVal,Tbl2.FxdActVal,Tbl2.FxdVarVal,
		Case When Tbl2.FxdVarValP >= 100 then 100 Else Tbl2.FxdVarValP End as 'FxdVarValP' From 
			(Select Tbl1.DocNum,Tbl1.U_FCost,Tbl1.FxdPlnVal,Tbl1.FxdActVal,Tbl1.FxdVarVal,
			Case When Tbl1.FxdPlnVal > 0 then(((Tbl1.FxdVarVal) / (Tbl1.FxdPlnVal)) * 100) End as 'FxdVarValP'  From 
				(Select Tbl.DocNum,Tbl.U_FCost,Tbl.FxdPlnVal,Tbl.FxdActVal,((Tbl.FxdActVal) - (Tbl.FxdPlnVal)) 
				as 'FxdVarVal' From
					(Select T0.DocNum,T1.U_FCost,Sum((T0.PlannedQty) * (T1.U_UnitCost)) as 'FxdPlnVal',
					Sum((T1.U_Totfcst) + (T1.U_Adnl1)) as 'FxdActVal'  from OWOR T0
					Inner Join [@PSSIT_WOR4] T1 on T1.U_Pordno=T0.DocNum
				    where T0.Docnum>=@FPordNo and  T0.Docnum<=@TPordNo and T0.Status=@Status 
					Group by  T0.DocNum,T1.U_FCost,T0.PlannedQty,T1.U_UnitCost,T1.U_Totfcst,T1.U_Adnl1
		) Tbl ) Tbl1) Tbl2

end




























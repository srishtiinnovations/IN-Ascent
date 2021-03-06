set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go


--exec Production_Cost_Tool_Report '4469','4469','R'
ALTER procedure [dbo].[Production_Cost_Tool_Report]
(
    @FPordNo Numeric(20),
    @TPordNo Numeric(20),
    @Status  nvarchar(10) )

as
begin

		Select Tbl3.DocNum,Tbl3.Tools,Tbl3.U_Planqty,Tbl3.TLPlnVal,
		Tbl3.U_Qty,Tbl3.U_TotCost,Tbl3.TLVarStrk,Tbl3.TLVarVal,
		Case when Tbl3.TLVarValP >= 100 then 100 Else Tbl3.TLVarValP End as 'TLVarValP' From
			(Select Tbl2.DocNum,Tbl2.Tools,Tbl2.U_Planqty,Tbl2.TLPlnVal,
			Tbl2.U_Qty,Tbl2.U_TotCost,Tbl2.TLVarStrk,Tbl2.TLVarVal,
			Case When Tbl2.TLPlnVal > 0 then (((Tbl2.TLVarVal) / (Tbl2.TLPlnVal)) * 100) End as 'TLVarValP' From
				(Select Tbl1.DocNum,Tbl1.Tools,Tbl1.U_Planqty,Tbl1.TLPlnVal,
				Tbl1.U_Qty,Tbl1.U_TotCost,Tbl1.TLVarStrk,((Tbl1.U_TotCost) -(Tbl1.TLPlnVal)) as 'TLVarVal' from
					(Select Tbl.DocNum,Tbl.Tools,Tbl.U_Planqty,Tbl.TLPlnVal,
					Tbl.U_Qty,Tbl.U_TotCost,Tbl.TLVarStrk from
						(select T0.DocNum,((T2.U_Toolcode) + '-' + (T2.U_TLname)) as 'Tools',T1.U_Planqty,
						((T1.U_Planqty) * (T3.U_Cpno)) as 'TLPlnVal',T2.U_Qty,T2.U_TotCost,
						((T2.U_Qty) - (T1.U_Planqty)) as 'TLVarStrk'  from OWOR T0
						Inner Join [@PSSIT_OPEY] T1 on T1.U_WORNo=T0.DocNum
						Inner Join [@PSSIT_PEY3] T2 on T2.U_Prdentno=T1.DocNum
						Inner Join [@PSSIT_OTLS] T3 on T3.Code=T2.U_ToolCode
					    where T0.Docnum>=@FPordNo and  T0.Docnum<=@TPordNo and T0.Status=@Status 
		) Tbl ) Tbl1) Tbl2) Tbl3

end



























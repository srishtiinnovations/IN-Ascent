set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go




--exec Production_Cost_ManPow_Report '4469','4469','R'
ALTER procedure [dbo].[Production_Cost_ManPow_Report]
(
    @FPordNo Numeric(20),
    @TPordNo Numeric(20),
    @Status  nvarchar(10) )

as
begin

		select Tbl2.DocNum,Tbl2.Emp,Tbl2.MPPlnTime,Tbl2.MPPlnVal,Tbl2.U_WrkTime,Tbl2.U_Totcost,
		Tbl2.MPVarTime,Tbl2.MPVarVal,Case When Tbl2.MPPlnVal > 0 then (((Tbl2.MPVarVal)/(Tbl2.MPPlnVal)) * 100) End 
        as 'MPVarValP' from
			(Select Tbl1.DocNum,Tbl1.Emp,Tbl1.MPPlnTime,Tbl1.MPPlnVal,Tbl1.U_WrkTime,Tbl1.U_Totcost,
			((Tbl1.U_WrkTime) - (Tbl1.MPPlnTime)) as 'MPVarTime',((Tbl1.U_Totcost) - (Tbl1.MPPlnVal)) as 'MPVarVal' from
				(select Tbl.DocNum,Tbl.Emp,Tbl.MPPlnTime,((Tbl.MPPlnVal) * (Tbl.MPPlnTime)) as 'MPPlnVal',
				Tbl.U_WrkTime,Tbl.U_Totcost  from
					(select T0.Docnum,((T4.U_Empid) + '-' + (T4.U_Empnam)) as 'Emp', ((T3.U_Reqtime) * (T3.U_Reqno)) 
					as 'MPPlnTime', ((T4.U_Labrate)/60) as 'MPPlnVal',T2.U_WrkTime,T2.U_Totcost from OWOR T0
					inner Join [@PSSIT_OPEY] T1 on T1.U_WORNo=T0.DocNum
					inner Join [@PSSIT_PEY2] T2 on T2.U_Prdentno=T1.DocNum
					inner Join [@PSSIT_RTE2] T3 on T3.U_Skilgrp=T2.U_LGCode
					inner Join [@PSSIT_OLBR] T4 on T4.Code=T2.U_Lrcode
					inner Join [@PSSIT_ORTE] T5 on T5.code=T3.U_Rteid 
					inner join [@PSSIT_WOR2] T6 on T6.U_Pordno=T0.Docnum
					Where T6.U_Rteid=T5.Code and  T6.U_Oprcode=T3.U_Oprcode 
					and  T0.Docnum>=@FPordNo 
					and T0.Docnum<=@TPordNo and T0.Status=@Status 
		) Tbl) Tbl1) Tbl2

end




























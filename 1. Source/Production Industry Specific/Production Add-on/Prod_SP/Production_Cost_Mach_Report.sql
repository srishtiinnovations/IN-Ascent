set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go
--exec Production_Cost_Mach_Report '28','28','R'
ALTER procedure [dbo].[Production_Cost_Mach_Report]
(
    @FPordNo Numeric(20),
    @TPordNo Numeric(20),
    @Status  nvarchar(10) )
as
begin

		select Tbl2.DocNum,Tbl2.Mach,Tbl2.MachPlnTime,Tbl2.MachPlnVal,Tbl2.U_Rntime,Tbl2.U_Qty,
		Tbl2.MachVarTime,Tbl2.MachVarVal, Case When  Tbl2.MachVarVal >=100 Then 100 
		else Tbl2.MachVarVal End as 'MachVarValP' From
			(select Tbl1.DocNum,Tbl1.Mach,Tbl1.MachPlnTime,Tbl1.MachPlnVal,Tbl1.U_Rntime,Tbl1.U_Qty,
			Tbl1.MachVarTime,Tbl1.MachVarVal, Case When Tbl1.MachPlnVal > 0 then (((Tbl1.MachVarVal) / (Tbl1.MachPlnVal)) * 100) 
			End as 'MachVarValP' from 
				(select Tbl.DocNum,Tbl.Mach,Tbl.MachPlnTime,Tbl.MachPlnVal,Tbl.U_Rntime,Tbl.U_Qty,
				((Tbl.U_Rntime) - (Tbl.MachPlnTime)) as 'MachVarTime',((Tbl.U_Qty) - (Tbl.MachPlnVal)) 
				as 'MachVarVal' from
					(select T0.DocNum,(T2.U_wcno + '-' + T2.U_wcname) as 'Mach',T2.U_type,
					Case when T2.U_type='Operation Time' then T3.U_Opertime 
					when T2.U_type='Setup Time' then T3.U_Setime  End as 'MachPlnTime',
					Case when T2.U_type='Operation Time' then ((T4.U_opercost/60) *  T3.U_Opertime )
					when T2.U_type='Setup Time' then ((T4.U_Setupcost/60) * T3.U_Setime)   End as 'MachPlnVal',
					Sum(T2.U_Rntime) as 'U_Rntime',Sum(T2.U_Qty) as 'U_Qty'  from OWOR T0
					inner Join [@PSSIT_OPEY] T1 on T1.U_WORNo=T0.DocNum
					inner Join [@PSSIT_PEY1] T2 on T2.DocEntry=T1.DocEntry
					inner Join [@PSSIT_RTE1] T3 on T3.U_wcno=T2.U_wcno
					inner Join [@PSSIT_PMWCHDR] T4 on T4.U_wcno=T2.U_wcno
				    inner join [@PSSIT_WOR2] T5 on T5.U_Pordno=T0.Docnum
					where T5.U_Oprcode=T3.U_Oprcode and 
                    T5.U_Rteid =T3.U_Rteid and 
					T0.Docnum>=@FPordNo and  T0.Docnum<=@TPordNo and T0.Status=@Status
					Group by T0.DocNum,T2.U_wcno,T2.U_wcname,T2.U_type,T3.U_Opertime ,
					T4.U_opercost,T4.U_Setupcost,T3.U_Setime   
		) Tbl ) Tbl1) Tbl2

end






























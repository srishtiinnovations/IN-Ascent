set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go





--exec [ProcessSheet_Mach_Report] 15
ALTER procedure [dbo].[ProcessSheet_Mach_Report]
(
   @PordNo Numeric(10))
   
as
begin	
--Select Code,((U_wcno) + '-' + (U_wcname)) as 'Machine' from [@PSSIT_PRN1]

Select T0.DocNum,T1.U_OprCode,((T5.U_wcno) + '-' + (T5.U_wcname)) as 'Machine' from OWOR T0
Inner Join [@PSSIT_WOR2] T1 on T1.U_Pordno =T0.DocNum
inner Join OITM T2 on T2.ItemCode=T0.ItemCode
Inner Join [@PSSIT_ORTE] T3 on T3.Code=T1.U_Rteid
Inner Join [@PSSIT_RTE1] T4 on T4.U_Rteid=T3.Code 
inner join [@PSSIT_PRN1] T5 on T5.Code=T1.U_OprCode
where T0.Itemcode=T3.U_ItemCode and T0.Docnum=@PordNo
Group by T0.DocNum,T1.U_OprCode,T5.U_wcno,T5.U_wcname

end


























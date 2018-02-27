alter procedure [dbo].[Production_Report]
(
    @FDate datetime,
    @TDate datetime)
   
as
begin
---Modified by Manimaran-----s	
--select T0.ItemCode,T1.ItemName,T0.Docnum,Case When T0.Status='R' then 'Released' 
--When T0.Status='C' then 'Closed' End as 'Status', T0.PostDate,T0.DueDate,T0.PlannedQty,
--T2.U_Docdt,((T2.U_Scode) + '-' + (T2.U_Sdesc)) as 'Shift', ((T2.U_Oprcode) + '-' + (T2.U_Oprname)) 
--as 'Operation',T2.U_ProdQty,T2.U_Passqty,T2.U_Rewrkqty,T2.U_scrapqty
-- from OWOR T0
--inner join OITM T1 on T1.ItemCode=T0.ItemCode
--inner join [@PSSIT_OPEY] T2 on T2.U_WORNo=T0.DocNum
--Where T2.U_DocDt>=@FDate and T2.U_DocDt<=@TDate
----Group by T0.ItemCode,T1.ItemName,T0.Docnum,T0.Status, T0.PostDate,
----T0.DueDate,T0.PlannedQty,T2.U_Docdt,T2.U_Scode,T2.U_Sdesc,T2.U_Oprcode,
----T2.U_Oprname,T2.U_ProdQty,T2.U_Passqty,T2.U_Rewrkqty,T2.U_scrapqty
set dateformat dmy
select T0.ItemCode,T1.ItemName,T0.Docnum,Case When T0.Status='R' then 'Released' 
When T0.Status='C' then 'Closed' End as 'Status', T0.PostDate,T0.DueDate,T0.PlannedQty,
T2.U_Docdt,((T2.U_Scode) + '-' + (T2.U_Sdesc)) as 'Shift', ((T2.U_Oprcode) + '-' + (T2.U_Oprname)) 
as 'Operation',T2.U_ProdQty,T2.U_Passqty,sum(isnull(T3.U_Rewrkqty,0)) U_Rewrkqty,sum(isnull(T4.U_scrapqty,0)) U_scrapqty
 from OWOR T0
inner join OITM T1 on T1.ItemCode=T0.ItemCode
inner join [@PSSIT_OPEY] T2 on T2.U_WORNo=T0.DocNum
inner join [@PSSIT_PEY5] t3 on t3.DocEntry = t2.DocEntry 
inner join [@PSSIT_PEY6] t4 on t4.DocEntry = t2.DocEntry 
Where T2.U_DocDt>= @FDate and T2.U_DocDt<= @TDate
Group by T0.ItemCode,T1.ItemName,T0.Docnum,T0.Status, T0.PostDate,
T0.DueDate,T0.PlannedQty,T2.U_Docdt,T2.U_Scode,T2.U_Sdesc,T2.U_Oprcode,
T2.U_Oprname,T2.U_ProdQty,T2.U_Passqty
end

























GO



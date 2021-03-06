set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go



--exec Production_Cost_Report 0,0,'C'
ALTER procedure [dbo].[Production_Cost_Report]
(
    @FPordNo Numeric(20),
    @TPordNo Numeric(20),
    @Status  nvarchar(10) )

as
begin

	If @FPordNo>0 and @TPordNo>0 and Len(@Status)>0 
			select T0.Docnum,T0.ItemCode,T1.ItemName,
			Case When T0.Status='R' then 'Released' When 
			T0.Status='C' then 'Closed' End as 'Status',
			T0.PostDate,T0.DueDate,T0.PlannedQty,
			T0.CmpltQty,T0.RjctQty 
			from OWOR T0
			inner join OITM T1 on T1.ItemCode=T0.Itemcode
			where T0.Docnum>=@FPordNo and T0.Docnum<=@TPordNo and T0.Status=@Status
	Else if @FPordNo>0 and @TPordNo>0 and Len(@Status)=0
			select T0.Docnum,T0.ItemCode,T1.ItemName,
			Case When T0.Status='R' then 'Released' When 
			T0.Status='C' then 'Closed' End as 'Status',
			T0.PostDate,T0.DueDate,T0.PlannedQty,
			T0.CmpltQty,T0.RjctQty 
			from OWOR T0
			inner join OITM T1 on T1.ItemCode=T0.Itemcode
			where T0.Docnum>=@FPordNo and T0.Docnum<=@TPordNo 
	Else if @FPordNo>0 and @TPordNo=0 and Len(@Status)=0
			select T0.Docnum,T0.ItemCode,T1.ItemName,
			Case When T0.Status='R' then 'Released' When 
			T0.Status='C' then 'Closed' End as 'Status',
			T0.PostDate,T0.DueDate,T0.PlannedQty,
			T0.CmpltQty,T0.RjctQty 
			from OWOR T0
			inner join OITM T1 on T1.ItemCode=T0.Itemcode
			where T0.Docnum=@FPordNo 
    Else if @FPordNo=0 and @TPordNo>0 and Len(@Status)=0
			select T0.Docnum,T0.ItemCode,T1.ItemName,
			Case When T0.Status='R' then 'Released' When 
			T0.Status='C' then 'Closed' End as 'Status',
			T0.PostDate,T0.DueDate,T0.PlannedQty,
			T0.CmpltQty,T0.RjctQty 
			from OWOR T0
			inner join OITM T1 on T1.ItemCode=T0.Itemcode
			where  T0.Docnum=@TPordNo 							
    Else if @FPordNo=0 and @TPordNo=0 and Len(@Status)>0
			select T0.Docnum,T0.ItemCode,T1.ItemName,
			Case When T0.Status='R' then 'Released' When 
			T0.Status='C' then 'Closed' End as 'Status',
			T0.PostDate,T0.DueDate,T0.PlannedQty,
			T0.CmpltQty,T0.RjctQty 
			from OWOR T0
			inner join OITM T1 on T1.ItemCode=T0.Itemcode
			where  T0.Status=@Status
    Else if @FPordNo=0 and @TPordNo=0 and Len(@Status)=0
			select T0.Docnum,T0.ItemCode,T1.ItemName,
			Case When T0.Status='R' then 'Released' When 
			T0.Status='C' then 'Closed' End as 'Status',
			T0.PostDate,T0.DueDate,T0.PlannedQty,
			T0.CmpltQty,T0.RjctQty 
			from OWOR T0
			inner join OITM T1 on T1.ItemCode=T0.Itemcode
end
























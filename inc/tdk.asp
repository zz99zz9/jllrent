<%tdksql="select * From [tdk] where id="&tdkid
Set tdkrs= Server.CreateObject("ADODB.Recordset")
tdkrs.open tdksql,conn,1,1
t=tdkrs("t")
d=tdkrs("d")
k=tdkrs("k")
tdkrs.close
set tdkrs=nothing
%>

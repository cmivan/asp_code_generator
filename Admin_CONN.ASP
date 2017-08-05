<%
dim conns,connstr,time1,time2,mdb
time1=timer   
  mdb="fanren_db/8fbdc9ab21bf4a61.mdb"
    on error resume next
    connstr="driver=microsoft access driver (*.mdb);dbq="+server.mappath(mdb)
set conns=server.createobject("adodb.connection") 
    conns.open connstr
if err then
   err.clear
set conns = nothing
   response.write "暂时无法读取数据..."
   response.end
end if
%>
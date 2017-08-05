<%
option explicit
response.buffer=true
numcode
function numcode()
	response.expires = -1
	response.addheader "pragma","no-cache"
	response.addheader "cache-ctrol","no-cache"
	on error resume next
	dim znum,i,j
	dim ados,ados1
	randomize timer
	znum = cint(8999*rnd+1000)
	session("checkcode") = znum
	dim zimg(4),nstr
	nstr=cstr(znum)
	for i=0 to 3
		zimg(i)=cint(mid(nstr,i+1,1))
	next
	dim pos
	set ados=server.createobject("adodb.stream")
	ados.mode=3
	ados.type=1
	ados.open
	set ados1=server.createobject("adodb.stream")
	ados1.mode=3
	ados1.type=1
	ados1.open
	ados.loadfromfile(server.mappath("body.fix"))
	ados1.write ados.read(1280)
	for i=0 to 3
		ados.position=(9-zimg(i))*320
		ados1.position=i*320
		ados1.write ados.read(320)
	next	
	ados.loadfromfile(server.mappath("head.fix"))
	pos=lenb(ados.read())
	ados.position=pos
	for i=0 to 9 step 1
		for j=0 to 3
			ados1.position=i*32+j*320
			ados.position=pos+30*j+i*120
			ados.write ados1.read(30)
		next
	next
	response.contenttype = "image/bmp"
	ados.position=0
	response.binarywrite ados.read()
	ados.close:set ados=nothing
	ados1.close:set ados1=nothing
	if err then session("checkcode") = 9999
end function
'asp code created by blueidea.com web team v37 2003-7-25
%>
Set fso = CreateObject("Scripting.FileSystemObject")
target_dir=fso.GetAbsolutePathName(".")+"\report"
filename = fso.GetAbsolutePathName(".")+"\2019_01_31_17_13_0000_datalog.csv"
separator=","

annee=0
mois=0
jour=0
artice=""



Set f = fso.OpenTextFile(filename)
Set f_model = fso.OpenTextFile("./modele/modele.html")
Set f_not_repeated = fso.OpenTextFile("./modele/not_repeated.txt")
Set f_repeated = fso.OpenTextFile("./modele/repeated_body.txt")

Set shell = CreateObject("Wscript.Shell")

report_content=f_model.ReadAll()
f_model.close()


first_columns=split(f.Readline,separator)
first_line=split(f.Readline,separator)


localdate=first_line(GetColumnIndex("LocalDate"))
annee=split(localdate,"/")(2)
mois=split(localdate,"/")(1)
jour=split(localdate,"/")(0)
article=first_line(GetColumnIndex("Nom d'article"))



Do Until f_not_repeated.AtEndOfStream
  var = f_not_repeated.ReadLine
  report_content=replace(report_content,"["+var+"]",first_line(GetColumnIndex(var)))
Loop
f_not_repeated.close()


table="<table border=1><tr>"

Do Until f_repeated.AtEndOfStream
  var = f_repeated.ReadLine
  vars=split(var,",")
  var_id = vars(0)
  var_name = vars(1)
  var_align = vars(2)
  table=table+"<td style='text-align: "+ var_align +"'>"+var_name+"</td>"
Loop
table=table+"</tr>"


f.close()

Set f = fso.OpenTextFile(filename)
if not f.AtEndOfStream then f.ReadLine
Do Until f.AtEndOfStream
  f_repeated.close()
  Set f_repeated = fso.OpenTextFile("./modele/repeated_body.txt")
  line=f.ReadLine
  line=replace(line,"""","")
  table=table+"<tr>"
  Do Until f_repeated.AtEndOfStream
    var = f_repeated.ReadLine
    vars=split(var,",")
    var_id = vars(0)
    var_name = vars(1)
    var_align = vars(2)
    table=table+"<td style='text-align: "+ var_align +"'>"+split(line,separator)(GetColumnIndex(var_id))+"</td>"
  Loop
  table=table+"</tr>"
Loop
f_repeated.close()
table=table+"</table>"


report_content=replace(report_content,"[table]",table)

save_file annee,mois,jour,article

print(GetColumnIndex("Vitesse Agitation (%)"))
'for each col in columns
'	WScript.Echo col
'next 
'Do Until f.AtEndOfStream
'  WScript.Echo f.ReadLine
'Loop

f.Close

function print(text)
   WScript.Echo text
End Function


function GetColumnIndex(col_name)
   i=0
   for each col in first_columns
      i=i+1
	  'print("compare "+col+ " "+ col_name)
	  if col = col_name then 
		GetColumnIndex=i-1
		Exit function
	  end if
   next
   GetColumnIndex=0
End Function

function save_file(an,mois,day,art)
   tdir=target_dir+"\"+cstr(an)+"\"+right("00"+cstr(mois),2)+"\"+right("00"+cstr(day),2)
   If not fso.FolderExists(tdir) Then 
      print(tdir)
      shell.Run "cmd /c mkdir "+tdir,0,true
   End if
   outFile=tdir+"\"+cstr(an)+"-"+right("00"+cstr(mois),2)+"-"+right("00"+cstr(day),2)+"-"+art+".html"
   print(outFile)
   Set objFile = fso.CreateTextFile(outFile,True)
   objFile.Write report_content & vbCrLf
   objFile.Close
   print_html_file(outFile)
End Function

sub print_html_file(file_name)
	surl=file_name

	dim bpttd_ready, istatus
	set oie=wscript.createobject("internetexplorer.application","ie_")
	do while oie.readystate<>4 : wscript.sleep 50 : loop
	on error resume next
	istatus=oie.querystatuswb(6)
	if err.number<>0 then
		wscript.echo "Cannot find the printer. Operation aborted."
		oie.quit
		set oie=nothing
		wscript.quit err.number
	end if
	on error goto 0
	oie.navigate surl
	do while oie.readystate<>4 : wscript.sleep 50 : loop
	bpttd_ready=false
	oie.execwb 6,2
	do while not bpttd_ready : wscript.sleep 50 : loop
	oie.quit
	set oie=nothing
	wscript.quit

end sub


sub ie_PrintTemplateTeardown(pDisp)
    bpttd_ready=true    'global bpttd_ready; no dim here
end sub
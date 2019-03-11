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


first_columns=split(getLine(f.Readline),separator)
'first_line=split(f.Readline,separator)
'for each co in first_columns
'  print(co)
'next

lot=""
heure_depart=""
heure_fin=""
date_report=""
utilisateur=""
article=""

flag=""
Do Until f.AtEndOfStream
  'print(getLine(f.Readline))
  current_line = split(getLine(f.Readline),separator)
  vlocaldate=current_line(GetColumnIndex("LocalDate"))
  varticle=current_line(GetColumnIndex("Nom d'article"))
  vheure=current_line(GetColumnIndex("LocalTime"))
  vutilisateur=current_line(GetColumnIndex("Nom d'utilisateur"))
  vlot=current_line(GetColumnIndex("N. Cycle"))
  flag=current_line(GetColumnIndex("Flag"))
  if flag="0" then exit do
  if date_report="" then
    date_report=vlocaldate
    annee=split(vlocaldate,"/")(2)
	mois=split(vlocaldate,"/")(1)
	jour=split(vlocaldate,"/")(0)
  end if
  if article="" then article=varticle
  if lot="" then lot=vlot
  if heure_debut="" then heure_debut=vheure
  if vheure<>"" then heure_fin=vheure
  if utilisateur="" then utilisateur=vutilisateur
Loop

if flag="1" then WScript.quit

report_content=replace(report_content,"[N. Cycle]",lot)
report_content=replace(report_content,"[heure_debut]",heure_debut)
report_content=replace(report_content,"[heure_fin]",heure_fin)
report_content=replace(report_content,"[LocalDate]",date_report)
report_content=replace(report_content,"[Nom d'utilisateur]",utilisateur)


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
  line=getLine(f.ReadLine)
  flag=split(line,separator)(GetColumnIndex("Flag"))
  if flag="0" then exit do
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
f.close()
f_repeated.close()
table=table+"</table>"


report_content=replace(report_content,"[table]",table)

save_file annee,mois,jour,article
if flag="0" then BlankFile(filename)

'for each col in columns
'	WScript.Echo col
'next 
'Do Until f.AtEndOfStream
'  WScript.Echo f.ReadLine
'Loop


function BlankFile(fi)
on error resume next
	Set objFile = fso.OpenTextFile(fi, 2)


	fso.writeline ""

	fso.Close

end function

function print(text)
   WScript.Echo text
End Function

function getLine(s)
   s=replace(s,"""","")
   s=left(s,len(s)-2)
   getLine=s
end function

function GetColumnIndex(col_name)
   i=0
   
   for each col in first_columns
      i=i+1
	  'if col_name="Temperature" then print("compare "+col+ " "+ col_name)
	  if col = col_name then 
		GetColumnIndex=i-1
		'if col_name="Temperature" then print("trouvé i=" & i-1)
		Exit function
	  end if
   next
   GetColumnIndex=0
End Function

function save_file(an,mois,day,art)
   tdir=target_dir+"\"+cstr(an)+"\"+right("00"+cstr(mois),2)+"\"+right("00"+cstr(day),2)
   If not fso.FolderExists(tdir) Then 
      'print(tdir)
      shell.Run "cmd /c mkdir "+tdir,0,true
   End if
   outFile=tdir+"\"+cstr(an)+"-"+right("00"+cstr(mois),2)+"-"+right("00"+cstr(day),2)+"-"+art+".html"
   Set objFile = fso.CreateTextFile(outFile,True)
   objFile.Write report_content & vbCrLf
   objFile.Close
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

end sub


sub ie_PrintTemplateTeardown(pDisp)
    bpttd_ready=true    'global bpttd_ready; no dim here
end sub
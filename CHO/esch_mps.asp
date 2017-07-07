<%response.charset="UTF-8"%>
<%response.Codepage = "65001"%>
<html>

<head>
<title>Production Schedule List</title>
</head>

<body bgcolor="#FFFFFF">
<p align="center"> <b><font face="Times New Roman" color="#008000"> <font size="6">Chongqing plant Production 
  Schedule List</font></font></b> </p>
<p></p> 
<table width="75%" border="1" align="center"> 
<tr> 
    <td width="13%" bgcolor="#008000"> 
      <p align="center"><b><font face="Times New Roman" color="#FFFFFF">Sun</font></b></p> 
    </td> 
    <td width="13%" bgcolor="#008000"> 
      <p align="center"><b><font face="Times New Roman" color="#FFFFFF">Mon</font></b></p> 
    </td> 
    <td width="13%" bgcolor="#008000"> 
      <p align="center"><b><font face="Times New Roman" color="#FFFFFF">Tue</font></b></p> 
    </td> 
    <td width="13%" bgcolor="#008000"> 
      <p align="center"><b><font face="Times New Roman" color="#FFFFFF">Wed</font></b></p> 
    </td> 
    <td width="13%" bgcolor="#008000"> 
      <p align="center"><b><font face="Times New Roman" color="#FFFFFF">Thu</font></b></p> 
    </td> 
    <td width="13%" bgcolor="#008000"> 
      <p align="center"><b><font face="Times New Roman" color="#FFFFFF">Fri</font></b></p> 
    </td> 
    <td width="13%" bgcolor="#008000"> 
      <p align="center"><b><font face="Times New Roman" color="#FFFFFF">Sat</font></b></p> 
    </td> 
     
  </tr> 
<tr> 
<% 
i=0 
j=0 
k=0 
dt=date() 
dt=dt-21 
a=weekday(dt,2) 
 
if a<7 then 
   for j=1 to a 
      response.write "<td width='13%'>&nbsp;</td>"  
      k=k+1  
   next  
end if 
 
 
do while i< 22 
     a=weekday(dt,2) 
     d=day(dt) 
     if d<10 then  
       days="0" & trim(d) 
     else days=trim(d) 
     end if 
 
     m=month(dt) 
     if m<10 then  
       months="0" & trim(m) 
     else months=trim(m) 
     end if 
 
     years=trim(year(dt)) 
 
     intranetLocation = ".\data\" 
     name="mps" & years & months & days & "l1.gif" 

     set objfilesys=server.createobject("scripting.filesystemobject") 

     if objfilesys.fileexists(server.mappath(intranetLocation & name )) then   
        response.write "<td width='13%' bgcolor='#CAFFCA'><p align='center'><b><font face='Times New Roman' color='#008000'><a href='eschPerDay.asp?v_date=" & dt & "'>" & dt & "</a></font></b></td>" 
 
     else  
 
         name="mps" & years & months & days & ".txt" 
         set objfilesys=server.createobject("scripting.filesystemobject") 
         if objfilesys.fileexists(server.mappath(intranetLocation & name )) then   
            response.write "<td width='13%' bgcolor='#CAFFCA'><p align='center'><b><font face='Times New Roman' color='#008000'><a href='eschPerDay.asp?v_date=" & dt & "'>" & dt & "</a></font></b></td>" 
 
         else  
 
             response.write "<td width='13%' bgcolor='#CAFFCA'><p align='center'><b><font face='Times New Roman' color='#008000'>"& dt &"</a></font></b></td>" 
             ' response.write <br> 
         end if   
          
     end if 
 
   
     k=k+1 
     if k=7 then  
        response.write "</tr><tr>" 
        k=0 
    end if 
  
    i=i+1 
    dt=dt+1 
loop   
 
if k<7 and k<>0 then  
 for j=1 to 7-k 
    response.write "<td width='13%'>&nbsp;</td>" 
  next 
end if 
     
%>  
  </tr> 
</table>
</body>  
</html>  

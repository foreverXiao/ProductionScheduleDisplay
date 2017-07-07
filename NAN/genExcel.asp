<%
if not isdate(request("v_date")) then
            response.end 
end if


dt = cdate(request("v_date"))

d = day(dt)
if d<10 then 
   days="0" & trim(d)
else
   days=trim(d)
end if

m=month(dt)
if m<10 then 
   months="0" & trim(m)
else
   months=trim(m)
end if

years=trim(year(dt)) 

set objfilesys = server.createobject("scripting.filesystemobject")
intranetLocation = request("loc")
continueOrNot = true



responseText = ""

if continueOrNot then
        name="mps" & years & months & days  & ".txt"


       if objfilesys.fileexists(server.mappath(intranetLocation & name)) then

            set objfile1= objfilesys.getfile(server.mappath(intranetLocation & name))
            set objfile1 = nothing

            Set adoStream = Server.CreateObject("ADODB.Stream")
     
            adoStream.Charset = "UTF-8" 
            adoStream.Open 
            adoStream.LoadFromFile server.mappath(intranetLocation & name) 'change this to point to your text file

            Dim mulArray 
            mulArray = split(adoStream.ReadText,"^")

            set adoStream = nothing


            curProdLine = 999 
            linesCount = 0 'counter for production line

            responseText = "<table border=1><tr><td>Line</td><td>lot no</td><td>item</td><td>quantity</td><td>start time</td><td>working minutes</td><td>screw</td><td>vip</td><td>remark</td></tr>"
            For intIndex = LBound(mulArray) To UBound(mulArray) 
                    dim mulItems
                    mulItems = split(mulArray(intIndex),"@")
                    daysToCover = cdate(mulItems(3)) - cdate(mulItems(2)) + 2 ' startTime1, endTime1
                    curProdLine = mulItems(0)
                    pixelsPerDay = mulItems(1)
                    startD = cdate(mulItems(2))
                    offSetFromStart = mulItems(4)
                    startTime = dateadd("n",offSetFromStart,startD)
                    workingMinutes = mulItems(5)
                    txt_lot_no = mulItems(6)
                    txt_item_no = mulItems(7)
                    planned_production_qty = mulItems(8)
                    planned_production_qty = left(planned_production_qty,len(planned_production_qty)-2)
                    txt_order_key = mulItems(9)
                    txt_VIP = mulItems(10)
                    txt_remark = mulItems(11)
                    leftPosForRSDspan = mulItems(12)
                    SPANandETD = mulItems(13)
                    marginTop = "" ''mulItems(14)
                    pullScrew = mulItems(15)
                    linesList = mulItems(16)
                    
                    if len(pullScrew) > 3 then  'color: #FFD700 ==> to  yellow
                       pullScrew = "yellow"
                    end if

                    'vbCrLF
                   responseText = responseText & "<tr><td>" & curProdLine & "</td><td style = 'mso-number-format:\@;'>" & txt_lot_no & "</td><td>" & txt_item_no & "</td><td>" & planned_production_qty & "</td><td>" & startTime & "</td><td>" & workingMinutes & "</td><td>" & pullScrew & "</td><td>" & txt_VIP & "</td><td>" & txt_remark & "</td></tr>"
                    
            Next 

            responseText = responseText & "</table>"
   

       end if

end if

set objfilesys = nothing

if len(responseText) = 0 then
        responseText = "No data"
end if 

'response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
response.ContentType = "application/vnd.ms-excel"
response.AddHeader "content-disposition", "attachment; filename=mps" &  years & months & days &".xls" 
response.Write(responseText)


%> 



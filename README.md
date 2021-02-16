- 👋 Hi, I’m @pankajmakwana-1319
- 👀 I’m interested in ...
- 🌱 I’m currently learning ...
- 💞️ I’m looking to collaborate on ...
- 📫 How to reach me ...

<!---
pankajmakwana-1319/pankajmakwana-1319 is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->
'*******************************************************************************************

' PURPOSE:

' This script is used to get Batch Size from Process Order Number and update custom property

'*******************************************************************************************

' Version:1.0.0    Date:29-Nov-2018   Author:Bansari Dave

' Comment: Created new script

' Reference ID: NA

' -----------------------------------------------------------------------------------------

 

Function Event_OnComplete(Data)

 

Dim objData, objXMlPars, objHTTPDOM, objExeResp, objXML, objResponseDOM

Dim strWib, strParExp, strBehaviorParamValue

Dim strEquipmentCode, strBehaviorResponse, strProcessOrderNumber

Dim strCustProp, strCustPropValues, strEventParamValues, strSQL, strTaskMsg

 

'****************Inputs**********************************************

 

strCustProp = "Usage Count"

 

'****************Input from Data XML**********************************************

 

Set objData = CreateObject("msxml2.DOMDocument.4.0")

objData.Async = False

objData.LoadXML Data

strEquipmentCode = objData.SelectSingleNode("//Equipment/@Code").Text

strWib = "PSSB_ET_Get_and_Set_EquipmentCustomProperty.wib"

strBehaviorParamValue = "<VariableValues><Variable name='CS_BEHAVIOREXECUTED' category='Recipe Info' description='Behavior Executed'/></VariableValues>"

 

 

'***************Behavior Input Parameters******************

 

Set objXMlPars = CreateObject("msxml2.DOMDocument.4.0")

objXMlPars.Async = False

strParExp = "<BehaviorValues><Object name='parEquipment'></Object><Object name='parAction'>Read</Object><Object name='parCustomProperties'>" + strCustProp + "</Object><Object name='parValues'></Object></BehaviorValues>"

objXMlPars.LoadXML strParExp

objXMlPars.SelectSingleNode("//Object[@name='parEquipment']").Text = strEquipmentCode

strParExp = objXMlPars.XML

Set objXMlPars = Nothing

 

'**************Call Behavior Execution Web Service****************

 

Set objHTTPDOM = CreateObject("DMIRABehaviorsRuntime.Execution")

 

'***********Create Response DOM***************

 

Set objExeResp = CreateObject("msxml2.DOMDocument.4.0")

objExeResp.Async = "False"

 

'***********Load Response********************

 

If Not objExeResp.LoadXML(objHTTPDOM.ExecuteBehavior(strWib, strParExp, strBehaviorParamValue)) Then

'Web service response can not be loaded into a parser. Need to raise an error

  Err.Raise 1000, "Event Script", "Error reading custom properties of an equipment. Raw response from web service: " & objHTTPDOM.responseText

End If

 

If objExeResp.SelectSingleNode("//CS_RESULT").Text = "CS_EVENTFAILED" Then

  Err.Raise 1000, "Event Script", "Error while extracting custom property values: " & objExeResp.SelectSingleNode("//CS_MSG").Text

  Set objHTTPDOM = Nothing

  Set objExeResp = Nothing

  Set objData = Nothing

  Set objXMlPars = Nothing

  Exit Function

End If

 

strBehaviorResponse = objExeResp.SelectSingleNode("//CS_PARAMETERVALUE").Text

 

 

Set objXML = CreateObject("msxml2.DOMDocument.4.0")

objXML.Async = False

 

objXML.LoadXML Data

 

strProcessOrderNumber = objXML.SelectSingleNode("//Parameters/Parameter[@Name='Process Order Number']/@Value").Text

strSQL = "SELECT TotalOrderQuantity FROM DMI_Manu32.dbo.OM_Orders WITH (NOLOCK) WHERE OrderNumber='" & strProcessOrderNumber & "' AND ActiveTF=1 AND DeletedTF=0"

strTaskMsg = "<?xml version='1.0' encoding='utf-8'?>" & "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" & "<soap:Body>" & "<GetXML xmlns='http://DMI.SA.GUI/ESigWS'> <SQL>" & strSQL & "</SQL> </GetXML>" & "</soap:Body>" & "</soap:Envelope>"

Set objXML = CreateObject("MSXML2.DOMDocument.6.0")

objXML.LoadXML strTaskMsg

 

Set objHTTPDOM = CreateObject("msxml2.ServerXMLHTTP.6.0")

objHTTPDOM.Open "POST", "http://localhost/ESig/ESig.asmx?wsdl", False

objHTTPDOM.setRequestHeader "SOAPAction", "http://DMI.SA.GUI/ESigWS/GetXML"

objHTTPDOM.setRequestHeader "Content-Type", "text/xml; charset=utf-8"

objHTTPDOM.Send objXML.XML

Set objResponseDOM = CreateObject("msxml2.DOMDocument.6.0")

objResponseDOM.LoadXML (objHTTPDOM.responseText)


 objResponseDOM.setProperty "SelectionNamespaces", "xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' xmlns:z='#RowsetSchema'"

If Not (objResponseDOM.SelectSingleNode("//soap:Body") Is Nothing) Then

  objResponseDOM.LoadXML  objResponseDOM.SelectSingleNode("//soap:Body").FirstChild.Text

Else

  Call Err.Raise(3, "Error", "Process Order Number Not Found in Database")

End If

 

If Not(objResponseDOM.selectSingleNode("//z:row/@TotalOrderQuantity") Is Nothing) Then

  strCustPropValues = objResponseDOM.selectSingleNode("//z:row/@TotalOrderQuantity").Text

 Else

  Call Err.Raise(3,"Error","Batch Size for particular Process Order Number Not Found in Database")

 End If

 

'****************Input from Data XML**********************************************

 

Set objData = CreateObject("msxml2.DOMDocument.4.0")

objData.Async = False

objData.LoadXML Data

strEventParamValues = objData.SelectSingleNode("//Parameter[@Name='Process Order Number']/@Value").Text

strWib = "PSSB_ET_Get_and_Set_EquipmentCustomProperty.wib"

strBehaviorParamValue = "<VariableValues><Variable name='CS_BEHAVIOREXECUTED' category='Recipe Info' description='Behavior Executed'/></VariableValues>"

 

'***************Behavior Input Parameters******************

 

Set objXMlPars = CreateObject("msxml2.DOMDocument.4.0")

objXMlPars.Async = False

strBehaviorResponse = CDbl(strBehaviorResponse) + CDbl(strCustPropValues)

strParExp = "<BehaviorValues><Object name='parEquipment'></Object><Object name='parAction'>Write</Object><Object name='parCustomProperties'>" + strCustProp + "</Object><Object name='parValues'>" + CStr(strBehaviorResponse) + "</Object></BehaviorValues>"

objXMlPars.LoadXML strParExp

objXMlPars.SelectSingleNode("//Object[@name='parEquipment']").Text = strEquipmentCode

strParExp = objXMlPars.XML

Set objXMlPars = Nothing

 

'**************Call Behavior Execution Web Service****************

 

Set objHTTPDOM = CreateObject("DMIRABehaviorsRuntime.Execution")

 

'***********Create Response DOM***************

 

Set objExeResp = CreateObject("msxml2.DOMDocument.4.0")

objExeResp.Async = "False"

 

'***********Load Response********************

 

If Not objExeResp.LoadXML(objHTTPDOM.ExecuteBehavior(strWib, strParExp, strBehaviorParamValue)) Then

'Web service response can not be loaded into a parser. Need to raise an error

  Err.Raise 1000, "Event Script", "Error writing custom properties of an equipment. Raw response from web service: " & objHTTPDOM.responseText

End If

 

End Function

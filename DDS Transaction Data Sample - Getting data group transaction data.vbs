'This example reads the most recent transactions from a device in the DDS, 
'parses the individual data elements and stores them into a jagged dictionary

'APIs used:
'   CxDds.DdsClient
'		- Connect
'		- GetRecentSuccessfulDataGroupTXs
'		- GetDataGroupTxDataWithRefs
'	CxScript.Dictionary


Option Explicit 

'DDS service to connect to - currently hard-coded
Dim strUisSiteService : strUisSiteService = "C4PROD.DDS"

'Start date of how far to look back for successful transactions
Dim dNewerThan : dNewerThan = Now() - 1

'Device, data group type and ordinal to look for
Dim strDeviceId : strDeviceId = "BUTCHER_RD"
Dim strDataGroupType : strDataGroupType = "AWTHIST"
Dim iDataGroupOrd : iDataGroupOrd = 1

'Create DDS Client object and connect to service
On Error Resume Next
Err.Clear
Dim ddsClient : Set ddsClient = CreateObject("CxDds.DdsClient")
ddsClient.Connect(strUisSiteService)

If Err.Number <> 0 Then
	'Error connecting to DDS
	WScript.Echo "Error connecting to " & strUisSiteService
	WScript.Echo Err.Description
	WScript.Quit	
End If

On Error Goto 0

'Get transaction data for supplied device id, data group type, and ordinal
Dim strError : strError = ""
Dim txDictionary : Set txDictionary = GetRecentTransactionData(strDeviceId, strDataGroupType, iDataGroupOrd, dNewerThan, strError)

If strError = "" Then
	'Parse the results
	Call OutputTransactionDictionary(txDictionary)
Else
	WScript.Echo "Error: " & strError
End If

'----------------------------------------
' Functions and Subs
'----------------------------------------	
	'Get the data from the most recent successful transactions
	'Inputs:
	'	strDeviceId - the ID of the device
	'	strDataGroupType - the data group type
	'	iDataGroupOrd - the ordinal or instance of the data group type specified
	'	dNewerThan - all successful transactions newer than this date will be returned
	'	strError - the function will populate this variable with any error occurred during the function
	'Output: a jagged CygNet dictionary
	'	key: the database key of the transaction - there really isn't any use for this other than to maintain uniqueness
	'	value: a CygNet dictionary
	'		key: data element id
	'		value: value of the data element
	Function GetRecentTransactionData(ByVal strDeviceId, ByVal strDataGroupType, ByVal iDataGroupOrd, ByVal dNewerThan, ByRef strError)		
		'clear error string
		strError = ""
		
		'Create a store container for the transaction data
		'	key = database key
		'   value = dictionary of element ID and value pairs
		Dim transactionDictionary : Set transactionDictionary = CreateObject("CxScript.Dictionary")
		
		'Get recent successful data group transactions
		On Error Resume Next
		Err.Clear
		Dim strTransXML : strTransXML = ddsClient.GetRecentSuccessfulDataGroupTXs(strDeviceId, strDataGroupType, iDataGroupOrd, dNewerThan)
		
		If Err.Number <> 0 Then
			'There was an error getting the transactions
			strError = "Error issuing command to get recent data group transactions (GetRecentSuccessfulDataGroupTXs): " & Err.Description
			Set GetRecentTransactionData = transactionDictionary
			Exit Function 
		End If
		On Error Goto 0
				
		'Create XML object and load transactions
		Dim objXmlDoc : Set objXmlDoc = CreateObject("Msxml2.DOMDocument")
		If Not objXmlDoc.loadXML(strTransXML) Then			
			strError = "Error - failed to load transactions XML"
			Set GetRecentTransactionData = transactionDictionary
			Exit Function
		End If
		
		'Parse the transaction list
		Dim rootNode : Set rootNode = objXmlDoc.selectSingleNode("dgTxList")
		
		Dim child
		For Each child In rootNode.childNodes	
			'Load the individual data group transaction information
			Dim childXmlDoc : Set childXmlDoc = CreateObject("Msxml2.DOMDocument")
			childXmlDoc.loadXML(child.xml)
			
			'Get the db Key and get the transaction data
			Dim dbKey : dbKey = childXmlDoc.documentElement.attributes.getNamedItem("dbKey").text
			Dim dbTimeStamp : dbTimeStamp = childXmlDoc.documentElement.attributes.getNamedItem("timestamp").text
			
			'Parse transaction results and store them to the transaction dictionary
			Dim strGetTransactionDataDictionaryError : strGetTransactionDataDictionaryError = ""
			Dim results : Set results = GetTransactionDataDictionary(dbKey, strGetTransactionDataDictionaryError)	
			
			'Check for error
			If strGetTransactionDataDictionaryError = "" Then
				'Post result
				transactionDictionary.SetKeyValue dbKey, results
			End If
		Next	
		
		Set GetRecentTransactionData = transactionDictionary
	End Function
	
	'Get the element name and values for the supplied transaction key
	'Inputs:
	'	transactionDbKey - the key of the transaction
	'	strError - the function will populate this variable with any error occurred during the function
	'Output: a CygNet dictionary
	'	key: the data group element id
	'	value: the value for the data group element id	
	Function GetTransactionDataDictionary(ByVal transactionDbKey, ByRef strError)
		Dim dictionary : Set dictionary = CreateObject("CxScript.Dictionary")
	
		'Get the transaction data
		Dim strTranXml : strTranXml = ddsClient.GetDataGroupTxDataWithRefs(transactionDbKey)
				
		'Load the transaction data and parse results
		Dim dbTransDataDoc : Set dbTransDataDoc = CreateObject("Msxml2.DOMDocument")
		dbTransDataDoc.loadXML(strTranXml)
		
		'Parse each element and put element name and value into a dictionary
		Dim dgElement
		For Each dgElement In dbTransDataDoc.documentElement.childNodes
			Dim dgElementName : dgElementName = dgElement.nodeName
			Dim dgElementValue : dgElementValue = dgElement.text
	
			dictionary.SetKeyValue dgElementName, dgElementValue
		Next
	
		'Return the results
		Set GetTransactionDataDictionary = dictionary
	End Function
	
	
	'Output the transaction data dictionary
	'Inputs:  a jagged CygNet dictionary (from GetRecentTransactionData)
	'	key: the database key of the transaction - there really isn't any use for this other than to maintain uniqueness
	'	value: a CygNet dictionary
	'		key: data element id
	'		value: value of the data element	
	'Output: echos to screen
	Sub OutputTransactionDictionary(ByVal transactionDictionary)
		'Parse results
		Dim arrKeyList : arrKeyList = transactionDictionary.GetKeyList
		
		'Read transaction dictionary
		Dim dbKey
		For Each dbKey In arrKeyList
			'For each transaction, process the data elements and values
			WScript.Echo "--------------------------"
			WScript.Echo dbKey
			WScript.Echo "--------------------------"
			Dim myTx : Set myTx = transactionDictionary.Value(dbKey)
			Dim arrMyTxKeys : arrMyTxKeys = myTx.GetKeyList
			
			Dim kvp
			For Each kvp In myTx
				WScript.Echo kvp.key & " = " & kvp.value
			Next
		Next
		
		WScript.Echo "--------------------------"
		WScript.Echo "End"		
	End Sub		
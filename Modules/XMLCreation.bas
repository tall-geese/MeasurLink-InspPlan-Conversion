Attribute VB_Name = "XMLCreation"
'Testing the ability to Create XML files that MeasurLink can read for us.
Public Const XML_SCHEMA_ATTR = "xmlns:xsd"
Const XML_SCHEMA_VALUE = "http://www.w3.org/2001/XMLSchema"
Const XML_SCHEMA_INST_ATTR = "xmlns:xsi"
Const XML_SCHEMA_INST_VALUE = "http://www.w3.org/2001/XMLSchema-instance"
Const XML_SCHEMA_LOCATION_ATTR = "xsi:schemaLocation"
Const XML_SCHEMA_LOCATION_VALUE = "http://qifstandards.org/xsd/qif3 ../QIFApplications/QIFDocument.xsd"
Const XML_SCHEMA_VERSION_ATTR = "versionQIF"
Const XML_SCHEMA_VERSION_VALUE = "3.0.0"
Const XML_SCHEMA_NAMESPACE_VALUE = "http://qifstandards.org/xsd/qif3"

Sub CreateXMLTest()

    Dim doc As MSXML2.DOMDocument60
    Dim styles As MSXML2.DOMDocument60
    Dim output As MSXML2.DOMDocument60
    Set doc = New MSXML2.DOMDocument60
    Dim xmlVersion As MSXML2.IXMLDOMProcessingInstruction
    Dim xmlSchema As MSXML2.IXMLDOMAttribute, xmlInstance As MSXML2.IXMLDOMAttribute, xmlLocation As MSXML2.IXMLDOMAttribute, QIFVersion As MSXML2.IXMLDOMAttribute
    Dim root As MSXML2.IXMLDOMNode
    Dim product As MSXML2.IXMLDOMElement
    Dim characteristics As MSXML2.IXMLDOMElement
    Dim fso As FileSystemObject

    'idMax Value???? is this really necessary??
    
    'Setup Test Example array
    Dim testArr(0 To 4, 1 To 3) As Variant
    testArr(0, 1) = "0_008_00"
    testArr(1, 1) = "Variable"
    testArr(2, 1) = "0.139"
    testArr(3, 1) = "0.141"
    testArr(4, 1) = "0.143"
    
    testArr(0, 2) = "0_021_00"
    testArr(1, 2) = "Variable"
    testArr(2, 2) = "0.036"
    testArr(3, 2) = "0.038"
    testArr(4, 2) = "0.040"
    
    testArr(0, 3) = "0_023_00"
    testArr(1, 3) = "Attribute"
    testArr(2, 3) = ""
    testArr(3, 3) = ""
    testArr(4, 3) = ""
    
'    Set xmlVersion = doc.createNode(NODE_PROCESSING_INSTRUCTION, "xml", XML_SCHEMA_NAMESPACE_VALUE)
'    xmlVersion1 = doc.createAttribute("version")
'    xmlVersion1.Value = Chr(34) & "1.0" & Chr(34)
'    xmlVersion.Attributes.setNamedItem xmlVersion1
'    Set xmlVersion = doc.createProcessingInstruction("xml", "version=" & Chr(34) & "1.0" & Chr(34))
'    doc.appendChild xmlVersion
'    Set Encoding = doc.createProcessingInstruction("xml", "encoding=" & Chr(34) & "UTF-8" & Chr(34))
'    doc.appendChild Encoding
    
    'Root Node, QIF
    Set root = doc.createNode(NODE_ELEMENT, "QIFDocument", XML_SCHEMA_NAMESPACE_VALUE)
    doc.appendChild root
    Set xmlSchema = doc.createAttribute(XML_SCHEMA_ATTR)
    xmlSchema.Value = XML_SCHEMA_VALUE
    Set xmlInstance = doc.createAttribute(XML_SCHEMA_INST_ATTR)
    xmlInstance.Value = XML_SCHEMA_INST_VALUE
    Set xmlLocation = doc.createAttribute(XML_SCHEMA_LOCATION_ATTR)
    xmlLocation.Value = XML_SCHEMA_LOCATION_VALUE
    Set QIFVersion = doc.createAttribute(XML_SCHEMA_VERSION_ATTR)
    QIFVersion.Value = XML_SCHEMA_VERSION_VALUE
    
    root.Attributes.setNamedItem xmlSchema
    root.Attributes.setNamedItem xmlInstance
    root.Attributes.setNamedItem xmlLocation
    root.Attributes.setNamedItem QIFVersion
    
    'Level-1, Product and Characteristics
    Set product = doc.createNode(MSXML2.NODE_ELEMENT, "Product", XML_SCHEMA_NAMESPACE_VALUE)
    Set characteristics = doc.createNode(MSXML2.NODE_ELEMENT, "Characteristics", XML_SCHEMA_NAMESPACE_VALUE)
    
    root.appendChild newChild:=product
    root.appendChild newChild:=characteristics
    
    'Product Tree
        'Level-2 (Product) -> Header, PartSet
    Set headerNode = doc.createNode(MSXML2.NODE_ELEMENT, "Header", XML_SCHEMA_NAMESPACE_VALUE)
    product.appendChild newChild:=headerNode
    Set partSet = doc.createNode(MSXML2.NODE_ELEMENT, "PartSet", XML_SCHEMA_NAMESPACE_VALUE)
    partSet.setAttribute "n", "1"
    product.appendChild newChild:=partSet
    
        'Level-3 (Header,PartSet) -> Name, Part
    Set nameNode = doc.createNode(MSXML2.NODE_ELEMENT, "Name", XML_SCHEMA_NAMESPACE_VALUE)
    nameNode.Text = "DRW-00717-02_RAJ_FA_SYLVAC_TEST8"
    Set partNode = doc.createNode(MSXML2.NODE_ELEMENT, "Part", XML_SCHEMA_NAMESPACE_VALUE)
    partNode.setAttribute "id", "0"   '<--unique ID
    headerNode.appendChild nameNode
    partSet.appendChild partNode
    
        'Level-4 (Part) -> Header, CharacteristicNominalIds
    Set partHeaderNode = doc.createNode(MSXML2.NODE_ELEMENT, "Header", XML_SCHEMA_NAMESPACE_VALUE)
    Set partCharsNode = doc.createNode(MSXML2.NODE_ELEMENT, "CharacteristicNominalIds", XML_SCHEMA_NAMESPACE_VALUE)
    partCharsNode.setAttribute "n", "3" '<-- number of features for this routine
    partNode.appendChild partHeaderNode
    partNode.appendChild partCharsNode
    
        'Level-5 (Header, CharacteristicNominalIDs) -> Name, IDs
    Set partHeaderNameNode = doc.createNode(MSXML2.NODE_ELEMENT, "Name", XML_SCHEMA_NAMESPACE_VALUE)
    partHeaderNameNode.Text = "DRW-00717-10_RAJ"
    partHeaderNode.appendChild partHeaderNameNode
    
    'TODO: This should probably be done when we are creating the Characteristic noms so we dont have to solve for or remember the IDs
'    For i = 0 To 2 '<-- need to iterate throught the IDs of our features, leaving this placeholder loop here
'        Set partCharsIDNode = doc.createNode(MSXML2.NODE_ELEMENT, "ID", XML_SCHEMA_NAMESPACE_VALUE)
'        partCharsIDNode.Text = i
'        partCharsNode.appendChild partCharsIDNode
'    Next i
    
    'Characteristics Tree
        'Level 1 (Characterisitcs) -> CharacteristicDefinitions, CharacterisitcNominals, CharacteristicItems
    Set charDefsNode = doc.createNode(MSXML2.NODE_ELEMENT, "CharacteristicDefinitions", XML_SCHEMA_NAMESPACE_VALUE)
    Set charNomsNode = doc.createNode(MSXML2.NODE_ELEMENT, "CharacteristicNominals", XML_SCHEMA_NAMESPACE_VALUE)
    Set charItemsNode = doc.createNode(MSXML2.NODE_ELEMENT, "CharacteristicItems", XML_SCHEMA_NAMESPACE_VALUE)
    charDefsNode.setAttribute "n", "3" '<-- number of features
    charNomsNode.setAttribute "n", "3" '<-- number of features
    charItemsNode.setAttribute "n", "3" '<-- number of features
    characteristics.appendChild charDefsNode
    characteristics.appendChild charNomsNode
    characteristics.appendChild charItemsNode
    
        
    For i = 1 To 3 '<--- iterate through our list of features
        'CharacteristicDefinitions
        Dim style As String
        If testArr(1, i) = "Variable" Then
            style = "Linear"
        Else
            style = "Attribute"
        End If
        
        
'        style = "Attribute"  '<--- grab the offset cell or wherever we will define the "Linear/Variable" vs "Attribute" for a dim
        Set charDef = doc.createNode(MSXML2.NODE_ELEMENT, "UserDefined" & style & "CharacteristicDefinition", XML_SCHEMA_NAMESPACE_VALUE)
        charDef.setAttribute "id", i '<---Unique ID
        charDefsNode.appendChild charDef
        
        Set charDefName = doc.createNode(MSXML2.NODE_ELEMENT, "Name", XML_SCHEMA_NAMESPACE_VALUE)
        charDefName.Text = testArr(0, i) '<--feature name
        Set charDefMeas = doc.createNode(MSXML2.NODE_ELEMENT, "WhatToMeasure", XML_SCHEMA_NAMESPACE_VALUE)
        charDefMeas.Text = testArr(0, i) '<--feature name
        charDef.appendChild charDefName
        charDef.appendChild charDefMeas
        
        If style = "Linear" Then
            Set charDefTol = doc.createNode(MSXML2.NODE_ELEMENT, "Tolerance", XML_SCHEMA_NAMESPACE_VALUE)
            
            Set charDefMax = doc.createNode(MSXML2.NODE_ELEMENT, "MaxValue", XML_SCHEMA_NAMESPACE_VALUE)
            charDefMax.Text = testArr(4, i) '<-- Upper Limit
            Set charDefMin = doc.createNode(MSXML2.NODE_ELEMENT, "MinValue", XML_SCHEMA_NAMESPACE_VALUE)
            charDefMin.Text = testArr(2, i) '<-- Lower Limit
            Set charDefLimit = doc.createNode(MSXML2.NODE_ELEMENT, "DefinedAsLimit", XML_SCHEMA_NAMESPACE_VALUE)
            charDefLimit.Text = "true"
            
            charDefTol.appendChild charDefMax
            charDefTol.appendChild charDefMin
            charDefTol.appendChild charDefLimit
            
            charDef.appendChild charDefTol
        End If
        
        'CharacteristicNominals
        Set charNom = doc.createNode(MSXML2.NODE_ELEMENT, "UserDefined" & style & "CharacteristicNominal", XML_SCHEMA_NAMESPACE_VALUE)
        charNom.setAttribute "id", i + 3  '<---Unique ID (TODO: should be the number of features we have)
        charNomsNode.appendChild charNom
        
        Set charDefID = doc.createNode(MSXML2.NODE_ELEMENT, "CharacteristicDefinitionId", XML_SCHEMA_NAMESPACE_VALUE)
        charDefID.Text = i '<--Link Definition ID
        charNom.appendChild charDefID
        
            'Part Header information needs to reference the Nominal Ids for the Part Library Creation
        Set partCharsIDNode = doc.createNode(MSXML2.NODE_ELEMENT, "Id", XML_SCHEMA_NAMESPACE_VALUE)
        partCharsIDNode.Text = i + 3
        partCharsNode.appendChild partCharsIDNode
        
        If style = "Linear" Then
            Set charNomTarget = doc.createNode(MSXML2.NODE_ELEMENT, "TargetValue", XML_SCHEMA_NAMESPACE_VALUE)
            charNomTarget.Text = testArr(2, i) '<-- Nominal Value
            charNom.appendChild charNomTarget
        Else
                'Pass
            Set charNomPass = doc.createNode(MSXML2.NODE_ELEMENT, "PassValues", XML_SCHEMA_NAMESPACE_VALUE)
            charNomPass.setAttribute "n", "1"
            charNom.appendChild charNomPass
            Set passString = doc.createNode(MSXML2.NODE_ELEMENT, "StringValue", XML_SCHEMA_NAMESPACE_VALUE)
            passString.Text = "Pass"
            charNomPass.appendChild passString
            
                'Fail
            Set charNomFail = doc.createNode(MSXML2.NODE_ELEMENT, "FailValues", XML_SCHEMA_NAMESPACE_VALUE)
            charNomFail.setAttribute "n", "1"
            charNom.appendChild charNomFail
            Set failString = doc.createNode(MSXML2.NODE_ELEMENT, "StringValue", XML_SCHEMA_NAMESPACE_VALUE)
            failString.Text = "Fail"
            charNomFail.appendChild failString
        End If
        
        'CharacteristicItems
        Set charItem = doc.createNode(MSXML2.NODE_ELEMENT, "UserDefined" & style & "CharacteristicItem", XML_SCHEMA_NAMESPACE_VALUE)
        charItem.setAttribute "id", i + 6 '<---Unique ID (TODO: should be TWICE the number of features we have)
        charItemsNode.appendChild charItem
        
        Set charNomID = doc.createNode(MSXML2.NODE_ELEMENT, "CharacteristicNominalId", XML_SCHEMA_NAMESPACE_VALUE)
        charNomID.Text = i + 3 '<--Link Nominal ID (TODO: should be equal to the number of features we have)
        charItem.appendChild charNomID
        
        Set charItemDesc = doc.createNode(MSXML2.NODE_ELEMENT, "Description", XML_SCHEMA_NAMESPACE_VALUE)
        charItemDesc.Text = "DRW-00717-10_RAJ." & testArr(0, i) '<-- Format of "Part_Rev.Feature", like "1642652_D.0_026_01"
        charItem.appendChild charItemDesc
        
        Set charItemName = doc.createNode(MSXML2.NODE_ELEMENT, "Name", XML_SCHEMA_NAMESPACE_VALUE)
        charItemName.Text = testArr(0, i) '<-- FeatureName
        charItem.appendChild charItemName
    Next i

    'XSLT
    Set fso = New FileSystemObject
    Set styles = New MSXML2.DOMDocument60
    Set output = New MSXML2.DOMDocument60
    styles.async = False
    Dim stylesText As String
    stylesText = fso.GetFile("C:\Users\mdieckman\Desktop\styles.xml").OpenAsTextStream.ReadAll

'    styles.LoadXML bstrXML:=stylesText
'    doc.transformNodeToObject styles, output
    
    Set ts = fso.CreateTextFile("C:\Users\mdieckman\Desktop\New_Test.QIF")
    Dim writer As MSXML2.MXXMLWriter60
    Set writer = New MSXML2.MXXMLWriter60
    Dim reader As MSXML2.SAXXMLReader60
    Set reader = New MSXML2.SAXXMLReader60
    writer.indent = True
    writer.Encoding = "utf-8"
    writer.omitXMLDeclaration = False
    Set reader.contentHandler = writer
    reader.Parse doc
    ts.Write (writer.output)
    ts.Close
    
'    writer.output = doc
    

    'Output
'    output.Save ("C:\Users\mdieckman\Desktop\XMLTest.QIF")

End Sub


Sub testArr()
    Dim testArr(0 To 4, 1 To 3) As Variant
    
    testArr(0, 1) = "0_008_00"
    testArr(1, 1) = "Variable"
    testArr(2, 1) = ".139"
    testArr(3, 1) = ".141"
    testArr(4, 1) = ".143"
    
    testArr(0, 2) = "0_021_00"
    testArr(1, 2) = "Variable"
    testArr(2, 2) = ".036"
    testArr(3, 2) = ".038"
    testArr(4, 2) = ".040"
    
    testArr(0, 3) = "0_008_00"
    testArr(1, 3) = "Attribute"
    testArr(2, 3) = ""
    testArr(3, 3) = ""
    testArr(4, 3) = ""

    Debug.Print ("check this out")

End Sub

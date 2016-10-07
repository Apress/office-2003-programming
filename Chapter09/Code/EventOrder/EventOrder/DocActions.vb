Option Explicit On 
Imports ST = Microsoft.Office.Interop.SmartTag
Imports W = Microsoft.Office.Interop.Word

Public Class DocActions
	Implements Microsoft.Office.Interop.SmartTag.ISmartDocument
	'TODO: Change to 'Order' as opposed to 'Simple'
	Const cNAMESPACE As String = "urn:schemas-bravocorp-com.namespaces.event.simple"
	Const cXNS As String = "xmlns:ns='urn:schemas-bravocorp-com.namespaces.event.simple'"

	'Number of Elements
	Public Const cTYPES As Integer = 4

	'Main Elements
	Const cORDER As String = cNAMESPACE & "#Order"
	Const cSHOW As String = cNAMESPACE & "#Show"
	Const cCONTACT As String = cNAMESPACE & "#Customer"
  Const cITEMS As String = cNAMESPACE & "#Item"

	'TODO: Make next string dynamic
  Dim ConnectString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\clients\RealOfficeVB\Chapters\Chapter09\data.mdb"
  'Dim ConnectString As String = "Server=coreserv;Database=BravoCorp;User ID=sa;Password=cr3d3ra;Trusted_Connection=False"

	'State Variables - for rememboring user selections
	Dim mlSelectedShow As Long
	Dim mlSelectedCompany As Long
	Dim mlSelectedContact As Long
	Dim mlSelectedBooth As Long

  Dim oDoc As W.Document

#Region "SmartDocControlSetup"

	Public Sub SmartDocInitialize(ByVal ApplicationName As String, ByVal Document As Object, _
	 ByVal SolutionPath As String, ByVal SolutionRegKeyRoot As String) _
	 Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.SmartDocInitialize

		'Initialize Selection States
		mlSelectedShow = -1
		mlSelectedCompany = -1
		mlSelectedContact = -1
		mlSelectedBooth = -1

	End Sub

	Public ReadOnly Property SmartDocXmlTypeCount() As Integer _
		Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.SmartDocXmlTypeCount
		Get
			Return cTYPES
		End Get
	End Property

	Public ReadOnly Property SmartDocXmlTypeCaption(ByVal XMLTypeID _
	 As Integer, ByVal LocaleID As Integer) As String Implements _
	 Microsoft.Office.Interop.SmartTag.ISmartDocument.SmartDocXmlTypeCaption

		Get
			Select Case XMLTypeID
				Case 1

          Return "Bravo Corp: event-site Order"
				Case 2

					Return "Show Information"
				Case 3

          Return "Customer Information"
				Case 4

					Return "Order Information"
				Case Else

			End Select
		End Get
	End Property

	Public ReadOnly Property SmartDocXmlTypeName(ByVal XMLTypeID As Integer) As String _
		Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.SmartDocXmlTypeName
		Get
			Select Case XMLTypeID
				Case 1
					Return cORDER
				Case 2
					Return cSHOW

				Case 3
					Return cCONTACT
				Case 4
					Return cITEMS

				Case Else

			End Select
		End Get
	End Property

	Public ReadOnly Property ControlCount(ByVal XMLTypeName As String) _
		As Integer Implements _
		Microsoft.Office.Interop.SmartTag.ISmartDocument.ControlCount

		Get
			Select Case XMLTypeName
				Case cORDER
					ControlCount = 1
				Case cSHOW
					ControlCount = 3
				Case cCONTACT
					ControlCount = 6
				Case cITEMS
					ControlCount = 6
				Case Else
			End Select
		End Get
	End Property

	Public ReadOnly Property ControlID(ByVal XMLTypeName As String, _
		ByVal ControlIndex As Integer) As Integer _
		Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.ControlID
		Get
			Select Case XMLTypeName
				Case cORDER
					ControlID = ControlIndex
				Case cSHOW
					ControlID = ControlIndex + 100
				Case cCONTACT
					ControlID = ControlIndex + 200
				Case cITEMS
					ControlID = ControlIndex + 300
			End Select
		End Get
	End Property

	Public ReadOnly Property ControlNameFromID(ByVal ControlID As Integer) _
		As String Implements _
		Microsoft.Office.Interop.SmartTag.ISmartDocument.ControlNameFromID

		Get

			Return ControlID.ToString
		End Get
	End Property

	Public ReadOnly Property ControlCaptionFromID(ByVal ControlID As Integer, _
		ByVal ApplicationName As String, ByVal LocaleID As Integer, _
		ByVal Text As String, ByVal Xml As String, ByVal Target As Object) _
		As String Implements _
		Microsoft.Office.Interop.SmartTag.ISmartDocument.ControlCaptionFromID

		Get
			Select Case ControlID
				Case 1
					Return "To create a  new Event-site order, simply move through the document and make your selections."
				Case 101				' 
					Return "Select Desired Event/Show"
				Case 102				'Location TextBox
					Return "Location:"
				Case 103				'Insert Button
					Return "Insert Show Information"

				Case 201
					Return "Select Exhibitor Company Name:"
				Case 202
					Return "Select Company Contact Placing Order:"
				Case 203
					Return "Phone Number:"
				Case 204
					Return "Email:"
				Case 205
					Return "Select Booth Number:"
				Case 206
					Return "Insert Client Information"

				Case 301
					Return "Select Order Item:"
				Case 302
					Return "Part Number:"
				Case 303
					Return "Quantity:"
				Case 304
					Return "Price:"
				Case 305
					Return "Insert Line Item"
				Case 306
					Return "Submit Order"

				Case Else
			End Select
		End Get
	End Property

	Public ReadOnly Property ControlTypeFromID(ByVal ControlID As Integer, _
	 ByVal ApplicationName As String, ByVal LocaleID As Integer) _
	 As Microsoft.Office.Interop.SmartTag.C_TYPE Implements _
		Microsoft.Office.Interop.SmartTag.ISmartDocument.ControlTypeFromID

		Get
			Select Case ControlID
				Case 1
					Return ST.C_TYPE.C_TYPE_LABEL
				Case 101				'select show
					Return ST.C_TYPE.C_TYPE_COMBO
				Case 102				'location
					Return ST.C_TYPE.C_TYPE_TEXTBOX
				Case 103				'insert button
					Return ST.C_TYPE.C_TYPE_BUTTON

				Case 201				'company
					Return ST.C_TYPE.C_TYPE_COMBO
				Case 202				'customer contact
					Return ST.C_TYPE.C_TYPE_COMBO
				Case 203				'phone#
					Return ST.C_TYPE.C_TYPE_TEXTBOX
				Case 204				'email
					Return ST.C_TYPE.C_TYPE_TEXTBOX
				Case 205				'select booth
					Return ST.C_TYPE.C_TYPE_COMBO
				Case 206				'insert button
					Return ST.C_TYPE.C_TYPE_BUTTON

				Case 301				'select Part desc
					Return ST.C_TYPE.C_TYPE_COMBO
				Case 302				'part#
					Return ST.C_TYPE.C_TYPE_TEXTBOX
				Case 303				'qty
					Return ST.C_TYPE.C_TYPE_TEXTBOX
				Case 304				'price
					Return ST.C_TYPE.C_TYPE_TEXTBOX
				Case 305				'insert button
					Return ST.C_TYPE.C_TYPE_BUTTON
				Case 306				'submit button
					Return ST.C_TYPE.C_TYPE_BUTTON

				Case Else
			End Select
		End Get
	End Property

#End Region


#Region "Populate and Respond"

	Public Sub PopulateListOrComboContent(ByVal ControlID As Integer, _
	 ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal _
	 Text As String, ByVal Xml As String, ByVal Target As Object, ByVal _
	 Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, ByRef _
	 List As System.Array, ByRef Count As Integer, ByRef InitialSelected As Integer) _
	 Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateListOrComboContent

		Select Case ControlID
			Case 101			'shows
				FillListOrComboControlWithDBData("Select * from qryShows", _
				ConnectString, Count, List)

				InitialSelected = mlSelectedShow
				Props.Write("w", 200)
			Case 201			'company
				FillListOrComboControlWithDBData("Select * from qryCompanies", _
				ConnectString, Count, List)
				InitialSelected = mlSelectedCompany
				Props.Write("w", 200)
			Case 202			'contact
				FillListOrComboControlWithDBData("Select * from qryContacts", _
				ConnectString, Count, List)

				InitialSelected = mlSelectedContact
				Props.Write("w", 200)
			Case 205			'booth
				FillListOrComboControlWithDBData("Select * from qryBooths", _
				ConnectString, Count, List)

				InitialSelected = mlSelectedBooth
				Props.Write("w", 200)
			Case 301			'part#
				FillListOrComboControlWithDBData("Select * from qryProducts", _
				ConnectString, Count, List)

				InitialSelected = -1
				Props.Write("w", 200)
			Case Else

		End Select

	End Sub

  Public Sub PopulateTextboxContent(ByVal ControlID As Integer, _
  ByVal ApplicationName As String, ByVal LocaleID As Integer, _
  ByVal Text As String, ByVal Xml As String, ByVal Target As Object, _
  ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, _
  ByRef Value As String) _
  Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateTextboxContent
    Select Case ControlID
      Case 102   'location
        Props.Write("w", 200)
      Case 203   'phone#
        Props.Write("w", 200)
      Case 204   'email
        Props.Write("w", 200)
      Case 302   'part desc
        Props.Write("w", 200)
      Case 303   'qty
        Props.Write("w", 50)
      Case 304   'price
        Props.Write("w", 50)
      Case Else

    End Select
  End Sub


#End Region


#Region "RespondToUsers"

  Public Sub InvokeControl(ByVal ControlID As Integer, ByVal ApplicationName _
  As String, ByVal Target As Object, ByVal Text As String, ByVal Xml As _
  String, ByVal LocaleID As Integer) _
  Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.InvokeControl

    Dim xNode As W.XMLNode
    Dim rng As W.Range
    rng = Target
    xNode = rng.XMLNodes(1)

    Try
      Select Case ControlID
        Case 103
          xNode.SelectSingleNode("//ns:Show/ns:Name", cXNS).Range.Text = _
           xNode.SmartTag.SmartTagActions("101").TextboxText
          xNode.SelectSingleNode("//ns:Show/ns:Location", cXNS).Range.Text = _
           xNode.SmartTag.SmartTagActions("102").TextboxText

        Case 206
          xNode.SelectSingleNode("//ns:Customer/ns:CompanyName", cXNS).Range.Text = _
           xNode.SmartTag.SmartTagActions("201").TextboxText
          xNode.SelectSingleNode("//ns:Customer/ns:Name", cXNS).Range.Text = _
           xNode.SmartTag.SmartTagActions("202").TextboxText
          xNode.SelectSingleNode("//ns:Customer/ns:Phone", cXNS).Range.Text = _
          xNode.SmartTag.SmartTagActions("203").TextboxText
          xNode.SelectSingleNode("//ns:Customer/ns:Email", cXNS).Range.Text = _
           xNode.SmartTag.SmartTagActions("204").TextboxText
          xNode.SelectSingleNode("//ns:Customer/ns:Booth", cXNS).Range.Text = _
           xNode.SmartTag.SmartTagActions("205").TextboxText()

          'End If
          'rs.Close()
        Case 305
          If xNode.SmartTag.SmartTagActions("301").ListSelection = -1 Then
            MsgBox("Please select a Product item before attempting an Insert.", _
             MsgBoxStyle.Exclamation, "Pick Something, Anything")
            Exit Sub
          Else
            Dim strSku As String = xNode.SmartTag.SmartTagActions("302").TextboxText
            Dim strDesc As String = xNode.SmartTag.SmartTagActions("301").TextboxText
            Dim strQty As String = xNode.SmartTag.SmartTagActions("303").TextboxText
            Dim strPrice As String = xNode.SmartTag.SmartTagActions("304").TextboxText
            With rng.Application.Selection
              .Tables(1).Rows(.Tables(1).Rows.Count - 2).Select()

              If .Tables(1).Rows(.Tables(1).Rows.Count - 2).Range.Fields(1).Result.Text <> "$   0.00" Then
                .InsertRowsBelow(1)
                .Cells(5).Formula("=Product(Left)", "$#,##0.00;($#,##0.00)")
              End If
              xNode = .Rows.Last.Range.XMLNodes(1)

            End With
            With xNode.ChildNodes
              .Item(1).Range.Text = strSku
              .Item(2).Range.Text = strDesc
              .Item(3).Range.Text = strQty
              .Item(4).Range.Text = strPrice
            End With
          End If
          rng.Application.Selection.Tables(1).Range.Fields.Update()

        Case 306

          If oDoc.CustomDocumentProperties("OrderID").value = "0" Then
            'Invoke the Submit Order Web Service
            'Send the Target docs XML data file
            Dim boPO As New BravoOrders.ProcessOrders
            Dim strOrderID As String
            Dim strXML As String = oDoc.XMLNodes(1).XML(True)

            strXML = Replace(strXML, "xmlns", "xmlns:ns")

            strOrderID = boPO.SubmitNewOrder(oDoc.XMLNodes(1).XML(True))

            'TODO: Check for existence first
            With oDoc
              .CustomDocumentProperties("OrderID") = strOrderID
              With .Application
                .ActiveWindow.ActivePane.View.SeekView = _
                 W.WdSeekView.wdSeekCurrentPageHeader
                .Selection.HeaderFooter.Range.Text = "Submitted Order# " & strOrderID
                .Selection.Font.Name = "Tunga"
                .Selection.Font.Bold = True
                .Selection.Font.Size = 20
                .ActiveWindow.ActivePane.View.SeekView = W.WdSeekView.wdSeekMainDocument
              End With
            End With
            MsgBox("This order has been submitted. The Order Number is " & _
             strOrderID & ".", MsgBoxStyle.Information, "Order Submitted")
          Else
            MsgBox("This order has already been submitted.", MsgBoxStyle.Information, _
             "Submitted Order")
          End If

        Case Else

      End Select
    Catch ex As Exception
      MsgBox(Err.Description)
    End Try

  End Sub

  Public Sub OnListOrComboSelectChange(ByVal ControlID As Integer, _
   ByVal Target As Object, ByVal Selected As Integer, ByVal Value As String) _
   Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.OnListOrComboSelectChange

    Dim xNode As W.XMLNode
    xNode = Target.XMLNodes(1)
    Try
      Select Case ControlID
        Case 101    'shows
          mlSelectedShow = Selected
          Dim rs As New ADODB.Recordset
          rs.Open("Select * from tblShows Where Name = '" & Value & "'", _
           ConnectString, ADODB.CursorTypeEnum.adOpenStatic)

          xNode.SmartTag.SmartTagActions("102").TextboxText = rs.Fields("Location").Value

          rs.Close()
          rs = Nothing
        Case 201    'company
          mlSelectedCompany = Selected
        Case 202    'contact
          mlSelectedContact = Selected
          Dim rs As New ADODB.Recordset
          rs.Open("Select * From tblCustomers Where ContactName='" & Value & "'", _
           ConnectString, ADODB.CursorTypeEnum.adOpenStatic)

          xNode.SmartTag.SmartTagActions("203").TextboxText = rs.Fields("Phone").Value
          xNode.SmartTag.SmartTagActions("204").TextboxText = rs.Fields("Email").Value

          rs.Close()
          rs = Nothing

        Case 205    'booth
          mlSelectedBooth = Selected
        Case 301    'part#
          Dim rs As New ADODB.Recordset
          rs.Open("Select * From tblProducts Where ProductName='" & Value & "'", _
           ConnectString, ADODB.CursorTypeEnum.adOpenStatic)



          xNode.SmartTag.SmartTagActions("302").TextboxText = rs.Fields("ProductID").Value
          'coasdfmment


          xNode.SmartTag.SmartTagActions("303").TextboxText = "1"
          'commetn
          xNode.SmartTag.SmartTagActions("304").TextboxText = _
           Format(rs.Fields("UnitPrice").Value, "C")

          rs.Close()
          rs = Nothing
        Case Else

      End Select
    Catch ex As Exception

    End Try

  End Sub

#End Region


#Region "Empties"

	Public Sub ImageClick(ByVal ControlID As Integer, ByVal ApplicationName As String, ByVal Target As Object, ByVal Text As String, ByVal Xml As String, ByVal LocaleID As Integer, ByVal XCoordinate As Integer, ByVal YCoordinate As Integer) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.ImageClick

	End Sub


	Public Sub OnCheckboxChange(ByVal ControlID As Integer, ByVal Target As Object, ByVal Checked As Boolean) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.OnCheckboxChange

	End Sub



	Public Sub OnPaneUpdateComplete(ByVal Document As Object) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.OnPaneUpdateComplete

	End Sub

	Public Sub OnRadioGroupSelectChange(ByVal ControlID As Integer, ByVal Target As Object, ByVal Selected As Integer, ByVal Value As String) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.OnRadioGroupSelectChange

	End Sub

	Public Sub OnTextboxContentChange(ByVal ControlID As Integer, ByVal Target As Object, ByVal Value As String) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.OnTextboxContentChange

	End Sub

	Public Sub PopulateActiveXProps(ByVal ControlID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal Text As String, ByVal Xml As String, ByVal Target As Object, ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, ByVal ActiveXPropBag As Microsoft.Office.Interop.SmartTag.ISmartDocProperties) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateActiveXProps

	End Sub

	Public Sub PopulateCheckbox(ByVal ControlID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal Text As String, ByVal Xml As String, ByVal Target As Object, ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, ByRef Checked As Boolean) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateCheckbox

	End Sub

	Public Sub PopulateDocumentFragment(ByVal ControlID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal Text As String, ByVal Xml As String, ByVal Target As Object, ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, ByRef DocumentFragment As String) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateDocumentFragment

	End Sub

	Public Sub PopulateHelpContent(ByVal ControlID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal Text As String, ByVal Xml As String, ByVal Target As Object, ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, ByRef Content As String) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateHelpContent

	End Sub

	Public Sub PopulateImage(ByVal ControlID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal Text As String, ByVal Xml As String, ByVal Target As Object, ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, ByRef ImageSrc As String) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateImage

	End Sub



	Public Sub PopulateOther(ByVal ControlID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal Text As String, ByVal Xml As String, ByVal Target As Object, ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateOther

	End Sub

	Public Sub PopulateRadioGroup(ByVal ControlID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal Text As String, ByVal Xml As String, ByVal Target As Object, ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, ByRef List As System.Array, ByRef Count As Integer, ByRef InitialSelected As Integer) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateRadioGroup

	End Sub




#End Region


#Region "Helper Functions"
	Public Function FillListOrComboControlWithDBData(ByVal strSQL As String, _
	 ByVal strConnect As String, ByRef iCount As Long, ByRef List As Array)

		Try

      Dim rs As New ADODB.Recordset
			'Dim ary

			rs.Open(strSQL, strConnect, ADODB.CursorTypeEnum.adOpenStatic)

			'If Not DataCache Is Nothing Then 'Must be the first time to fill with data.
			'	iCount = DataCache.GetUpperBound(0) + 1
			'	List = DataCache
			'	For i = i To iCount
			'		List(i) = DataCache(0, i).value
			'	Next

      'Else

			If Not rs.EOF Then
				rs.MoveLast()
				rs.MoveFirst()
				iCount = rs.RecordCount

				Do Until rs.EOF
					List(rs.AbsolutePosition) = rs(1).Value
					rs.MoveNext()
				Loop

			End If

			rs.Close()
			'End If



		Catch ex As Exception
			MsgBox(Err.Description)
		End Try

	End Function

#End Region

End Class

' © 2008 - 2018 Arno van Boven - Vovin IT Services
' You may freely use this code for commercial and non-commercial purposes,
' but please be so kind as to let my name in.

' Requires Commence RM 3.1 or higher

' For documentation on the ActiveX controls, see the ActiveX Control reference on MSDN.
' All of the controls on this form are part of a group of custom controls that are found in the MSCOMCTL.OCX file.
' When distributing your application, install the MSCOMCTL.OCX file in the user's Microsoft Windows SYSTEM folder.
' Commence ships with this library by default, and virtually all Windows machines have it, so you *should* be fine.

' IMPORTANT NOTE:
' This script DOES NOT follow the guidelines of how to script ActiveX controls as CCorp has (unofficially) published.
' Scripting ActiveX controls the CCorp way is sometimes the only way, but it is a completely unintelligible one.
' The main difference is how controls are referenced:
'
' in this script, I will use
' Set ctl = Form.RunTime.GetControl(ActivexControlName)
'
' whereas CCorp says it should be done thus:
' Set ctl = Form.Control(ActivexControlName)
'
' The way in which properties and methods etc. are controlled is completely different for these methods!
' The methods used here are the same as how you would use them from a regular VB program
' It is VITAL that you understand that NOT ALL ActiveX controls can be controlled this way!!
' Commence -despite several pleas on my behalf- has so far not revealed why this is the case, or when to use which method

' This script is available free of charge and without any limitations, but:
' If you use and/or deploy it, please be so kind as to let my name in
' Created by: Arno van Boven # viewfinder@vovin.nl

Option Explicit

' Common Controls constants
' not all of these are used, but when a property is manipulated I include the full set of possible values
' --- ListView constants ---
'style
Const lvwIcon = 0 '(Default) Icon. Each ListItem object is represented by a full-sized (standard) icon and a text label. 
Const lvwSmallIcon = 1 'SmallIcon. Each ListItem is represented by a small icon and a text label that appears to the right of the icon. The items appear horizontally. 
Const lvwList = 2 'List. Each ListItem is represented by a small icon and a text label that appears to the right of the icon. Each ListItem appears vertically and on its own line with information arranged in columns. 
Const lvwReport = 3 ' Report. Each ListItem is displayed with its small icons and text labels. You can provide additional information about each ListItem. The icons, text labels, and information appear in columns with the leftmost column containing the small icon, followed by the text label. Additional columns display the text for each of the item's subitems. 
'labeledit
Const lvwAutomatic = 0	'Label Editing is automatic 
Const lvwManual = 1	'Label Editing must be invoked 
'sorting
Const lvwAscending = 0 	'(Default) Ascending order.
Const lvwDescending  = 1 	'Descending order.
' --- Statusbar constants ---
Const sbrNormal = 0	'Normal. StatusBar is divided into panels. 
Const sbrSimple = 1	'Simple. StatusBar has only one large panel and SimpleText.
'panel style
Const sbrText = 0 'Text and/or bitmap displayed. 
Const sbrCaps = 1 'Caps Lock status displayed. 
Const sbrNum = 2 'Number Lock status displayed. 
Const sbrIns = 3 'Insert key status displayed. 
Const sbrScrl = 4 'Scroll Lock status displayed. 
Const sbrTime = 5 'Time displayed in System format. 
Const sbrDate = 6 'Date displayed in System format. 
Const sbrKana = 7 'Kana. displays the letters KANA in bold when scroll lock is enabled, and dimmed when disabled. 
'bevel style
Const sbrNoBevel = 0 'No bevel. 
Const sbrInset = 1 'Bevel inset. 
Const sbrRaised = 2 'Bevel raised.
'auto-size
Const sbrNoAutoSize = 0 'No Autosizing. 
Const sbrSpring = 1 'Extra space divided among panels. 
Const sbrContents = 2 'Fit to contents. 

Const CMC_DELIM = "#@!##!@#" 'delimiter for Commence
Const VF_FILTER_DEFAULT = "-any-"

' variables that reference controls on the form
Dim oLvw 'listview
Dim oCboCats 'combobox with category filter
Dim oCboTypes 'combobox with view type filter
Dim oSearchBox 'textbox
Dim oChkCase 'checkbox
Dim oSbr 'statusbar

' ICommenceConversation object
Dim oConv   	' Commence Conversation object

' global variables
Dim oCmc 'Commence database object
Dim arrCats	'list of category names
Dim dictViews ' dictionary of views
Dim arrPanelText 'array to store text values of statusbar panels in
Dim iSortColumn 'holds the column last used for sort order

' --- events ---
Sub Form_OnLoad()
    Set oCmc = Application.Database ' get a reference to Commence
End Sub

Sub Form_OnSave()
    Call Form_OnCancel
    Form.Cancel 'there is nothing to save
End Sub

Sub Form_OnCancel()
    Call CleanUp
End Sub

Sub Form_OnEnterTab(ByVal TabName)
    ' assume form has just a single Tab, so no TabControl on form!
    Set oLvw = Form.RunTime.GetControl("ListView1")
    oLvw.FullRowSelect = True
    'create some columns
    oLvw.ColumnHeaders.Add ,,"View name", (oLvw.Width / 4) - 100
    oLvw.ColumnHeaders.Add ,,"Type", (oLvw.Width / 4) - 100
    oLvw.ColumnHeaders.Add ,,"Category", (oLvw.Width / 4) - 100
    oLvw.ColumnHeaders.Add ,,"File name", (oLvw.Width / 4) - 100
    'set ListView display mode to Report, this makes columnheader bar visible
    oLvw.View = lvwReport 'Report mode
    oLvw.SortOrder = lvwAscending 
    oLvw.Sorted = True
    oLvw.LabelEdit = lvwManual 'prevents users from being able to edit listitem values
    Set oSearchBox = Form.RunTime.GetControl("TextBox1")
    Set oCboCats = Form.RunTime.GetControl("ComboBox1")
    Set oCboTypes = Form.RunTime.GetControl("ComboBox2")
    Set oChkCase = Form.RunTime.GetControl("CheckBox1")
    Set oSbr = Form.RunTime.GetControl("StatusBar1")
    oSbr.Style = sbrNormal
    arrCats = GetCategories() 'populates arrCats withcategory names
    Call GetViewDataDictionary()
    Call PopulateComboBoxWithCategories(oCboCats, arrCats)
    Call PopulateComboBoxWithTypes(oCboTypes, dictViews)
    Call PopulateListBox(dictViews)

End Sub

Sub Form_OnClick(ByVal ControlID)

    Select Case ControlID

        Case "CommandButton1"
            Call Form_OnCancel ' will not get raised by Form.Cancel so invoke manually
            Form.Cancel ' does not raise Form_OnCancel
            Exit Sub

        Case "ComboBox1", "ComboBox2"
            Call FilterViewData(oSearchBox.Text, oCboCats.Text, oCboTypes.Text)

        Case "CheckBox1"
            ' only perform if there is text to filter for
            If Len(oSearchBox.Text ) > 0 Then
                Call FilterViewData(oSearchBox.Text, oCboCats.Text, oCboTypes.Text)
            End If

    End Select

End Sub

Sub Form_OnChange(ByVal ControlID)

    Select Case ControlID

    Case "TextBox1"
        Call FilterViewData(oSearchBox.Text, oCboCats.Text, oCboTypes.Text)

    End Select

End Sub

Sub Form_OnActiveXControlEvent(ByVal ControlName, ByVal EventName, ByVal ParameterArray)
 'pass events for each control to event handler routine for the control

	Select Case ControlName

		Case "ListView1" 'ListView1 was assigned to oLvw variable
			Call HandleListView1Events(EventName, ParameterArray)

	End Select

End Sub

' --- end events ---

' --- helper routines ---
Sub HandleListView1Events(EventName, params)
	'handle the events for the listview
	Dim oColHead

	Select Case EventName

		Case "DblClick"
			'note: does not receive any params
			If oLvw.ListItems.Count = 0 Then Exit Sub
			Call ShowView(oLvw.ListItems(oLvw.SelectedItem.Index).Text)

		Case "ColumnClick"
			'How do we get the index of the selected column?
			'params will only contain 1 element here: the text, not the key or index
			'the only way seem to be to iterate thru all headers and compare the names
			'theoretically this may fail since names need not be unique
			For Each oColHead In oLvw.ColumnHeaders
				If oColHead.Text = params(0) Then
					oLvw.SortKey = oColHead.Index - 1
					'reverse sort order if we clicked the same column
					If oColHead.Index - 1 = iSortColumn Then
						oLvw.SortOrder = ToggleSortOrder(oLvw.SortOrder)
					End If
					iSortColumn = oColHead.Index - 1 'store current columnheader
					Exit For
				End If
			Next 'oColHead

	End Select

End Sub

Sub GetViewDataDictionary
    ' create a dictionary so that we can get something along the lines of:
    ' Dictionary("ViewName1", {ViewInfo1})
    ' Dictionary("ViewName2", {ViewInfo2})
    ' Dictionary("ViewNameN", {ViewInfoN})
    ' where {ViewInfoN} refers to ViewInfo objects

    Dim i, arrViews, j, dde, buffer, oViewInfo

    Set dictViews = CreateObject("Scripting.Dictionary")

    For i = 0 To UBound(arrCats) 'iterate over categories
        arrViews = GetViewNames(arrCats(i)) 'get all views in that category
        For j = 0 To UBound(arrViews) 'iterate over views in the category
            ' construct DDE command
            ' note that GetViewDefinition is an *undocumented* function!
            dde = "[GetViewDefinition(" & dq(arrViews(j)) & "," & dq(CMC_DELIM) & ")]"
            ' View information: View Name, View Type, Category, FileName (RPX)
            ' execute DDE request
            buffer = Split(DDERequest(dde), CMC_DELIM) 'buffer is an array with details on the view
            ' store view details
            Set oViewInfo = New ViewInfo 'create a new instance of the ViewInfo class
            oViewInfo.Name = buffer(0) 'name
            oViewInfo.ViewType = buffer(1) 'type
            oViewInfo.Category = buffer(2) 'category
            oViewInfo.File = buffer(3) 'filename
            dictViews.Add oViewInfo.Name, oViewInfo 'store in dictionary
        Next 'j
    Next 'i

End Sub

Sub FilterViewData(ByVal searchText, ByVal catName, ByVal viewType) 'using an array
    ' Filter the listview
    ' We do not actually query the listview values,
    ' but instead we build a new dataset to populate it
    ' Of course we *could* just filter the ListView,
    ' but that would require that we have know what column (SubItem) to look in
    ' Using a dataset with easily identifiable object properties is simply more convenient

	Dim i, key
    ' create a dictionary we can use to filter
    ' note that we do not perform a deep-clone, i.e. the ViewInfo objects are the same
    Dim fltDict : Set fltDict = CloneDictionary(dictViews)
    Dim keys : keys = fltDict.Keys
    Dim items : items = fltDict.Items

	'filter views
    ' We perform 3 similar but slightly different operations in this routine,
    ' so it should probably be refactored, but I am too lazy to do that.

    ' filter by category name
	If catName <> VF_FILTER_DEFAULT Then
		For i = 0 To UBound(items)
            If Not items(i).Category = catName Then
                fltDict.Remove keys(i)
            End If
		Next 'i
	End If
    
    'redefine leftover keys and items
    ' this obviously screams refactoring or even recursion
    ' but I am just being lazy
    keys = fltDict.Keys
    items = fltDict.Items

    ' filter by viewtype
	If viewType <> VF_FILTER_DEFAULT Then
		For i = 0 To UBound(items)
			If Not items(i).ViewType = viewType Then
				fltDict.Remove keys(i)
			End If
		Next 'i
	End If

    'redefine leftover keys and items
    keys = fltDict.Keys
    items = fltDict.Items

    ' filter by search string
	If Len(searchText) > 0 Then
		For i = 0 To UBound(items)
            ' case insensitive
            If oChkCase.Value = 0 Then
                If Instr(1, items(i).Name, searchText, 1) = 0 Then
                    fltDict.Remove keys(i)
                End If
            Else ' case sensitive
                If Instr(items(i).Name, searchText) = 0 Then
                    fltDict.Remove keys(i)
                End If
            End If
		Next 'i
	End If

	Call PopulateListBox(fltDict) ' redraw listbox based on filtered values
    Set fltDict = Nothing
End Sub

Sub PopulateListBox(dict)
	Dim i
	Dim oListItem

	'remove any existing listitems
	oLvw.ListItems.Clear
	'loop all View objects
    Dim items : items = dict.Items
	For i = 0 To dict.Count - 1
		Set oListItem = oLvw.ListItems.Add(,,items(i).Name)
		oListItem.ToolTipText = "Double-click to open view " & oListItem.Text
		oListItem.SubItems(1) = items(i).ViewType
		oListItem.SubItems(2) = items(i).Category
		oListItem.SubItems(3) = items(i).File
	Next 'i

	'set statusbar
	Redim arrPanelText(0)
	arrPanelText(0) = oLvw.ListItems.Count & " views"
	Call AddStatusBarPanel(arrPanelText)

	'release objects
	Set oListItem = Nothing

End Sub

Sub PopulateComboBoxWithCategories(ctl, arr)
	Dim i
	ctl.AddItem VF_FILTER_DEFAULT 'always include this item
	For i = 0 To UBound(arr)
		ctl.AddItem(arr(i))
	Next 'i
	ctl.ListIndex = 0 'put focus on first item
End Sub

Sub PopulateComboBoxWithTypes(ctl, dict)
	'populate combobox with types of views
	'types of views in Commence are fixed, however, they may change
	'therefore, we will evaluate all the views in the database,
	'and create a unique entry for every viewtype found
	'i'll use a dictionary, they are easy to work with
	Dim i

	ctl.AddItem VF_FILTER_DEFAULT 'always include this item

    ' create a list of unique types
    Dim d : Set d = CreateObject("Scripting.Dictionary")
    Dim items : items = dict.Items

    For i = 0 To UBound(items)
        If Not d.Exists(items(i).ViewType) Then
            d.Add items(i).ViewType, "" ' we don't care about the item value
        End If
	Next 'i

    Dim keys : keys = d.keys

	For i = 0 To UBound(keys)
        ctl.AddItem keys(i)
	Next 'i

	ctl.ListIndex = 0 'put focus on first item
    Set d = Nothing

End Sub

Function ToggleSortOrder(ByVal i)
	If i = lvwAscending Then
		ToggleSortOrder = lvwDescending
	Else
		ToggleSortOrder = lvwAscending
	End If
End Function

Function dq(ByVal s)
	dq = Chr(34) & s & Chr(34)
End Function

Sub CleanUp() ' not needed but can't hurt
    On Error Resume Next
	Set oLvw = Nothing
	Set oSearchBox = Nothing
	Set oCboCats = Nothing
	Set oCboTypes = Nothing
	Set oChkCase = Nothing
	Set oSbr = Nothing
    Set dictViews = Nothing
    Set oCmc = Nothing
End Sub

Sub AddStatusBarPanel(ByVal arr) 'refreshes the statusbar control
	Dim i
	Dim oPnl

	'clear all panels
	oSbr.Panels.Clear
	'adds panels to Satusbar displaying string arrPnlText(index)
	For i = 0 To UBound(arr)
		Set oPnl = oSbr.Panels.Add
		oPnl.AutoSize = sbrContents
		oPnl.Text = arr(i)
		oPnl.ToolTipText = arr(i)
		Set oPnl = Nothing
	Next 'i

End Sub

Function CloneDictionary(Dict)
    Dim newDict, key
    Set newDict = CreateObject("Scripting.Dictionary")

    For Each key in Dict.Keys
    newDict.Add key, Dict(key)
    Next
    newDict.CompareMode = Dict.CompareMode

    Set CloneDictionary = newDict
End Function

' --- Commence DDE routines ---
Function GetCategoryCount()
	'request category count
	Dim request : request = "[GetCategoryCount()]"
    GetCategoryCount = DDERequest(request)
End Function

Function GetCategories()
	'get list of category names
	Dim request : request = "[GetCategoryNames(" & dq(CMC_DELIM) & ")]"
	GetCategories = Split(DDERequest(request), CMC_DELIM)
End Function

Function GetViewNames(ByVal sCategoryName)
	'get list of fieldnames
	Dim request : request = "[GetViewNames(" & dq(sCategoryName) & "," & dq(CMC_DELIM) & ")]"
    GetViewNames = Split(DDERequest(request),CMC_DELIM)
End Function

Sub ShowView(ByVal sView)
	'show view
	Dim request : request = "[ShowView(" & dq(sView) & ")]"
	DDEExecute request
End Sub

Function DDERequest(ByVal strRequest)
    'performs a DDE Request
    Set oConv = oCmc.GetConversation("Commence", "GetData")
    DDERequest = oConv.Request(strRequest)
    Set oConv = Nothing
End Function

Sub DDEExecute(ByVal strRequest)
    'performs a DDE Execute
    Set oConv = oCmc.GetConversation("Commence", "GetData")
    oConv.Execute(strRequest)
    Set oConv = Nothing
End Sub

' ==== Classes ===
Class ViewInfo
    Private m_Name
    Private m_Type
    Private m_Category
    Private m_FileName

    Property Let Name(value)
        m_Name = value
    End Property

    Property Get Name
        Name = m_Name
    End Property
        
    Property Let ViewType(value)
        m_Type = value
    End Property

    Property Get ViewType
        ViewType = m_Type
    End Property

    Property Let Category(value)
        m_Category = value
    End Property

    Property Get Category
        Category = m_Category
    End Property

    Property Let File(value)
        m_FileName = value
    End Property

    Property Get File
        File = m_FileName
    End Property

End Class
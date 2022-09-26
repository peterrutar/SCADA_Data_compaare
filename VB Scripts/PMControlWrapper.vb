Const PARAMTYPE_ANALOG = 1
Const PARAMTYPE_BINARY = 2
Const PARAMTYPE_DIGITAL = 4
Const PARAMTYPE_TEXT = 8

Const IDX_NAME = 0
Const IDX_TYPE = 1
Const IDX_VALUE = 2
Const IDX_INFO = 3
Const IDX_MIN_LIMIT = 4
Const IDX_MAX_LIMIT = 5

Sub PMControlWrapper_Init()
End Sub

Sub PMControlWrapper_AdjustFontSettings( Byval template )	'Textformatierung global einstellen

	Dim si, strScreen
	
	If HMIRuntime.ActiveScreen.ObjectName("CustomRecipe")Then
		Set strScreen = HMIRuntime.ActiveScreen
	Else 
		Set strScreen = HMIRuntime.Screens("Home.Picture:Recipe_1")
	End If	
			
	For Each si In strScreen.ScreenItems					

		If si.ObjectName <> template.ObjectName Then
		
			Select Case si.Type
				Case "HMIButton", _
					"HMIComboBox", _
					"HMITextField", _
					"HMIIOField"
					si.FontSize = template.FontSize
					si.FontBold = template.FontBold
			End Select 
		
		End If
	Next	
End Sub

Function PMControlWrapper_GetScreenItemByName( Byval name )	'Objektauswahl
	Set PMControlWrapper_GetScreenItemByName = Nothing
	
On Error Resume Next
	
		Dim strScreen 
		Set strScreen =  HMIRuntime.Screens("Home.Picture:Recipe_1")
		Set PMControlWrapper_GetScreenItemByName = strScreen.ScreenItems( name )	
		
	Exit Function		
	
End Function

Sub PMControlWrapper_FillComboBoxWithRecipeNames( Byval cbx )	'Fill in the field for recipe selection

	If IsObject( cbx ) And Not cbx Is Nothing Then
		
		Dim w: Set w = CreateObject( "PMBRecipesWrapper.PMBRecipes" )		
		Dim recipes: recipes = w.GetRecipeNames()					
		
		cbx.NumberLines = UBound( recipes ) + 1					'Anzahl der Rezepte
		
		Dim i
		For i = 0 To UBound( recipes )
			cbx.Index = i + 1			
			cbx.Text = recipes(i)
		Next 
		Set w = Nothing
	End If
End Sub

Sub PMControlWrapper_FillComboBoxFromBinaryParam( Byval cbx, Byval values )		'Felder zur Auswahl von Binärparametern ausfüllen

	If IsObject( cbx ) And Not cbx Is Nothing And VarType( values ) = vbString Then
		
		cbx.NumberLines = 2
		Dim str: str = Split( values, "|" )		
		cbx.Index = 1			
		cbx.Text = str(0)
		cbx.Index = 2
		cbx.Text = str(1)
	End If
End Sub

Sub PMControlWrapper_FillComboBoxFromDigitalParam( Byval cbx, Byval values )	'Felder zur Auswahl von Digitalparametern ausfüllen		

	If IsObject( cbx ) And Not cbx Is Nothing And VarType( values ) = vbString Then
		
		Dim str: str = Split( values, "|" )
		Dim i
		Dim count: count = 0
		
		For i = 0 To UBound( str )
			If str(i) <> "" Then
				count = count + 1
			Else
				Exit For
			End If	
		Next
		
		cbx.NumberLines = count
		
		For i = 1 To cbx.NumberLines
			cbx.Index = i
			cbx.Text = str( i - 1 )
		Next
	End If
End Sub


Sub PMControlWrapper_FillLabelFromParam( Byval stLabel, Byval label )	'Beschriftung für Parametername ausfüllen

	If IsObject( stLabel ) And Not stLabel Is Nothing And VarType( label ) = vbString Then
		
		stLabel.Text = label & ":"
	End If
End Sub

Sub PMControlWrapper_FillIOValueFromAnalogParam( Byval ioValue, Byval min, Byval max )	'Grenzen der E/A-Felder festlegen (nur für Analoge Parameter)
	If IsObject( ioValue ) _
		And Not ioValue Is Nothing _		
		And Not IsEmpty( min ) _
		And Not IsEmpty( max ) _
		And IsNumeric( min ) _
		And IsNumeric( max ) Then			'Grenzen aus dem Rezeptsystem übernehmen
		
			ioValue.LimitMin = min			
			ioValue.LimitMax = max
	
	Elseif IsObject( ioValue ) _		
		And Not ioValue Is Nothing _		
		And IsEmpty( min ) _
		And IsEmpty( max ) Then				'Es liegen keine Grenzen im Rezeptsystem vor

			ioValue.LimitMin = 0		
			ioValue.LimitMax = 1.7976931348623099e+308
		
	End If
	
End Sub


Sub PMControlWrapper_FillUnitFromAnalogParam( Byval stUnit, Byval unit )	'Beschriftung für Parametereinheit ausfüllen (nur für Analoge Parameter)
	
	If IsObject( stUnit ) And Not stUnit Is Nothing And VarType( unit ) = vbString Then
		
		stUnit.Text = unit
	End If
End Sub

Sub PMControlWrapper_LoadRecipe( Byval recipeName )		'Load parameters of the selected recipe

	Dim w: Set w = CreateObject( "PMBRecipesWrapper.PMBRecipes" )
	
	Dim i, count
	
	Dim p: p = w.GetRecipeValues( recipeName, 0 )  ' !!!! **** Preveri zakaj je potrebno dodati step kot parameter
	
	On Error Resume Next
	Count = UBound(p, 1)
	
	If Err.Number <> 0 Then
		MsgBox "Recipe could not be loaded. The recipe may be being processed in the recipe system.", "Error"
		Exit Sub		
	End If
	
	On Error Goto 0
	
	Dim ts: Set ts = HMIRuntime.Tags.CreateTagSet()		'Returns an array with the recipe parameter names, values, information, ...
	Dim strScreen, si
		
	For i = 0 To count
		
		p( i, IDX_NAME ) = Replace( p( i, IDX_NAME ), " ", "_") 'Replace space with underliner
		
		Dim tagName: tagName = "PMC_" & p( i, IDX_NAME ) & "_Value"	'Name the WinCC Variables		
		ts.Add tagName
		ts( tagName ).Value = p( i, IDX_VALUE )			'Write recipe value to the WinCC tag
		
		' Debug
		HMIRuntime.Trace tagName & vbCrLf

		Dim cbx, ioValue, stLabel, stUnit
		
		Set stLabel = PMControlWrapper_GetScreenItemByName( "PMC_" & p( i, IDX_NAME ) & "_Label" )
		PMControlWrapper_FillLabelFromParam stLabel, p( i, IDX_NAME )
				
		Select Case p( i, IDX_TYPE)						'Type selection analog/binary/digital/text
			Case PARAMTYPE_ANALOG 
				Set stUnit = PMControlWrapper_GetScreenItemByName( "PMC_" & p( i, IDX_NAME ) & "_Unit" )
				Set ioValue = PMControlWrapper_GetScreenItemByName( "PMC_" & p( i, IDX_NAME ) & "_Value" )
				
				PMControlWrapper_FillIOValueFromAnalogParam ioValue, p( i, IDX_MIN_LIMIT ), p( i, IDX_MAX_LIMIT ) 				
				
				PMControlWrapper_FillUnitFromAnalogParam stUnit, p( i, IDX_INFO ) 
										
			Case PARAMTYPE_BINARY 
				Set cbx = PMControlWrapper_GetScreenItemByName( "PMC_" & p( i, IDX_NAME ) & "_Value" )

				PMControlWrapper_FillComboBoxFromBinaryParam cbx, p( i, IDX_INFO )

			Case PARAMTYPE_DIGITAL
				Set cbx = PMControlWrapper_GetScreenItemByName( "PMC_" & p( i, IDX_NAME ) & "_Value" )
				
				PMControlWrapper_FillComboBoxFromDigitalParam cbx, p( i, IDX_INFO )
			
			Case PARAMTYPE_TEXT 
			
		End Select	
	Next
			
		
	Set strScreen =  HMIRuntime.Screens("Home.Picture:Recipe_1")	
	For Each si In strScreen.ScreenItems
		
		If Left( si.ObjectName, 4 ) = "PMC_" _
			And Right( si.ObjectName, 6 ) = "_Value" Then
			
			Dim ColorIndexEnabled, ColorIndexDisabled
			ColorIndexEnabled = 27
			ColorIndexDisabled = 5
			
			For i = 0 To count
				If Mid( si.ObjectName, 5, Len( si.ObjectName ) - 10 ) = p( i, IDX_NAME )Then
					si.Enabled = True
					si.BackColor = -2147483648 + ColorIndexEnabled 
					If si.Type <> "HMIComboBox" Then 			'If no combo box, then change background of unit
						strScreen.ScreenItems(Replace( si.ObjectName, "_Value", "_Unit")).BackColor = -2147483648 + ColorIndexEnabled
					End If
					Exit For
				End If
				If i = UBound(p, 1)Then
					HMIRuntime.Tags(si.ObjectName).Write 0		'Parameters that do not exist are given the value 0
					si.Enabled = False							'I/O fields are disabled for these parameters
					si.BackColor = -2147483648 + ColorIndexDisabled
					If si.Type <> "HMIComboBox" Then 			'If no combo boxes and not in the recipe, then empty and change the background of the unit
						strScreen.ScreenItems(Replace( si.ObjectName, "_Value", "_Unit")).Text = "-"
						strScreen.ScreenItems(Replace( si.ObjectName, "_Value", "_Unit")).BackColor = -2147483648 + ColorIndexDisabled
					End If
					If si.Type = "HMIComboBox" Then		
						si.SelText = " "						'Clear combo boxes
					Elseif	si.Type = "HMIIOField" Then 
						si.LimitMin = 0							'Reset limits
					End If							
				End If
			Next
		End If
	Next

	ts.Write
	
	Set ts = Nothing
	Set w = Nothing
End Sub

Function PMControlWrapper_SaveRecipe( Byval recipeName )

	Dim w: Set w = CreateObject( "PMBRecipesWrapper.PMBRecipes" )

	Dim count: count = 0
	Dim names(), values()
	
	Dim si, p
	
	p = w.GetRecipeValues( recipeName )
	
	For Each si In HMIRuntime.Screens("Home.Picture:Recipe_1").ScreenItems
			
		If Left( si.ObjectName, 4 ) = "PMC_" _
			And Right( si.ObjectName, 6 ) = "_Value" Then
				
				Dim i
				
				For i = 0 To UBound(p, 1)
					p( i, IDX_NAME ) = Replace( p( i, IDX_NAME ), " ", "_") 'Replace space with underliner
					
					If Mid( si.ObjectName, 5, Len( si.ObjectName ) - 10 ) = p( i, IDX_NAME )Then
					
						Redim Preserve names(count)
						Redim Preserve values(count)
						names( count ) = Replace( Mid( si.ObjectName, 5, Len( si.ObjectName ) - 10 ), "_", " ")'Replace Underliner with space for the later filling of the recipe system
						values(count ) = HMIRuntime.Tags( si.ObjectName ).Read
				
						count = count + 1
					End If

				Next
		End If
	Next	
		
	Dim tempValues: tempValues = values
	Dim tempNames: tempNames = names
	Dim result
	
	On Error Resume Next
	
	w.UserName = HMIRuntime.Tags("@CurrentUserName").Read
	
	result = w.SetRecipeValues( recipeName, tempNames, tempValues )
	
	PMControlWrapper_SaveRecipe = result

	Set w = Nothing
End Function
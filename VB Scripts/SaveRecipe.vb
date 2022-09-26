Sub OnClick(ByVal Item)         
	
	Dim Language
	Set Language = HMIRuntime.Tags("Language")


	Dim cbxRecipe: Set cbxRecipe = ScreenItems( "cbxRecipe" )
	Dim fptMessage: Set fptMessage = ScreenItems( "fptMessage" )

	Dim result:	result = PMControlWrapper_SaveRecipe( cbxRecipe.Text )
	
	Language.Read
	If result = vbTrue Then
		If Language.Value = "en" Then
			fptMessage.MessageText = "Data successfully" & vbCrlf & "transferred to the" & vbCrlf & "recipe system."
		Else
			fptMessage.MessageText = "Daten erfolgreich" & vbCrlf & "in das Rezeptsystem" & vbCrlf & "übertragen."
		End If
	Else
	 	If Language.Value = "en" Then
			fptMessage.MessageText = "Error" & vbCrlf & "during data transfer" & vbCrlf & "to the recipe system!"
		Else
			fptMessage.MessageText = "Fehler" & vbCrlf & "beim Übertragen der" & vbCrlf & "Daten in das Rezeptsystem!"
		End If
	End if
	fptMessage.Visible = True
End Sub
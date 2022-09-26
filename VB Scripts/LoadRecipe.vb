Sub OnClick(Byval Item)     
             
	Dim cbxRecipe: Set cbxRecipe = ScreenItems( "cbxRecipe" )
	cbxRecipe.Index = cbxRecipe.SelIndex	
	PMControlWrapper_LoadRecipe( cbxRecipe.Text )
	
End Sub
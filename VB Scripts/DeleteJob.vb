Sub OnClick(Byval Item)     
	
	Dim o1: o1 =  "CIP"	    
    Dim o2: o2 =  ScreenItems("JobID").OutputValue
    
    If IsNumeric(o2) Then   'Check if the var is Number -> Required for funct .PMCDeleteOrder
    Else
    	o2 = -1             'If NaN -> for error handling
   	End If
    
    'HMIRuntime.Trace "o2 => "	& o2 & vbCrLf

	Dim jc:	Set jc =  ScreenItems("jobcontrol")

	Dim ok:	ok = jc.PMCDeleteOrder( o1, o2 )    

    Dim infoBox: Set infoBox = ScreenItems( "Info" )
	Dim infoMessage: Set infoMessage = ScreenItems( "txt" )	
	infoBox.Visible = True
	
    If ok = True Then
        infoMessage.Text = "Job successfully" & vbCrlf & "deleted"
    Else
        Select Case jc.PMCGetLastError()
            Case 1
                infoMessage.Text = "Job can NOT be deleted!" & vbCrlf & "The call of PMCBeginCreateOrder" & vbCrlf & "is missed before PMCEndCreateOrder."
            Case 2
                infoMessage.Text = "Job can NOT be deleted!" & vbCrlf & "The production unit was not found."
            Case 3
                infoMessage.Text = "Job can NOT be deleted!" & vbCrlf & "The recipe was not found or is" & vbCrlf & "not assigned to the selected production unit."
            Case 4
                infoMessage.Text = "Job can NOT be deleted!" & vbCrlf & "The plant settings doesn't allow the splitting" & vbCrlf & "of the job quantity into several batches."
            Case 5
                infoMessage.Text = "Job can NOT be deleted!" & vbCrlf & "The single batch quantity is smaller than " & vbCrlf & "the allowed min. prod. batch quantity."
            Case 6
                infoMessage.Text = "Job can NOT be deleted!" & vbCrlf & "The single batch quantity is bigger than" & vbCrlf & "the allowed max. prod. batch quantity."
            Case 8
                infoMessage.Text = "Job can NOT be deleted!" & vbCrlf & "The remaining quantity of the batch" & vbCrlf & "is bigger than the set quantity"
            Case 11
                infoMessage.Text = "Job can NOT be deleted!" & vbCrlf & "The production recipe" & vbCrlf & "couldn't be created"
            Case 12
                infoMessage.Text = "Job can NOT be deleted!" & vbCrlf & "The job was not found."
            Case 13
                infoMessage.Text = "Job can NOT be deleted!" & vbCrlf & "Error during checking the recipe parameters."
            Case 15
                infoMessage.Text = "Job can NOT be deleted!" & vbCrlf & "The job cannot be changed" & vbCrlf & "because it is currently in process."
            Case 16
                infoMessage.Text = "Job can NOT be deleted!" & vbCrlf & "The job cannot be deleted" & vbCrlf & "because it is currently in process."
            Case 17
                infoMessage.Text = "Job can NOT be deleted!" & vbCrlf & "The job name doesn't exist." & vbCrlf & "A job without name cannot be created."
            Case Else
                infoMessage.Text = "Job can NOT be deleted!" & vbCrlf & "Error Unknown " & jc.PMCGetLastError()
        End Select
    End If
		
	Set jc  = Nothing
	Set o1  = Nothing
    Set o2  = Nothing
      
End Sub
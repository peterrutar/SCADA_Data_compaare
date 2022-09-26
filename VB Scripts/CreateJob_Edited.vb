Sub OnClick(Byval Item)                                                              
	Dim o1, o2, o3, o4, o5, o6, o7, o8, o9, o10, o11, o12, o13, o14, o15, o16, o17, o18, o19, o20, o21

    Dim result: result = 1 

    'Order Name - REQUIRED
	o1 = ScreenItems("JobNumber").OutputValue	
	'Order Short Name 
    o2 = ""
	'Order Number
    o3 = ""
	'Creator Name
    o4 = ""
	'Client Name
    o5 = ""
	'Product Target
    o6 = ""
	'Customer order number
    o7 = ""
	'Free wildcard 1 
    o8 = ""
	'Free wildcard 2
    o9 = ""
	'Free wildcard 3
    o10 = ""
	'Free wildcard 4
    o11 = ""
	'Free wildcard 5
    o12 = ""
	'Free wildcard 6
    o13 = ""
	'Free wildcard 7
    o14 = ""
	'Free wildcard 8    
    o15 = ""
	'Free wildcard 9
    o16 = ""
	'Free wildcard 10
    o17 = ""
	'Notes
    o18 = ""
	'Planned duration in minutes - REQIRED (long)
    o19 = 1
	'Name of the production unit - REQIRED, the same as in Topology
    o20 = "CIP"
	'Recipe Name - REQIRED, must exist
    o21 = ScreenItems("JobRecipe").OutputValue

	Dim jc
	Set jc =  ScreenItems("jobcontrol")

	Dim ok 
	ok = jc.PMCBeginCreateOrder()    'initialize
	
	'HMIRuntime.Trace "jc.PMCBeginCreateOrder() => "	& ok & vbCrLf

	If ok = True Then
		
	   ok = jc.PMCCreateOrder (   o1, _  
								  o2, _ 
								  o3, _ 
								  o4, _ 
 								  o5, _ 
								  o6, _ 
								  o7, _ 
								  o8, _ 
								  o9, _ 
								 o10, _ 
								 o11, _ 
								 o12, _ 
 								 o13, _ 
								 o14, _ 
								 o15, _ 
								 o16, _ 
								 o17, _ 
								 o18, _ 
                                 o19, _
								 o20, _
								 o21 )  
				
		If ok = True Then
			
            ' 1. Generate batch: batch designation, batch quantity, remaining quantity
      		ok = jc.PMCCreateLoad( "", 1, 0 ) 'Batch
      		
      		'HMIRuntime.Trace "jc.PMCCreateLoad 1=> "	& ok & vbCrLf

			If ok = True Then
				
                'Enter target quantity: 1 = in absolute values, 2 = in number of batches
				jc.PMCSetOrderLoadType(1)
				
				'HMIRuntime.Trace "jc.PMCSetOrderLoadType => "	& ok & vbCrLf
				
				Dim lOrder
				lOrder = jc.PMCEndCreateOrder()		'Return value (ID) of the newly created order,
                                                    'should be saved for later editing
				
				Dim newID
        	    Set newID =  ScreenItems( "JobID" )	
        	    newID.OutputValue( lOrder )       	    
			
				'HMIRuntime.Trace "jc.PMCEndCreateOrder => "	& lOrder & vbCrLf
			Else
                result = 2
			End If
        Else
            result = 2
		End If
	Else
        result = 2	
	End If
	
    '
    Dim infoBox: Set infoBox = ScreenItems( "Info" )
	Dim infoMessage: Set infoMessage = ScreenItems( "txt" )
    infoBox.Visible = True

    If result = 1 Then
        infoMessage.Text = "Job successfully" & vbCrlf & "created!"
    Else        
        Select Case jc.PMCGetLastError()
            Case 1
                infoMessage.Text = "Job can NOT be created!" & vbCrlf & "The call of PMCBeginCreateOrder" & vbCrlf & "is missed before PMCEndCreateOrder."
            Case 2
                infoMessage.Text = "Job can NOT be created!" & vbCrlf & "The production unit was not found."
            Case 3
                infoMessage.Text = "Job can NOT be created!" & vbCrlf & "The recipe was not found or is" & vbCrlf & "not assigned to the selected production unit."
            Case 4
                infoMessage.Text = "Job can NOT be created!" & vbCrlf & "The plant settings doesn't allow the splitting" & vbCrlf & "of the job quantity into several batches."
            Case 5
                infoMessage.Text = "Job can NOT be created!" & vbCrlf & "The single batch quantity is smaller than " & vbCrlf & "the allowed min. prod. batch quantity."
            Case 6
                infoMessage.Text = "Job can NOT be created!" & vbCrlf & "The single batch quantity is bigger than" & vbCrlf & "the allowed max. prod. batch quantity."
            Case 8
                infoMessage.Text = "Job can NOT be created!" & vbCrlf & "The remaining quantity of the batch" & vbCrlf & "is bigger than the set quantity"
            Case 11
                infoMessage.Text = "Job can NOT be created!" & vbCrlf & "The production recipe" & vbCrlf & "couldn't be created"
            Case 12
                infoMessage.Text = "Job can NOT be created!" & vbCrlf & "The job was not found."
            Case 13
                infoMessage.Text = "Job can NOT be created!" & vbCrlf & "Error during checking the recipe parameters."
            Case 15
                infoMessage.Text = "Job can NOT be created!" & vbCrlf & "The job cannot be changed" & vbCrlf & "because it is currently in process."
            Case 16
                infoMessage.Text = "Job can NOT be created!" & vbCrlf & "The job cannot be deleted" & vbCrlf & "because it is currently in process."
            Case 17
                infoMessage.Text = "Job can NOT be created!" & vbCrlf & "The job name doesn't exist." & vbCrlf & "A job without name cannot be created."
            Case Else
                infoMessage.Text = "Job can NOT be created!" & vbCrlf & "Error Unknown"
        End Select
    End If


	Set jc  = Nothing
	Set o1  = Nothing
    Set o2  = Nothing
    Set o3  = Nothing
    Set o4  = Nothing   
    Set o5  = Nothing
    Set o6  = Nothing
    Set o7  = Nothing
    Set o8  = Nothing   
    Set o9  = Nothing   
    Set o10 = Nothing   
    Set o11 = Nothing   
    Set o12 = Nothing   
    Set o13 = Nothing   
    Set o14 = Nothing   
    Set o15 = Nothing   
    Set o16 = Nothing   
    Set o17 = Nothing   
    Set o18 = Nothing   
    Set o19 = Nothing   
    Set o20 = Nothing   
    Set o21 = Nothing 
    Set result = Nothing  
End Sub
Sub OnClick(Byval Item)                                    
	Dim o1, o2, o3, o4, o5, o6, o7, o8, o9, o10, o11, o12, o13, o14, o15, o16, o17, o18, o19, o20, o21
	
    'Order Name
	Set o1 =  HMIRuntime.Tags("API_OrderName") 
	o1.read
	'Order Short Name    
    Set o2 =  HMIRuntime.Tags("API_OrderShortName") 
    o2.read	  
    'Order Number
    Set o3 =  HMIRuntime.Tags("API_OrderNumber") 
    o3.read	
    'Creator Name
    Set o4 =  HMIRuntime.Tags("API_Builder") 
    o4.read   
    'Client Name  
    Set o5 =  HMIRuntime.Tags("API_ClientName") 
    o5.read	 	    
    'Product Target
    Set o6 =  HMIRuntime.Tags("API_ObjProd") 
    o6.read		    
    'Customer order number 
    Set o7 =  HMIRuntime.Tags("API_CustomerOrderNr") 
    o7.read	 
    'Free wildcard 1 
    Set o8 =  HMIRuntime.Tags("API_User1") 
    o8.read	    
    'Free wildcard 2
    Set o9 =  HMIRuntime.Tags("API_User2") 
    o9.read	     
    'Free wildcard 3
    Set o10 =  HMIRuntime.Tags("API_User3") 
    o10.read	    
    'Free wildcard 4
    Set o11 =  HMIRuntime.Tags("API_User4") 
    o11.read	    
    'Free wildcard 5
    Set o12 =  HMIRuntime.Tags("API_User5") 
    o12.read	    
    'Free wildcard 6
    Set o13 =  HMIRuntime.Tags("API_User6") 
    o13.read	    
    'Free wildcard 7
    Set o14 =  HMIRuntime.Tags("API_User7") 
    o14.read	    
    'Free wildcard 8
    Set o15 =  HMIRuntime.Tags("API_User8") 
    o15.read   
    'Free wildcard 9
    Set o16 =  HMIRuntime.Tags("API_User9") 
    o16.read	    
    'Free wildcard 10
    Set o17 =  HMIRuntime.Tags("API_User10") 
    o17.read
    'Notes
    Set o18 =  HMIRuntime.Tags("API_Remark") 
    o18.read	    
    'Planned duration in minutes 
    Set o19 =  HMIRuntime.Tags("API_PlanDuration") 
    o19.read	    
    'Name of the production unit
    Set o20 =  HMIRuntime.Tags("API_LineName") 
    o20.read	    
    'Recipe Name
    Set o21 =  HMIRuntime.Tags("API_RecipeName") 
    o21.read  	

	Dim jc
	Set jc =  ScreenItems("jobcontrol")

	Dim ok 
	ok = jc.PMCBeginCreateOrder()    'initialize
	
	HMIRuntime.Trace "jc.PMCBeginCreateOrder() => "	& ok & vbCrLf

	If ok = True Then
		
	   ok = jc.PMCCreateOrder (   o1.Value, _  
								  o2.Value, _ 
								  o3.Value, _ 
								  o4.Value, _ 
 								  o5.Value, _ 
								  o6.Value, _ 
								  o7.Value, _ 
								  o8.Value, _ 
								  o9.Value, _ 
								 o10.Value, _ 
								 o11.Value, _ 
								 o12.Value, _ 
 								 o13.Value, _ 
								 o14.Value, _ 
								 o15.Value, _ 
								 o16.Value, _ 
								 o17.Value, _ 
								 o18.Value, _ 
                                 o19.Value, _
								 o20.Value, _
								 o21.Value )  
				
		HMIRuntime.Trace "jc.PMCCreateOrder => "	& ok & vbCrLf
        '
        HMIRuntime.Trace "API_OrderName => "	& o1 & vbCrLf
        HMIRuntime.Trace "API_OrderShortName => "	& o2 & vbCrLf
        HMIRuntime.Trace "API_OrderNumber => "	& o3 & vbCrLf
        HMIRuntime.Trace "API_Builder => "	& o4 & vbCrLf
        HMIRuntime.Trace "API_ClientName => "	& o5 & vbCrLf
        HMIRuntime.Trace "API_ObjProd => "	& o6 & vbCrLf
        HMIRuntime.Trace "API_CustomerOrderNr => "	& o7 & vbCrLf
        HMIRuntime.Trace "API_User1 => "	& o8 & vbCrLf
        HMIRuntime.Trace "API_User2 => "	& o9 & vbCrLf
        HMIRuntime.Trace "API_User3 => "	& o10 & vbCrLf
        HMIRuntime.Trace "API_User4 => "	& o11 & vbCrLf
        HMIRuntime.Trace "API_User5 => "	& o12 & vbCrLf
        HMIRuntime.Trace "API_User6 => "	& o13 & vbCrLf
        HMIRuntime.Trace "API_User7 => "	& o14 & vbCrLf
        HMIRuntime.Trace "API_User8 => "	& o15 & vbCrLf
        HMIRuntime.Trace "API_User9 => "	& o16 & vbCrLf
        HMIRuntime.Trace "API_User10 => "	& o17 & vbCrLf
        HMIRuntime.Trace "API_Remark => "	& o18 & vbCrLf
        HMIRuntime.Trace "API_PlanDuration => "	& o19 & vbCrLf
        HMIRuntime.Trace "API_LineName => "	& o20 & vbCrLf
        HMIRuntime.Trace "API_RecipeName => "	& o21 & vbCrLf        

		If ok = True Then
			
            ' 1. Generate batch: batch designation, batch quantity, remaining quantity
      		ok = jc.PMCCreateLoad( "", 1, 0 ) 'Batch
      		
      		HMIRuntime.Trace "jc.PMCCreateLoad 1=> "	& ok & vbCrLf

			If ok = True Then
				
                'Enter target quantity: 1 = in absolute values, 2 = in number of batches
				jc.PMCSetOrderLoadType(1)
				
				HMIRuntime.Trace "jc.PMCSetOrderLoadType => "	& ok & vbCrLf
				
				Dim lOrder
				lOrder = jc.PMCEndCreateOrder()		'Return value (ID) of the newly created order,
                                                    'should be saved for later editing
				
				Dim newID
        	    Set newID =  HMIRuntime.Tags("API_OrderID")	
        	    newID.Write( lOrder )      	    
			
				HMIRuntime.Trace "jc.PMCEndCreateOrder => "	& lOrder & vbCrLf
				
			End If
		End If
		
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
End Sub
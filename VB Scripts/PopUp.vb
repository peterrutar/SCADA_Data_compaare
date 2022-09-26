Sub PopUp_Init()
End Sub

Sub AiPopUp (Byval prefix)
    'Variable
    Dim i, j

    'Check Which PopUp is NOT Open
    For i = 1 To 5
        If(ScreenItems("PopUp_Settings_" & i).Visible = False) Then
            'Check if PopUp for this Faceplate is allready open elsewhere
            For j = (i-1) To 1 Step -1
                If (ScreenItems("PopUp_Settings_" & j).TagPrefix = prefix) Then				
                'Exit funct because popup is allready open
                Exit Sub
                End If	
            Next
            'Define Picture Tag prefix -> from funct parameter
            ScreenItems("PopUp_Settings_" & i).TagPrefix = prefix
            'Define Picture Name -> from funct parameter
            ScreenItems("PopUp_Settings_" & i).PictureName = "AI_PopUp\AI_Settings"
            'Set Visibility of PopUp
            ScreenItems("PopUp_Settings_" & i).Visible = True		
            Exit Sub
        End If
    Next
End Sub

Sub DiPopUp (Byval prefix)
    'Variable
    Dim i, j

    'Check Which PopUp is NOT Open
    For i = 1 To 5
        If(ScreenItems("PopUp_Settings_" & i).Visible = False) Then
            'Check if PopUp for this Faceplate is allready open elsewhere
            For j = (i-1) To 1 Step -1
                If (ScreenItems("PopUp_Settings_" & j).TagPrefix = prefix) Then				
                'Exit funct because popup is allready open
                Exit Sub
                End If	
            Next
            'Define Picture Tag prefix -> from funct parameter
            ScreenItems("PopUp_Settings_" & i).TagPrefix = prefix	
            'Define Picture Name -> from funct parameter
            ScreenItems("PopUp_Settings_" & i).PictureName = "DI_PopUp\DI_Settings"	
            'Set Visibility of PopUp
            ScreenItems("PopUp_Settings_" & i).Visible = True
            Exit Sub
        End If
    Next
End Sub

Sub MotorAnalogPopUp (Byval prefix)
    'Variable
    Dim i, j

    'Check Which PopUp is NOT Open
    For i = 1 To 5
        If(ScreenItems("PopUp_Settings_" & i).Visible = False) Then
            'Check if PopUp for this Faceplate is allready open elsewhere
            For j = (i-1) To 1 Step -1
                If (ScreenItems("PopUp_Settings_" & j).TagPrefix = prefix) Then				
                'Exit funct because popup is allready open
                Exit Sub
                End If	
            Next
            'Define Picture Tag prefix -> from funct parameter
            ScreenItems("PopUp_Settings_" & i).TagPrefix = prefix	
            'Define Picture Name -> from funct parameter
            ScreenItems("PopUp_Settings_" & i).PictureName = "Motor_Analog_PopUp\Motor_Analog_Settings"	
            'Set Visibility of PopUp
            ScreenItems("PopUp_Settings_" & i).Visible = True
            Exit Sub
        End If
    Next
End Sub

Sub MotorPopUp (Byval prefix)
    'Variable
    Dim i, j

    'Check Which PopUp is NOT Open
    For i = 1 To 5
        If(ScreenItems("PopUp_Settings_" & i).Visible = False) Then
            'Check if PopUp for this Faceplate is allready open elsewhere
            For j = (i-1) To 1 Step -1
                If (ScreenItems("PopUp_Settings_" & j).TagPrefix = prefix) Then				
                'Exit funct because popup is allready open
                Exit Sub
                End If	
            Next
            'Define Picture Tag prefix -> from funct parameter
            ScreenItems("PopUp_Settings_" & i).TagPrefix = prefix	
            'Define Picture Name -> from funct parameter
            ScreenItems("PopUp_Settings_" & i).PictureName = "Motor_PopUp\Motor_Settings" 		
            'Set Visibility of PopUp
            ScreenItems("PopUp_Settings_" & i).Visible = True
            Exit Sub
        End If
    Next
End Sub

Sub MotorSiemensPopUp (Byval prefix)
    'Variable
    Dim i, j

    'Check Which PopUp is NOT Open
    For i = 1 To 5
        If(ScreenItems("PopUp_Settings_" & i).Visible = False) Then
            'Check if PopUp for this Faceplate is allready open elsewhere
            For j = (i-1) To 1 Step -1
                If (ScreenItems("PopUp_Settings_" & j).TagPrefix = prefix) Then				
                'Exit funct because popup is allready open
                Exit Sub
                End If	
            Next
            'Define Picture Tag prefix -> from funct parameter
            ScreenItems("PopUp_Settings_" & i).TagPrefix = prefix	
            'Define Picture Name -> from funct parameter
            ScreenItems("PopUp_Settings_" & i).PictureName = "Motor_Siemens_PopUp\Motor_Siemens_Settings"
            'Set Visibility of PopUp
            ScreenItems("PopUp_Settings_" & i).Visible = True
            Exit Sub
        End If
    Next
End Sub

Sub PIDPopUp (Byval prefix)
    'Variable
    Dim i, j

    'Check Which PopUp is NOT Open
    For i = 1 To 5
        If(ScreenItems("PopUp_Settings_" & i).Visible = False) Then
            'Check if PopUp for this Faceplate is allready open elsewhere
            For j = (i-1) To 1 Step -1
                If (ScreenItems("PopUp_Settings_" & j).TagPrefix = prefix) Then				
                'Exit funct because popup is allready open
                Exit Sub
                End If	
            Next
            'Define Picture Tag prefix -> from funct parameter
            ScreenItems("PopUp_Settings_" & i).TagPrefix = prefix
            'Define Picture Name -> from funct parameter
            ScreenItems("PopUp_Settings_" & i).PictureName = "PID_PopUp\PID_Settings"
            'Set Visibility of PopUp
            ScreenItems("PopUp_Settings_" & i).Visible = True		
            Exit Sub
        End If
    Next
End Sub

Sub ValvePopUp (Byval prefix)
    'Variable
    Dim i, j

    'Check Which PopUp is NOT Open
    For i = 1 To 5
        If(ScreenItems("PopUp_Settings_" & i).Visible = False) Then
            'Check if PopUp for this Faceplate is allready open elsewhere
            For j = (i-1) To 1 Step -1
                If (ScreenItems("PopUp_Settings_" & j).TagPrefix = prefix) Then				
                'Exit if parameters are set
                Exit Sub
                End If	
            Next
            'Define Picture Tag prefix -> from funct parameter
            ScreenItems("PopUp_Settings_" & i).TagPrefix = prefix
            'Define Picture Name -> Motor
            ScreenItems("PopUp_Settings_" & i).PictureName = "Valve_PopUp\Valve_Settings"	
            'Set Visibility of PopUp
            ScreenItems("PopUp_Settings_" & i).Visible = True
            Exit Sub
        End If
    Next
End Sub
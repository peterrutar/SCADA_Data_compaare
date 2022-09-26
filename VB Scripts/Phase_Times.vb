Sub Phase_Times () 

Dim start, ende
Dim completetime

Dim varPhases
Dim phases, p,name,value, nRow

start = CDate(CurrentBatch.FormatUTCDateTime  ( CurrentBatch.StartTimeUTC))

If Len (CurrentBatch.EndTimeUTC) > 0 Then  
  ende = CDate(CurrentBatch.FormatUTCDateTime  ( CurrentBatch.EndTimeUTC))
else
 ende = Now 
End If 

completetime = CDbl((ende-start)*86400)


Set varPhases = Variables("RelativePhases")
Set phases = CurrentBatch.Phases


varPhases.ReDim phases.Count +1, 2,vbTrue 
nRow = 1
For Each p In phases 

        Dim startphase
        Dim phasetime

        startphase = CDate(CurrentBatch.FormatUTCDateTime (p.TimeUTC))
        phasetime = (startphase-start)*86400
        start = startphase

        value = phasetime/completetime*100

        If (value <> 0) and (Len (name) > 0) Then 
            varPhases(nRow,0).Value = name        
            varPhases(nRow,1).Value =  FormatNumber(value,2) & "%"         
            nRow = nRow + 1
        End If 

        name = p.Name
Next


value = (ende-start)/completetime*100*86400

If (value <> 0) and (Len (name) > 0) Then 
  varPhases(nRow,0).Value = name        
  varPhases(nRow,1).Value =  FormatNumber(value,2) & "%"  
else
 nRow = nRow-1       
End If 

varPhases.ReDim nRow+1, 2,vbTrue 

End Sub
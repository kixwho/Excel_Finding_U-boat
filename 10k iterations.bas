Attribute VB_Name = "Module2"
Sub RunSimulation()

    Dim shipX As Double, shipY As Double
    Dim startX As Double, startY As Double
    Dim x As Double, y As Double
    Dim dx As Double, dy As Double
    Dim i As Long, m As Long
    Dim successCount As Long
    Dim pick As Long
    
    Dim lookup(1 To 8, 1 To 2) As Double
    Dim r As Long
    
    ' ---- Read lookup table A3:C10 ----
    For r = 1 To 8
        lookup(r, 1) = Range("B" & (r + 2)).Value   ' dx
        lookup(r, 2) = Range("C" & (r + 2)).Value   ' dy
    Next r
    
    ' ---- Read ship location ----
    shipX = Range("E3").Value
    shipY = Range("F3").Value
    
    ' ---- Read starting position ----
    startX = Range("B16").Value
    startY = Range("C16").Value
    
    successCount = 0
    
    ' ---- Main Monte Carlo loop ----
    For i = 1 To 10000
        
        x = startX
        y = startY
        
        For m = 1 To 100
            
            ' uniform random pick 1 to 8
            pick = Int(8 * Rnd) + 1
            
            dx = lookup(pick, 1)
            dy = lookup(pick, 2)
            
            x = x + dx
            y = y + dy
            
            If x = shipX And y = shipY Then
                successCount = successCount + 1
                Exit For
            End If
            
        Next m
        
    Next i
    
    ' ---- Write result ----
    Range("E11").Value = successCount / 10000#
    Range("E12").Value = Format(successCount / 10000#, "0.00%")
    
    MsgBox "Simulation complete. Success rate = " & Format(successCount / 10000#, "0.00%")

End Sub


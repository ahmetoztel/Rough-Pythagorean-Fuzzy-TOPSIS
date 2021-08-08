

Type Member
m As Double
v As Double
End Type

Type RRough
u As Double
l As Double
End Type

Type Rough
u As Member
l As Member
End Type









Sub RPF_TOPSIS()

Dim x, z As Integer

'x = InputBox("Alternatif sayısını giriniz")
'y = InputBox("Kriter sayısını giriniz")
'z = InputBox("Uzman sayısını giriniz")
x = 12
y = 8
z = 9
Dim EXP() As Member
ReDim EXP(z, x, x) As Member

Dim Karar() As Rough
ReDim Karar(z, x, x) As Rough

Dim rkarar() As Rough
ReDim rkarar(x, x) As Rough

Dim difmat() As RRough
ReDim difmat(x, x) As RRough

Dim mulmat() As RRough
ReDim mulmat(x, x) As RRough

Dim detmat() As Double
ReDim detmat(x, x) As Double

Dim wemat() As Double
ReDim wemat(x, x) As Double

Dim up As Member
Dim lo As Member
Dim sayu As Member
Dim sayl As Member


Dim Weight() As Double
ReDim Weight(y) As Double

Dim disp() As Double
ReDim disp(x) As Double

Dim disn() As Double
ReDim disn(x) As Double

Dim RC() As Double
ReDim RC(x) As Double

For k = 1 To z
For j = 1 To y

    For i = 1 To x
    
        
           
        
        
        
        
        
        EXP(k, i, j).m = 0.1 + (Cells(i + (k - 1) * x, j) - 1) * 0.2
        EXP(k, i, j).v = 0.9 - (Cells(i + (k - 1) * x, j) - 1) * 0.2

        
        
                    
        
        Cells(i + (k - 1) * x, y + 2 + j) = CStr(EXP(k, i, j).m) & ";" & CStr(EXP(k, i, j).v)
        
                
        
       
 
     
            
            
    Next i

Next j
Next k


Top = 0
For i = 1 To x
    For j = 1 To y
    

        For t = 1 To z
        up.m = 0
        lo.m = 0
       
        sayu.m = 0
        sayl.m = 0
        
        up.v = 0
        lo.v = 0
       
        sayu.v = 0
        sayl.v = 0
        
            For s = 1 To z
                If EXP(t, i, j).m <= EXP(s, i, j).m Then
                
                sayu.m = sayu.m + 1
                up.m = up.m + EXP(s, i, j).m
                
                End If
                
                
                If EXP(t, i, j).m >= EXP(s, i, j).m Then
                
                sayl.m = sayl.m + 1
                lo.m = lo.m + EXP(s, i, j).m
                
                End If
                
                If EXP(t, i, j).v <= EXP(s, i, j).v Then
                
                sayu.v = sayu.v + 1
                up.v = up.v + EXP(s, i, j).v
                
                End If
                
                
                If EXP(t, i, j).v >= EXP(s, i, j).v Then
                
                sayl.v = sayl.v + 1
                lo.v = lo.v + EXP(s, i, j).v
                
                End If
                
              
            Next s
                
           Karar(t, i, j).l.m = lo.m / sayl.m
           Karar(t, i, j).u.m = up.m / sayu.m
           
           Karar(t, i, j).l.v = lo.v / sayl.v
           Karar(t, i, j).u.v = up.v / sayu.v
        
        Cells(i + (t - 1) * x, 2 * y + 4 + j) = "[" & CStr(Round(Karar(t, i, j).l.m, 3)) & ";" & CStr(Round(Karar(t, i, j).u.m, 3)) & "]" & ";" & "[" & CStr(Round(Karar(t, i, j).l.v, 3)) & ";" & CStr(Round(Karar(t, i, j).u.v, 3)) & "]"
        
        rkarar(i, j).l.m = rkarar(i, j).l.m + Karar(t, i, j).l.m
        rkarar(i, j).u.m = rkarar(i, j).u.m + Karar(t, i, j).u.m
        
        rkarar(i, j).l.v = rkarar(i, j).l.v + Karar(t, i, j).l.v
        rkarar(i, j).u.v = rkarar(i, j).u.v + Karar(t, i, j).u.v
        
        
        Next t
    
    rkarar(i, j).l.m = rkarar(i, j).l.m / z
    rkarar(i, j).u.m = rkarar(i, j).u.m / z
    
    rkarar(i, j).l.v = rkarar(i, j).l.v / z
    rkarar(i, j).u.v = rkarar(i, j).u.v / z
    
    Cells(i, 3 * y + 6 + j) = "(" & "[" & CStr(Round(rkarar(i, j).l.m, 3)) & ";" & CStr(Round(rkarar(i, j).u.m, 3)) & "]" & ";" & "[" & CStr(Round(rkarar(i, j).l.v, 3)) & ";" & CStr(Round(rkarar(i, j).u.v, 3)) & "]" & ")"
    
    Weight(j) = Cells(x * z + 2, j)
    
    Next j
    
Next i


For i = 1 To x

For j = 1 To y
disp(i) = disp(i) + 0.5 * Weight(j) * ((1 - (rkarar(i, j).l.m) ^ 2) ^ 2 + (1 - (rkarar(i, j).u.m) ^ 2) ^ 2 + (rkarar(i, j).l.v) ^ 4 + (rkarar(i, j).u.v) ^ 4 + (2 - (rkarar(i, j).l.m) ^ 2 - (rkarar(i, j).u.m) ^ 2 - (rkarar(i, j).l.v) ^ 2 - (rkarar(i, j).u.v) ^ 2) ^ 2) ^ (0.5)
disn(i) = disn(i) + 0.5 * Weight(j) * ((1 - (rkarar(i, j).l.v) ^ 2) ^ 2 + (1 - (rkarar(i, j).u.v) ^ 2) ^ 2 + (rkarar(i, j).l.m) ^ 4 + (rkarar(i, j).u.m) ^ 4 + (2 - (rkarar(i, j).l.m) ^ 2 - (rkarar(i, j).u.m) ^ 2 - (rkarar(i, j).l.v) ^ 2 - (rkarar(i, j).u.v) ^ 2) ^ 2) ^ (0.5)
Next j

RC(i) = disn(i) / (disn(i) + disp(i))

Cells(i + 1, 4 * y + 8) = disp(i)
Cells(i + 1, 4 * y + 9) = disn(i)
Cells(i + 1, 4 * y + 11) = RC(i)
Next i
Cells(1, 4 * y + 8) = "the positive distance"
Cells(1, 4 * y + 9) = "the negative distance"
Cells(1, 4 * y + 11) = "the relative closeness"







End Sub

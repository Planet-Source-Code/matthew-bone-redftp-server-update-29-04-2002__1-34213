Attribute VB_Name = "modSecurity"
'// Global declares.
Global tmpENCDEC As String

Function EncryptINI(Strg$, Password$)

    '// Declares
    Dim B$, S$, I As Long, J As Long
    Dim A1 As Long, A2 As Long, A3 As Long, P$
    
    J = 1
   
    For I = 1 To Len(Password$)
        P$ = P$ & Asc(Mid$(Password$, I, 1))
    Next
    
    For I = 1 To Len(Strg$)
        A1 = Asc(Mid$(P$, J, 1))
        J = J + 1: If J > Len(P$) Then J = 1
        A2 = Asc(Mid$(Strg$, I, 1))
        A3 = A1 Xor A2
        B$ = Hex$(A3)
        
        If Len(B$) < 2 Then B$ = "0" + B$
        S$ = S$ + B$
    Next
    
    EncryptINI = S$
   
End Function

Function DecryptINI(Strg$, Password$)

    '// Declares
    Dim B$, S$, I As Long, J As Long
    Dim A1 As Long, A2 As Long, A3 As Long, P$
    J = 1
    
    For I = 1 To Len(Password$)
        P$ = P$ & Asc(Mid$(Password$, I, 1))
    Next
   
    For I = 1 To Len(Strg$) Step 2
        A1 = Asc(Mid$(P$, J, 1))
        J = J + 1: If J > Len(P$) Then J = 1
        B$ = Mid$(Strg$, I, 2)
        A3 = Val("&H" + B$)
        A2 = A1 Xor A3
        S$ = S$ + Chr$(A2)
    Next
    
    DecryptINI = S$
    
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Md()
Dim ldRow, lNRow, d, N, R_d, R_N_1, C_N_1, C_N_2, C_Umn, C_Umn1, R_d1, Umn As Integer

 ldRow = Cells(Columns.Count, 1).End(xlUp).Row - 1 'dlina d
 lNRow = Sheets("S2").Cells(Columns.Count, 1).End(xlUp).Row - 1 'dlina Nd
  R_d1 = 1 'row ud
  C_N_2 = 0 'col cd
    C_Umn1 = 1
 
    ' Pervii perebor d
    For d = 1 To ldRow Step 1
        
        R_d = R_d1 + 1 'row ud
        R_N_1 = 2 'row N
        C_N_1 = C_N_2 + 1 'col N
        Umn = 2 'row umnojenia
        C_Umn = C_Umn1 + 1
        
        ' Vtoroi perebor N
        For N = 1 To lNRow Step 1
            Cells(Umn, C_Umn) = 1500 * (3.1415 / 6) * (Cells(R_d, 1) ^ 3) * Sheets("S2").Cells(R_N_1, C_N_1)
            
            Umn = Umn + 1
            R_N_1 = R_N_1 + 1
            
        Next N
        
    C_Umn1 = C_Umn1 + 1
    C_N_2 = C_N_2 + 1
    R_d1 = R_d1 + 1
    
    Next d
End Sub


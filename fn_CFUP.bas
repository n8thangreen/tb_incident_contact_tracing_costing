Attribute VB_Name = "fn_CFUP"
Function CFUP(n_latent As Double)
'Application.Volatile True
CFUP = Range("c_fup_appt").Value * n_latent
End Function

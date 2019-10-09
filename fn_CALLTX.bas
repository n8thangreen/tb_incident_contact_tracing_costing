Attribute VB_Name = "fn_CALLTX"
Function CALLTX(n_latent As Double)
'Application.Volatile True
CALLTX = Range("c_Tx").Value * n_latent
End Function

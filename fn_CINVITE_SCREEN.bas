Attribute VB_Name = "fn_CINVITE_SCREEN"
Function CINVITE_SCREEN(n_id As Double, n_screen As Double)
'Application.Volatile True
PENSION = 1 + Range("p_pensionNI").Value
C_TESTS = Range("c_blood").Value * n_screen
T_ADMIN = Range("t_admin_appt").Value * n_id + Range("t_admin_post").Value * n_screen
CINVITE_SCREEN = (Range("c_apptnurse").Value * n_screen + Range("c_nurse_3_hr_outside").Value * T_ADMIN) * PENSION + C_TESTS
End Function

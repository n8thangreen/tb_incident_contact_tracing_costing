Attribute VB_Name = "fn_CSITE_SCREEN_NURSE"
Function CSITE_SCREEN_NURSE(n_id As Double, n_screen As Double)
'Application.Volatile True
PENSION = 1 + Range("p_pensionNI").Value
C_TESTS = Range("c_blood").Value * n_screen
T_ADMIN = Range("t_admin_id").Value * n_id + Range("t_admin_post").Value * n_screen

TSITE = Range("t_site_screen").Value

C_PEOPLE = ((4 * Range("c_nurse_7_hr_outside").Value + Range("c_hpp_hr_outside").Value) * TSITE + Range("c_nurse_3_hr_outside").Value * T_ADMIN) * PENSION

C_OTHER = C_TESTS + Range("c_inc_meet_BIRM").Value + (5 * Range("c_drive").Value * Range("d_site").Value)

CSITE_SCREEN_NURSE = C_PEOPLE + C_OTHER
End Function


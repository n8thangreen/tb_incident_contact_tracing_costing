Attribute VB_Name = "fn_PATH_INFORM"
Function PATH_INFORM(n_id As Double)
'Application.Volatile True
PENSION = 1 + Range("p_pensionNI").Value
PATH_INFORM = Range("c_nurse_3_hr_outside").Value * Range("t_inform_pp").Value * n_id * PENSION
End Function


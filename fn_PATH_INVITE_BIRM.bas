Attribute VB_Name = "fn_PATH_INVITE_BIRM"
Function PATH_INVITE_BIRM(n_id As Double, n_screen As Double, n_latent As Double)
'Application.Volatile True
RA = Range("c_incid_meet_salary_BIRM").Value + Range("c_phoneRA_BIRM").Value + Range("c_siteRA_BIRM").Value
'screen = CINVITE_SCREEN(n_id, n_screen) + CALLTX(n_latent) + Range("c_meeting_review_BIRM").Value
screen = CINVITE_SCREEN(n_id, n_screen) + CFUP(n_latent) + Range("c_meeting_review_BIRM").Value
PATH_INVITE_BIRM = RA + screen
End Function

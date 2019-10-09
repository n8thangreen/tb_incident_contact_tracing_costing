Attribute VB_Name = "fn_PATH_SITE_BIRM"
Function PATH_SITE_BIRM(n_id As Double, n_screen As Double, n_latent As Double)
'Application.Volatile True
RA = Range("c_incid_meet_salary_BIRM").Value + Range("c_phoneRA_BIRM").Value + Range("c_siteRA_BIRM").Value

'not including treatment in this model
'screen = CSITE_SCREEN_PHLEB(n_id, n_screen) + CALLTX(n_latent) + Range("c_meeting_review_BIRM").Value

If n_screen > 25 Then
    screen = CSITE_SCREEN_PHLEB(n_id, n_screen)
ElseIf n_screen <= 25 Then
    screen = CSITE_SCREEN_NURSE(n_id, n_screen)
Else
    screen = -999999 'error code
End If

PATH_SITE_BIRM = RA + screen + CFUP(n_latent) + Range("c_meeting_review_BIRM").Value
End Function


' The "companyName" the the full, unabbreviated name of your organization.
companyName = "COMPANY_NAME_REPLACE"
' The "companyAbbr" is the abbreviated name of your organization.
companyAbbr = "COMPANY_ABBR_REPLACE"
' The "companyDomain" is the domain to use for sending emails. Generated report emails will appear
' to have been sent by "COMPUTERNAME@domain.com"
companyDomain = Replace(Replace(GetObject("LDAP://RootDSE").Get("DefaultNamingContext"), ",DC=","."), "DC=","")
' The "toEmail" is a valid email address where notifications will be sent.
toEmail = "TO_EMAIL_REPLACE"
'The "enableEmail" setting is for enabling (TRUE) or disabling (FALSE) the sendEmail() function.
enableEmail = FALSE
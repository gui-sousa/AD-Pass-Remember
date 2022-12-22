#Import AD Module
Import-Module ActiveDirectory

#Link ADFS
$AdfsLink = "https://adfs.bwglab.tk/adfs/portal/updatepassword/"


#Create warning dates for future password expiration
$SevenDayWarnDate = (get-date).adddays(7).ToLongDateString()
$ThreeDayWarnDate = (get-date).adddays(3).ToLongDateString()
$OneDayWarnDate = (get-date).adddays(1).ToLongDateString()

#Email Variables
$MailSender = "Segurança da Informação - INFRA BWG <monitoramento@bwg.com.br>"
$Subject = "Sua senha de rede irá expirar em breve"
$EmailStub1 = "Olá!"
$EmailStub2 = "`n`nPassando pra informar que sua senha de acesso ao computador e VPN irá expirar daqui a:"
$EmailStub3 = "dia(s), na"
$EmailStub4 = ". `n`nMas não se assuste! Defina uma nova senha acessando:`n"
$SMTPServer = "aspmx.l.google.com"

#Formato UTF8
$encode = [System.Text.Encoding]::UTF8

$EmailBody = @"
<!DOCTYPE html><html id="exportHTML" style="background-color: #f0f0f0; height: 100%;"><head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
	<style type="text/css"> 
		table {border-collapse:separate;}
		.ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td {line-height: 100%;}
		.ExternalClass {width: 100%;}
	</style>
</head>
<body style="width: 100%; font-family: Tahoma, Geneva, sans-serif; color: #717171; font-weight: 300;margin:0; padding:0; font-size:17px; line-height:150%; border-image-width:0;">
	<style type="text/css"> 
		table {border-collapse:separate;}
		.ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td {line-height: 100%;}
		.ExternalClass {width: 100%;}
	</style>
	<table style="width: 100%; background-color: #f0f0f0; text-align: center; margin:0; padding:0; margin:0; padding:0; border-image-width:0; border:0; border-spacing:0; border-collapse:collapse;">
		<tbody>
			<tr style="margin:0; padding:0; border-image-width:0">
				<td style="padding-bottom: 0px; padding-top: 30px; padding-left: 15px;padding-right: 15px;">
					<table style="width: 100%; max-width: 600px; display: inline-block; vertical-align: top; background-color: #FFFFFF; text-align: center; margin:0; padding:0; border-image-width:0; border:0; border-spacing:0; border-collapse:collapse; margin-bottom:50px;">
						<tbody>
							<tr style="margin:0; padding:0; border-image-width:0; ">
								<td style="margin:0; padding:0; border-image-width:0">
									<table style="margin:0; padding:0; border-image-width:0; border:0; border-spacing:0; border-collapse:collapse;">
										<tbody>
											<tr style="margin:0; padding:0; border-image-width:0">
												<td style="margin:0; padding:0; border-image-width:0">
													<img id="template-img-header" src="https://github.com/stefanobg/email-template-generator/blob/master/assets/header.jpg?raw=true" style="max-width:600px; padding-bottom: 0; border:0; outline:0; text-decoration:none; display:block;">
												</td>
											</tr>
										</tbody>
									</table>
									<table style="margin:0; padding:0; border-image-width:0; border:0; border-spacing:0; border-collapse:collapse">
										<tbody>
											<tr style="margin:0; padding:0; border-image-width:0">
												<td style="margin:0; padding:0; border-image-width:0">
													<img id="template-img-title" src="" style="max-width:600px; padding-bottom: 0;border:0; outline:0; text-decoration:none; display:block">
													<!-- https://unsplash.com/photos/7HuBEZaehSE Image used -->
												</td>
											</tr>
										</tbody>
									</table>
									<br>
									<br>
									<table style="margin:0; padding:0; border-image-width:0; border:0; border-spacing:0; border-collapse:collapse; display: inline-block; vertical-align: top;">
										<tbody>
											<tr style="margin:0; padding:0; border-image-width:0">
												<td style="padding-bottom:50px;padding-left:50px;padding-right:50px">
													<table style="margin:0; padding:0; border-image-width:0; border:0; border-spacing:0; border-collapse:collapse; margin-top: 12px;">
														<tbody>
															<tr style="margin:0; padding:0; border-image-width:0">
																<td id="template-text" style="text-align: left; color: #717171; font-weight:100; font-size: 18px; min-width: 500px; max-width: 500px;"><p style="margin-top: -18px;"><span style="color: rgb(0, 0, 0);">Olá</span><strong style="color: rgb(0, 0, 0);"> VarNome</strong><span style="color: rgb(0, 0, 0);">!</span></p><p style="margin-top: -18px;"><br></p><p style="margin-top: -18px;"><span style="color: rgb(0, 0, 0);">Passando pra informar que sua senha de acesso ao computador e VPN irá expirar daqui a </span><strong style="color: rgb(0, 0, 0);">VarDias</strong><span style="color: rgb(0, 0, 0);"> em:</span></p><p style="margin-top: -18px;"><br></p><ul><li><strong style="color: rgb(246, 163, 35);">VarData</strong></li></ul><p style="margin-top: -18px;"><br></p><p style="margin-top: -18px;"><span style="color: rgb(0, 0, 0);">Mas não se assuste! <span class="ql-cursor">﻿</span>Defina uma nova senha clicando no botão abaixo 😉</span></p></td>
															</tr>
															<tr><td id="template-button" style="margin-top: 20px;">
																<table style="min-width:100%;border-collapse:collapse;margin-bottom:-10px;margin-top:40px;" width="100%" cellspacing="0" cellpadding="0" border="0">
																	<tbody>
																		<tr>
																			<td style="padding-top:0;padding-right:10px;padding-bottom:10px;padding-left:10px" valign="top" align="center">
																				<a id="template-button-link" href="https://adfs.bwglab.tk/adfs/portal/updatepassword/" target="_blank" style="text-decoration: none;"> 
																					<table style="border-collapse:separate!important;border-radius:5px;background-color:#333;padding-right:50px;padding-left:50px;" cellspacing="0" cellpadding="0" border="0">
																						<tbody>
																							<tr>
																								<td style="font-size:16px;padding:16px" valign="middle" align="center">
																									<a id="template-text-link" href="https://adfs.bwglab.tk/adfs/portal/updatepassword/" style="font-weight:bold;line-height:100%;text-align:center;text-decoration:none;color:#fff;display:block;" target="_blank" data-saferedirecturl="#">ALTERE SUA SENHA</a>
																								</td>
																							</tr>
																						</tbody>
																					</table>
																				</a>
																			</td>
																		</tr>
																	</tbody>
																</table>
															</td></tr>
															<tr style="margin:0; padding:0; border-image-width:0">
																<td id="template-regards" style="text-align: center; color: #717171; font-weight:100"><br><br>INFRA BWG</td>
															</tr>
														</tbody>
													</table>
												</td>
											</tr>
										</tbody>
										</table><table style="margin:0; padding:0; border-image-width:0; border:0; border-spacing:0; border-collapse:collapse; background-color:#f0f0f0; width: 100%">
											<tbody>
												<tr style="margin:0; padding:0; border-image-width:0; text-align: center; font-size: 14px; line-height: 140%">
													<td style="margin:0; padding:0; border-image-width:0">
														<br><br>
														<a target="_blank" href="https://github.com/stefanobg/email-template-generator/" style="text-decoration: none; color:#717171; font-weight: 100; font-size:16px">E-mail Template Generator</a> © 2019<br><br><br><br>
													</td>
												</tr>
											</tbody>
										</table>
									</td></tr></tbody></table>
								</td>
							</tr>
						</tbody>
					</table>
				
			
		
	

</body></html>
"@

#Find accounts that are enabled and have expiring passwords
$users = Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False -and PasswordLastSet -gt 0 } `
 -Properties "Name", "EmailAddress", "msDS-UserPasswordExpiryTimeComputed" | Select-Object -Property "Name", "EmailAddress", `
 @{Name = "PasswordExpiry"; Expression = {[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed").tolongdatestring() }}

#check password expiration date and send email on match
foreach ($user in $users) {
     if ($user.PasswordExpiry -eq $SevenDayWarnDate) {
         $days = 7
         $EmailBody = $EmailBody.Replace("VarNome",$user.name)
         $EmailBody = $EmailBody.Replace("VarDias",$days)
         $EmailBody = $EmailBody.Replace("VarData",$SevenDayWarnDate)

         Send-MailMessage -To $user.EmailAddress -From $MailSender -SmtpServer $SMTPServer -Port "25" -Subject $Subject -Body $EmailBody -BodyAsHtml -Encoding $encode
     }
     elseif ($user.PasswordExpiry -eq $ThreeDayWarnDate) {
         $days = 3
         $EmailBody = $EmailBody.Replace("VarNome",$user.name)
         $EmailBody = $EmailBody.Replace("VarDias",$days)
         $EmailBody = $EmailBody.Replace("VarData",$SevenDayWarnDate)

         Send-MailMessage -To $user.EmailAddress -From $MailSender -SmtpServer $SMTPServer -Port "25" -Subject $Subject -Body $EmailBody -BodyAsHtml -Encoding $encode
     }
     elseif ($user.PasswordExpiry -eq $oneDayWarnDate) {
         $days = 1
         $EmailBody = $EmailBody.Replace("VarNome",$user.name)
         $EmailBody = $EmailBody.Replace("VarDias",$days)
         $EmailBody = $EmailBody.Replace("VarData",$SevenDayWarnDate)

         Send-MailMessage -To $user.EmailAddress -From $MailSender -SmtpServer $SMTPServer -Port "25" -Subject $Subject -Body $EmailBody -BodyAsHtml -Encoding $encode
     }
    else {}
 }
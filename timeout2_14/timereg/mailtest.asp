	<%						
							
							bodyHTML = "Hej der"
							
							Set Mail = Server.CreateObject("Persits.MailSender") 
							
							Mail.Host = "abizmail1.abusiness.dk"
							Mail.From = "sk@outzource.dk" ' Required
							Mail.FromName = "Hellco Tools ApS - Webagent" ' Optional 
                            Mail.Port = 25

							'Mail.AddAddress "mail@dekaeresmaa.dk"  ', "Lisa"
						
							Mail.Subject = "Produkt bestilling fra hjemmesiden."
							Mail.Body = bodyHTML
							Mail.IsHTML = True
													
							'Mail.AddEmbeddedImage "f:\webs\dekaeresmaa.dk\wwwroot\ill\logo.gif", "dks"
													
							'*** skal være slået fra. Ellers vises body indhold ikke.
							'Mail.CharSet = 2
												
							Mail.AddAddress "sk@outzource.dk", "SK"
							'Mail.AddAddress "mail@hellco.dk", "Hellco Tools ApS"
							'Mail.AddBcc ""& oRec("email") '&","& oRec("navn")
							
							'** Afsender mail ***
							'On Error Resume Next
                            'Mail.Queue = True
							Mail.Send
							%>

                            Maile ner sendt

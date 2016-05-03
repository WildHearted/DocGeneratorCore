using System;
using Microsoft.Exchange.WebServices.Data;


namespace DocGeneratorCore
	{

	public class eMail
		{
		public static bool SendEmail(
			string parRecipient,
			string parSubject,
			string parBody,
               bool parSendBcc = false)
			{
			try
				{
				
				Console.WriteLine("Preaparing to Send e-mail...");
				// Configure the Web Credentials
				WebCredentials objWebCredentials = new WebCredentials(
					username: Properties.AppResources.User_Credentials_UserName,
					password: Properties.AppResources.User_Credentials_Password,
					domain: Properties.AppResources.User_Credentials_Domain);

				//WebCredentials objWebCredentials = new WebCredentials(
				//	username: "ben.vandenberg",
				//	password: "Bernice05",
				//	domain: Properties.AppResources.DocGenerator_AccountDomain);

				// Configure the Exchange Web Service
				ExchangeService objExchangeService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
				objExchangeService.Credentials = objWebCredentials;

				// Uset EWS AutoDicovery to obtain the correct EWS URL's
				// --- use the switches to show/trace the calls to AutoDiscovery - only turn it on when debugging...
				objExchangeService.TraceEnabled = false;
				objExchangeService.TraceFlags = TraceFlags.AutodiscoverConfiguration;
				try
					{
					objExchangeService.AutodiscoverUrl(emailAddress: Properties.AppResources.Email_Sender_Address,
						validateRedirectionUrlCallback: RedirectionUrlValidationCallback);

					//objExchangeService.AutodiscoverUrl(emailAddress: "ben.vandenberg@za.didata.com",
					//	validateRedirectionUrlCallback: RedirectionUrlValidationCallback);
					}
				catch(Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverRemoteException exc)
					{
					Console.WriteLine("*** ERROR ****\nSending E-mail failed with AutodiscoveryRemoteException: {0}\n{1}\nError: {2}\n{3}",
						exc.HResult, exc.Message, exc.Error, exc.StackTrace);
					return false;
					}
				catch(AutodiscoverLocalException exc)
					{
					Console.WriteLine("*** ERROR ****\nSending E-mail failed with AutodiscoveryLocalException: {0}\n{1}\nTargetSite: {2}\n{3}",
						exc.HResult, exc.Message, exc.TargetSite, exc.StackTrace);
					return false;
					}
				catch(Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverResponseException exc)
					{
					Console.WriteLine("*** ERROR ****\nSending E-mail failed with AutodiscoveryResponseException: {0}\n{1}\nErrorCode: {2}\n{3}",
						exc.HResult, exc.Message, exc.ErrorCode, exc.StackTrace);
					return false;
					}


				Console.WriteLine("Connected Exchange at URL: {0}", objExchangeService.Url);

				// Configure the Email Messsage to send...
				EmailMessage objEmailMessage = new EmailMessage(service: objExchangeService);

				// Specify the e-mail receipient, and add it the Email Message
				EmailAddress objReceipientsEmailAddress = new EmailAddress(
					address: parRecipient,
					name: "You",
					routingType: "SMTP");

				objEmailMessage.ToRecipients.Add(emailAddress: objReceipientsEmailAddress);

				// Specify the Email Message's Subject...
				objEmailMessage.Subject = parSubject;

				// Specify the Email Message's Body
				MessageBody objMessageBody = new MessageBody();
				//objMessageBody.BodyType = BodyType.HTML;
				objMessageBody.BodyType = BodyType.Text;
				objMessageBody.Text = parBody;

				objEmailMessage.Body = objMessageBody;

				// Now send the message to exchange...
				objEmailMessage.Send();

				}
			catch(AccountIsLockedException exc)
				{
				Console.WriteLine("*** ERROR ****\nSending E-mail failed with AccountIsLocekedException: {0}\n{1}\nTargetSite: {2}\n{3}",
					exc.HResult, exc.Message, exc.TargetSite, exc.StackTrace);
				return false;
				}
			catch(ServerBusyException exc)
				{
				Console.WriteLine("*** ERROR ****\nSending E-mail failed with ServerBusyException: {0}\n{1}\nErrorCode: {2}\n{3}",
					exc.HResult, exc.Message, exc.ErrorCode, exc.StackTrace);
				return false;
				}
			catch(ServiceObjectPropertyException exc)
				{
				Console.WriteLine("*** ERROR ****\nSending E-mail failed with ServiceObjectPropertyException: {0}\n{1}\nPropertyName: {2}"
					+ "\nPropertyDefinition: {3}\nStackTrace:{4}", exc.HResult, exc.Message, exc.Name, exc.PropertyDefinition, exc.StackTrace);
				return false;
				}
			catch(ServiceRequestException exc)
				{
				Console.WriteLine("*** ERROR ****\nSending E-mail failed with ServiceRequestException: {0}\nMessage: {1}\nTargetSite: {2}"
					+ "\nInnerException: {3}\nStackTrace:{4}", exc.HResult, exc.Message, exc.TargetSite, exc.InnerException, exc.StackTrace);
				return false;
				}
			catch(ServiceResponseException exc)
				{
				Console.WriteLine("*** ERROR ****\nSending E-mail failed with ServiceResponseException: Hresult: {0}\nMessage: {1}\nTargetSite: {2}"
					+ "\nInnerException: {3}\nErrorCode: {4}\nResponse: {5}\nStackTrace:{6}",
					exc.HResult, exc.Message, exc.TargetSite, exc.InnerException, exc.ErrorCode, exc.Response, exc.StackTrace);
				return false;
				}
			catch(ServiceRemoteException exc)
				{
				Console.WriteLine("*** ERROR ****\nSending E-mail failed with ServiceRemoteException: Hresult: {0}\nMessage: {1}\nTargetSite: {2}"
					+ "\nInnerException: {3}\nData: {4}\nStackTrace:{5}",
					exc.HResult, exc.Message, exc.TargetSite, exc.InnerException, exc.Data, exc.StackTrace);
				return false;
				}
			catch(ServiceVersionException exc)
				{
				Console.WriteLine("*** ERROR ****\nSending E-mail failed with ServiceVersionException: \nHresult: {0}\nMessage: {1}\nTargetSite: {2}"
					+ "\nInnerException: {3}\nData: {4}\nStackTrace:{5}",
					exc.HResult, exc.Message, exc.TargetSite, exc.InnerException, exc.Data, exc.StackTrace);
				return false;
				}
			catch(Exception exc)
				{
				Console.WriteLine("*** ERROR ****\nSending E-mail failed with Exception: Hresult: {0}\nMessage: {1}\nTargetSite: {2}"
					+ "\nInnerException: {3}\nData: {4}\nStackTrace:{5}",
					exc.HResult, exc.Message, exc.TargetSite, exc.InnerException, exc.Data, exc.StackTrace);
				return false;
				}

			Console.WriteLine("E-mail was successfully send...");
			return true;
			}


		private static bool RedirectionUrlValidationCallback(string redirectionUrl)
			{
			// The default for the validation callback is to reject the URL.
			bool bresult = false;

			Uri redirectionUri = new Uri(redirectionUrl);

			// Validate the contents of the redirection URL. In this simple validation callback, the redirection 
			// URL is considered valid if it is using HTTPS to encrypt the authentication credentials. 
			if(redirectionUri.Scheme == "https")
				{
				bresult = true;
				}
			return bresult;
			}

		}
	}

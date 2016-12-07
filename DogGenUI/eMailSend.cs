using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using RazorEngine;
using RazorEngine.Templating;

namespace DocGeneratorCore
	{
	public enum enumEmailType
		{
		UserSuccessfulConfirmation = 0,
		UserErrorConfirmation = 1,
		TechnicalSupport = 2
		}

	public enum enumMessageClassification
		{
		Information = 1,
		Warning = 2,
		Error = 3
		}

	//++ eMail class
	/// <summary>
	/// 
	/// </summary>
	public class eMail
		{
		//+ Properties
		public EmailModel ConfirmationEmailModel { get; set; }
		public TechnicalSupportModel TechnicalEmailModel { get; set; }
		public string HTMLmessage { get; set; }

		//+ Class Variables
		static readonly string emailTemplateFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "EmailTemplates");

		//+ Methods

		public bool SendEmail(
			//ref CompleteDataSet parDataSet,
			string parReceipient,
			string parSubject,
			//string parBody,
			bool parSendBcc = false
			)
			{
			try
				{
				Console.WriteLine("Preaparing to Send e-mail...");
				// Configure the Web Credentials

				WebCredentials objWebCredentials = new WebCredentials(
					username: Properties.AppResources.DocGenerator_AccountName,
					password: Properties.AppResources.DocGenerator_Account_Password,
					domain: Properties.AppResources.DocGenerator_AccountDomain);

				// Configure the Exchange Web Service
				ExchangeService objExchangeService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
				objExchangeService.Credentials = objWebCredentials;

				// Use EWS AutoDicovery to obtain the correct EWS URL's
				// --- use the switches to show/trace the calls to AutoDiscovery - only turn it on when debugging...
				objExchangeService.TraceEnabled = false;
				objExchangeService.TraceFlags = TraceFlags.AutodiscoverConfiguration;
				try
					{
					objExchangeService.AutodiscoverUrl(emailAddress: Properties.AppResources.Email_Sender_Address,
						validateRedirectionUrlCallback: RedirectionUrlValidationCallback);
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
					address: parReceipient,
					name: "You",
					routingType: "SMTP");

				objEmailMessage.ToRecipients.Add(emailAddress: objReceipientsEmailAddress);

				// Specify the Email Message's Subject...
				objEmailMessage.Subject = parSubject;

				// Specify the Email Message's Body
				MessageBody objMessageBody = new MessageBody();
				objMessageBody.BodyType = BodyType.HTML;
				//objMessageBody.BodyType = BodyType.Text;
				objMessageBody.Text = this.HTMLmessage;

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

		/// <summary>
		/// Before using this method, make sure that the object ConfirmationEmail property is completely populated.
		/// </summary>
		/// <returns></returns>
		public bool ComposeHTMLemail(enumEmailType parEmailType)
			{
			// add the code to distingush email template depending on the parEmailType
			string strTemplateFile = String.Empty;
			var varTemplateSource = "";
			try
				{
				
				switch(parEmailType)
					{
					//+ UserSuccessfulConfirmation
					case (enumEmailType.UserSuccessfulConfirmation):
					// define the Email Template File path and Load the template into Razor
					strTemplateFile = Path.Combine(emailTemplateFolderPath, "HTMLuserEmail.cshtml");
					// Read the contents of the .cshtml file into a string, in order for the Razor Engine to compile and run it later on.
					varTemplateSource = File.ReadAllText(strTemplateFile);

					// Define and Load the email template into the Razor Engine
					var razorKey1 = new NameOnlyTemplateKey("EmailUserSuccessTemplateKey", ResolveType.Global, null);
					Engine.Razor.AddTemplate(razorKey1, new LoadedTemplateSource(varTemplateSource));

					// RunCompile the email with Razor and that compiled HTML email.
					StringBuilder sbEmailContent1 = new StringBuilder();
					using(StringWriter swEmailContent = new StringWriter(sbEmailContent1))
						Engine.Razor.RunCompile(razorKey1, swEmailContent, null, this.ConfirmationEmailModel);
						{
						this.HTMLmessage = sbEmailContent1.ToString();
						}
					break;

					//+ UserErrorConfiramation
					case (enumEmailType.UserErrorConfirmation):
					// define the Email Template File path and Load the template into Razor
					strTemplateFile = Path.Combine(emailTemplateFolderPath, "HTMLerrorEmail.cshtml");
					// Read the File into a string, in order for the Razor Engine to compile and run it later on.
					varTemplateSource = File.ReadAllText(strTemplateFile);

					// Define and Load the email template into the Razor Engine
					var razorKey2 = new NameOnlyTemplateKey("EmailUserErrorTemplateKey", ResolveType.Global, null);
					Engine.Razor.AddTemplate(razorKey2, new LoadedTemplateSource(varTemplateSource));

					// RunCompile the email with Razor and that compiled HTML email.
					StringBuilder sbEmailContent2 = new StringBuilder();
					using(StringWriter swEmailContent = new StringWriter(sbEmailContent2))
						Engine.Razor.RunCompile(razorKey2, swEmailContent, null, this.ConfirmationEmailModel);
						{
						this.HTMLmessage = sbEmailContent2.ToString();
						}
					break;

					//+ TechnicalSupport
					case (enumEmailType.TechnicalSupport):
					// define the Email Template File path and Load the template into Razor
					strTemplateFile = Path.Combine(emailTemplateFolderPath, "HTMLtechnicalEmail.cshtml");
					// Read the File into a string, in order for the Razor Engine to compile and run it later on.
					varTemplateSource = File.ReadAllText(strTemplateFile);

					// Define and Load the email template into the Razor Engine
					var razorKey3 = new NameOnlyTemplateKey("EmailTechnicalTemplateKey", ResolveType.Global, null);
					Engine.Razor.AddTemplate(razorKey3, new LoadedTemplateSource(varTemplateSource));

					// RunCompile the email with Razor and that compiled HTML email.
					StringBuilder sbEmailContent3 = new StringBuilder();
					using(StringWriter swEmailContent = new StringWriter(sbEmailContent3))
						Engine.Razor.RunCompile(razorKey3, swEmailContent, null, this.TechnicalEmailModel);
						{
						this.HTMLmessage = sbEmailContent3.ToString();
						}
					break;
					}

				return true;
				}
			catch(Exception exc)
				{
				Console.WriteLine("*** Exception Error\nException Hresult: {0}\n Message:{1}", exc.HResult, exc.Message);
				return false;
				}
			}
		}



	public class TechnicalSupportModel
		{
		public enumMessageClassification Classification { get; set; }
		public string EmailAddress { get; set; }
		public string MessageHeading { get; set; }
		public string Instruction { get; set; }
		public List<String> MessageLines { get; set; }
		
		}



	public class EmailModel
		{
		public string Name { get; set; }
		public string EmailAddress { get; set; }
		public int CollectionID { get; set; }
		public string CollectionTitle { get; set; }
		public string CollectionURL { get; set; }
		/// <summary>
		/// Set this value to TRUE of an error occurred and the documents in the collection could not be generated
		/// and add the error message to the Error property
		/// </summary>
		public bool Failed { get; set; }
		/// <summary>
		/// Set this property to reflect the error message that the ser will receive why the generation failed.
		/// The message will only appear if the Failed property was set to TRUE.
		/// </summary>
		public string Error { get; set; }
		public List<EmailGeneratedDocuments> EmailGeneratedDocs { get; set; }
		}

	public class EmailGeneratedDocuments
		{
		public string Title { get; set; }
		public bool IsSuccessful { get; set; }
		public string URL { get; set; }
		public List<string> Errors { get; set; }
		} // end EmailGeneratedDocuments class

	}
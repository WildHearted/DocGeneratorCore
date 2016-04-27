using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading;
using System.Text;
using System.Threading.Tasks;

namespace DocGenerator
	{

	public class eMail
		{
		public static bool SendEmail(
			string parRecipient,
			string parSubject,
			string parBody,
               bool parSendBcc = false)
			{
			// Credentials
			NetworkCredential objCredential = new NetworkCredential();
			objCredential = CredentialCache.DefaultNetworkCredentials;

			// Configure the SMTP client
			SmtpClient objSmtpClient = new SmtpClient();
			objSmtpClient.Host = Properties.AppResources.Email_SMTP_Host;
			objSmtpClient.Port = Convert.ToInt16(Properties.AppResources.Email_SMTP_Port);
			objSmtpClient.Credentials = objCredential;
			objSmtpClient.EnableSsl = true;
			objSmtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
			objSmtpClient.Timeout = Convert.ToInt16(Properties.AppResources.Email_SMTP_TimeOut);
			objSmtpClient.UseDefaultCredentials = true;

			// Specify the e-mail sender.
			// Create a mailing address that includes a UTF8 character in the display name.
			MailAddress objFromAddress = new MailAddress(
				address: Properties.AppResources.Email_Sender_Address, 
				displayName: Properties.AppResources.Email_Sender_Name + (char)0xD8 + " SDDP", 
				displayNameEncoding: Encoding.UTF8);

			// Set destinations for the e-mail message.
			MailAddress objToAddress = new MailAddress(address: parRecipient); // "ben@contoso.com");

			// Specify the message content.
			MailMessage objMessage = new MailMessage(
					from: "temp string",
					to: parRecipient);

			if(parSendBcc)
				{
				MailAddress objBcc = new MailAddress(
					address: Properties.AppResources.Email_Bcc_Address,
					displayName: "DocGenerator Technical Support");
				objMessage.Bcc.Add(objBcc);
				}

			objMessage.From = objFromAddress;
			objMessage.SubjectEncoding = Encoding.UTF8;
			objMessage.Subject = parSubject;
			objMessage.BodyEncoding = Encoding.UTF8;
			objMessage.IsBodyHtml = true;
			objMessage.Body = parBody;

			//increase the timeout to 5 minutes
			//objSmtpClient.Timeout = (60 * 5 * 1000);
			objMessage.Body = parBody.Replace("\n", "<br>");

			// No need for the attachments currently, just comment out for now.
			//if(parAttachments != null)
			//	{
			//	foreach(string attachment in parAttachments)
			//		{
			//		objMessage.Attachments.Add(new Attachment(attachment));
			//		}
			//	}
			try
				{
				objSmtpClient.Send(objMessage);
				}
			catch (Exception exc)
				{
				Console.WriteLine("*** ERROR ****\nSending E-mail failed with {0}\n{1}\n{2}", exc.HResult, exc.Message,exc.StackTrace);
				return false;
				}
			return true;
			}	
		}
	}

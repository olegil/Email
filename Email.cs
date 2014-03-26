using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Net;
using System.Net.Mail;

namespace Email
{
	/// <summary>
	/// Provides methods for e-mail processing.
	/// </summary>
	public class Email
	{
        #region Methods
		/// <summary>
		/// Sends an e-mail message using the supplied information.
		/// </summary>
		/// <param name="recipient">Semicolon-separated list of e-mail addresses</param>
		/// <param name="subject">Subject of the message</param>
		/// <param name="body">Body of the message</param>
		/// <param name="sender">Originating e-mail address</param>
		/// <param name="displayName">Name to display for the sender</param>
		/// <param name="serverName">Name of the mail server</param>
		public static void SendMail(string recipient, string subject, string body, string sender, string displayName, 
                                    string serverName)
		{
			RunSendMail(recipient, 
                        subject, 
                        body, 
                        sender, 
                        displayName, 
                        serverName, 
                        false, 
                        System.Text.Encoding.UTF8, 
                        null, 
                        "",
                        "", 
                        "", 
                        100000);
		}

        /// <summary>
        /// Sends an e-mail message using the supplied information.
        /// </summary>
        /// <param name="recipient">Semicolon-separated list of e-mail addresses</param>
        /// <param name="subject">Subject of the message</param>
        /// <param name="body">Body of the message</param>
        /// <param name="sender">Originating e-mail address</param>
        /// <param name="displayName">Name to display for the sender</param>
        /// <param name="serverName">Name of the mail server</param>
        /// <param name="replyTo">Reply-to address</param>
        public static void SendMail(string recipient, string subject, string body, string sender, string displayName,
                                    string serverName, string replyTo)
        {
            RunSendMail(recipient,
                        subject,
                        body,
                        sender,
                        displayName,
                        serverName,
                        false,
                        System.Text.Encoding.UTF8,
                        null,
                        "",
                        "",
                        replyTo,
                        100000);
        }

		/// <summary>
		/// Sends an e-mail message using the supplied information.
		/// </summary>
		/// <param name="recipient">Semicolon-separated list of e-mail addresses</param>
		/// <param name="subject">Subject of the message</param>
		/// <param name="body">Body of the message</param>
		/// <param name="sender">Originating e-mail address</param>
		/// <param name="displayName">Name to display for the sender</param>
		/// <param name="serverName">Name of the mail server</param>
		/// <param name="attachments">List of file attachments for the message</param>
        public static void SendMail(string recipient, string subject, string body, string sender, string displayName, 
                                    string serverName, IEnumerable<Attachment> attachments)
		{
			RunSendMail(recipient, 
                        subject, 
                        body, 
                        sender, 
                        displayName, 
                        serverName, 
                        false, 
                        System.Text.Encoding.UTF8, 
                        attachments, 
                        "",
                        "", 
                        "", 
                        100000);
		}

		/// <summary>
		/// Sends an e-mail message using the supplied information.
		/// </summary>
		/// <param name="recipient">Semicolon-separated list of e-mail addresses</param>
		/// <param name="subject">Subject of the message</param>
		/// <param name="body">Body of the message</param>
		/// <param name="sender">Originating e-mail address</param>
		/// <param name="displayName">Name to display for the sender</param>
		/// <param name="serverName">Name of the mail server</param>
		/// <param name="isBodyHtml">Set to True if the body contains HTML</param>
		public static void SendMail(string recipient, string subject, string body, string sender, string displayName, 
                                    string serverName, bool isBodyHtml)
		{
			RunSendMail(recipient, 
                        subject, 
                        body, 
                        sender, 
                        displayName, 
                        serverName, 
                        isBodyHtml, 
                        System.Text.Encoding.UTF8, 
                        null, 
                        "",
                        "", 
                        "", 
                        100000);
		}

		/// <summary>
		/// Sends an e-mail message using the supplied information.
		/// </summary>
		/// <param name="recipient">Semicolon-separated list of e-mail addresses</param>
		/// <param name="subject">Subject of the message</param>
		/// <param name="body">Body of the message</param>
		/// <param name="sender">Originating e-mail address</param>
		/// <param name="displayName">Name to display for the sender</param>
		/// <param name="serverName">Name of the mail server</param>
		/// <param name="isBodyHtml">Set to True if the body contains HTML</param>
		/// <param name="encoding">Text encoding for the subject and body</param>
		/// <param name="attachments">List of file attachments for the message</param>
        public static void SendMail(string recipient, string subject, string body, string sender, string displayName, 
                                    string serverName, bool isBodyHtml, System.Text.Encoding encoding, 
                                    IEnumerable<Attachment> attachments)
		{
			RunSendMail(recipient, 
                        subject, 
                        body, 
                        sender, 
                        displayName, 
                        serverName, 
                        isBodyHtml, 
                        encoding, 
                        attachments, 
                        "",
                        "", 
                        "", 
                        100000);
		}

		/// <summary>
		/// Sends an e-mail message using the supplied information.
		/// </summary>
		/// <param name="recipient">Semicolon-separated list of e-mail addresses</param>
		/// <param name="subject">Subject of the message</param>
		/// <param name="body">Body of the message</param>
		/// <param name="sender">Originating e-mail address</param>
		/// <param name="displayName">Name to display for the sender</param>
		/// <param name="serverName">Name of the mail server</param>
		/// <param name="userName">User name for the mail server</param>
		/// <param name="password">Password for the mail server</param>
		public static void SendMail(string recipient, string subject, string body, string sender, string displayName, 
                                    string serverName, string userName, string password)
		{
			RunSendMail(recipient, 
                        subject, 
                        body, 
                        sender, 
                        displayName, 
                        serverName, 
                        false, 
                        System.Text.Encoding.UTF8, 
                        null, 
                        userName,
                        password, 
                        "", 
                        100000);
		}

		/// <summary>
		/// Sends an e-mail message using the supplied information.
		/// </summary>
		/// <param name="recipient">Semicolon-separated list of e-mail addresses</param>
		/// <param name="subject">Subject of the message</param>
		/// <param name="body">Body of the message</param>
		/// <param name="sender">Originating e-mail address</param>
		/// <param name="displayName">Name to display for the sender</param>
		/// <param name="serverName">Name of the mail server</param>
		/// <param name="attachments">List of file attachments for the message</param>
		/// <param name="userName">User name for the mail server</param>
		/// <param name="password">Password for the mail server</param>
        public static void SendMail(string recipient, string subject, string body, string sender, string displayName, 
                                    string serverName, IEnumerable<Attachment> attachments, string userName, string password)
		{
			RunSendMail(recipient, 
                        subject, 
                        body, 
                        sender, 
                        displayName, 
                        serverName, 
                        false, 
                        System.Text.Encoding.UTF8, 
                        attachments, 
                        userName,
                        password, 
                        "", 
                        100000);
		}

		/// <summary>
		/// Sends an e-mail message using the supplied information.
		/// </summary>
		/// <param name="recipient">Semicolon-separated list of e-mail addresses</param>
		/// <param name="subject">Subject of the message</param>
		/// <param name="body">Body of the message</param>
		/// <param name="sender">Originating e-mail address</param>
		/// <param name="displayName">Name to display for the sender</param>
		/// <param name="serverName">Name of the mail server</param>
		/// <param name="isBodyHtml">Set to True if the body contains HTML</param>
		/// <param name="userName">User name for the mail server</param>
		/// <param name="password">Password for the mail server</param>
		public static void SendMail(string recipient, string subject, string body, string sender, string displayName, 
                                    string serverName, bool isBodyHtml, string userName, string password, int timeout = 100000)
		{
			RunSendMail(recipient, 
                        subject, 
                        body, 
                        sender, 
                        displayName, 
                        serverName, 
                        isBodyHtml, 
                        System.Text.Encoding.UTF8, 
                        null, 
                        userName,
                        password, 
                        "", 
                        timeout);
		}

		/// <summary>
		/// Sends an e-mail message using the supplied information.
		/// </summary>
		/// <param name="recipient">Semicolon-separated list of e-mail addresses</param>
		/// <param name="subject">Subject of the message</param>
		/// <param name="body">Body of the message</param>
		/// <param name="sender">Originating e-mail address</param>
		/// <param name="displayName">Name to display for the sender</param>
		/// <param name="serverName">Name of the mail server</param>
		/// <param name="isBodyHtml">Set to True if the body contains HTML</param>
		/// <param name="encoding">Text encoding for the subject and body</param>
		/// <param name="attachments">List of file attachments for the message</param>
		/// <param name="userName">User name for the mail server</param>
		/// <param name="password">Password for the mail server</param>
        public static void SendMail(string recipient, string subject, string body, string sender, string displayName, 
                                    string serverName, bool isBodyHtml, System.Text.Encoding encoding, 
                                    IEnumerable<Attachment> attachments, string userName, string password)
		{
			RunSendMail(recipient, 
                        subject, 
                        body, 
                        sender, 
                        displayName, 
                        serverName, 
                        isBodyHtml, 
                        encoding, 
                        attachments, 
                        userName,
                        password, 
                        "", 
                        100000);
		}
        #endregion

        #region Private routines
        private static void RunSendMail(string recipient, string subject, string body, string sender, string displayName, 
                                        string serverName, bool isBodyHtml, System.Text.Encoding encoding, 
                                        IEnumerable<Attachment> attachments, string userName, string password, 
                                        string replyTo, int timeout)
		{
			SmtpClient mailServer = null;
			MailMessage mailMsg = null;

			try 
            {
				mailServer = new SmtpClient(serverName);
				mailServer.Timeout = timeout;

				if (!string.IsNullOrEmpty(userName))
                {
					mailServer.Credentials = new NetworkCredential(userName, password);
				}

				mailMsg = BuildMessage(recipient, 
                                       subject, 
                                       body, 
                                       sender, 
                                       displayName, 
                                       isBodyHtml, 
                                       encoding, 
                                       attachments, 
                                       replyTo);
				mailServer.Send(mailMsg);
			}
            catch
            {
				throw;
			}
            finally
            {
				if (mailMsg != null)
                {
					mailMsg.Dispose();
				}

				if (mailServer != null)
                {
					mailServer.Dispose();
				}
			}
		}

        private static MailMessage BuildMessage(string recipient, string subject, string body, string sender, 
                                                string displayName, bool isBodyHtml, System.Text.Encoding encoding, 
                                                IEnumerable<Attachment> attachments, string replyTo)
		{
			MailMessage mailMsg = null;
			string[] toList = null;

			mailMsg = new MailMessage();
			mailMsg.IsBodyHtml = isBodyHtml;
			mailMsg.SubjectEncoding = encoding;
			mailMsg.BodyEncoding = encoding;

			toList = recipient.Split(new char[] {';'});
			for (int i = 0; i <= toList.GetUpperBound(0); i++)
            {
				mailMsg.To.Add(toList[i]);
			}

			mailMsg.From = new MailAddress(sender, displayName);
			mailMsg.Sender = new MailAddress(sender, displayName);
			mailMsg.Subject = subject;
			mailMsg.Body = body;

			if (!string.IsNullOrEmpty(replyTo))
            {
				mailMsg.ReplyToList.Add(replyTo);
				mailMsg.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
			}

			if (attachments != null)
            {
				foreach (Attachment attach in attachments)
                {
					mailMsg.Attachments.Add(new System.Net.Mail.Attachment(attach.FileName));
				}
			}

			return mailMsg;
		}
        #endregion
	}
}

//
//////////////////////////////////////////////////////////////////
//
//  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives..
//
// Source code and executable can be downloaded from
//  https://sourceforge.net/projects/mbox-viewer/  and
//  https://github.com/eneam/mboxviewer
//
//  Copyright(C) 2019  Enea Mansutti, Zbigniew Minciel
//
//  This program is free software; you can redistribute it and/or modify
//  it under the terms of the version 3 of GNU Affero General Public License
//  as published by the Free Software Foundation; 
//
//  This program is distributed in the hope that it will be useful,
//  but WITHOUT ANY WARRANTY; without even the implied warranty of
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU
//  Library General Public License for more details.
//
//  You should have received a copy of the GNU Library General Public
//  License along with this program; if not, write to the
//  Free Software Foundation, Inc., 51 Franklin St, Fifth Floor,
//  Boston, MA  02110 - 1301, USA.
//
//////////////////////////////////////////////////////////////////
//

// This is my first c# code. Need better Exception handling, etc everything needs more work -:)

using System;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Runtime.InteropServices;

using MailKit;
using MailKit.Net.Smtp;
using MailKit.Security;
using MailKit.Net.Imap;

using MimeKit;
using MimeKit.Utils;
using MimeKit.IO;
using MimeKit.Text;

using Google.Apis.Auth;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Util.Store;
using Google.Apis.Services;
using Google.Apis.Util;

namespace ForwardEmlFile
{
    class IniFile
    {
        public string path;
        public StringBuilder buff;
        public int buffLength;

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section,
            string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section,
                 string key, string def, StringBuilder retVal,
            int size, string filePath);

        /// <summary>
        /// INIFile Constructor.
        /// </summary>
        /// <PARAM name="INIPath"></PARAM>
        public IniFile(string INIPath)
        {
            path = INIPath;
            buffLength = 2048;
            buff = new StringBuilder(buffLength);
        }
        /// <summary>
        /// Write Data to the INI File
        /// </summary>
        /// <PARAM name="Section"></PARAM>
        /// Section name
        /// <PARAM name="Key"></PARAM>
        /// Key Name
        /// <PARAM name="Value"></PARAM>
        /// Value Name
        public void IniWriteValue(string Section, string Key, string Value)
        {
            WritePrivateProfileString(Section, Key, Value, this.path);
        }

        /// <summary>
        /// Read Data Value From the Ini File
        /// </summary>
        /// <PARAM name="Section"></PARAM>
        /// <PARAM name="Key"></PARAM>
        /// <PARAM name="Path"></PARAM>
        /// <returns></returns>
        public string IniReadValue(string Section, string Key)
        {

            int i = GetPrivateProfileString(Section, Key, "", this.buff,
                                            this.buffLength, this.path);
            return this.buff.ToString();
        }
    }

    class FileLogger
    {
        public string m_filePath = "";

        public FileLogger()
        {
        }
        public void Open(string filePath)
        {
            try
            {
                using (StreamWriter streamWriter = new StreamWriter(filePath, append: false))
                {
                    m_filePath = filePath;
                    streamWriter.Close();
                }
            }
            catch (Exception)
            {
                ; //ignore exception
            }
        }
        public void Log(string title, string text)
        {
            if (m_filePath.Length == 0)
                return;
            try
            {
                using (StreamWriter streamWriter = new StreamWriter(m_filePath, append: true))
                {
                    DateTime now = DateTime.Now;
                    streamWriter.WriteLine("[{0}] {1} {2}", now.ToString(), title, text);
                    streamWriter.Close();
                }
            }
            catch (Exception)
            {
                m_filePath = "";
            }
        }
        public void Log(string text)
        {
            if (m_filePath.Length == 0)
                return;
            try
            {
                using (StreamWriter streamWriter = new StreamWriter(m_filePath, append: true))
                {
                    DateTime now = DateTime.Now;
                    streamWriter.WriteLine("[{0}] {1}", now.ToString(), text);
                    streamWriter.Close();
                }
            }
            catch
            {
                m_filePath = "";
            }
        }
    }
    class FileUtils
    {
        public static bool CreateWriteCloseFile(string filePath, string text)
        {
            try
            {
                using (StreamWriter streamWriter = new StreamWriter(filePath, append: false))
                {

                    streamWriter.WriteLine("{0}", text);
                    streamWriter.Close();
                }
            }
            catch
            {
                ; // what to do ?
                return false;
            }
            return true;
        }
        public static bool CreateWriteCloseFile(string filePath, string text, string ex)
        {
            DateTime now = DateTime.Now;
            string txt = now.ToString() + " " + text + ex;
            bool retval = CreateWriteCloseFile(filePath, txt);
            return retval;
        }
    }
    class MailkitConvert
    {
        static readonly Encoding Latin1 = Encoding.GetEncoding(28591);

        /// <summary>
        /// Convert MimeMessage to string
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        public static string ToString(MimeMessage message)
        {
            using (MemoryStream MemStream = new MemoryStream())
            {
                string latin1;
                long length;

                var options = FormatOptions.Default.Clone();
                options.NewLineFormat = NewLineFormat.Dos;
                options.EnsureNewLine = true;

                message.WriteTo(options, MemStream);
                length = MemStream.Length;
                MemStream.Position = 0;

                using (var reader = new StreamReader(MemStream, Latin1))
                    latin1 = reader.ReadToEnd();

                return latin1;
            }
        }
    }

    class SMTPConfig
    {
        public SMTPConfig()
        {
            MailServiceName = "";
            SmtpServerAddress = "";
            SmtpServerPort = 0;
            UserAccount = "";
            UserPassword = "";
            EncryptionType = 0;
        }
        public override string ToString()
        {
            string outString;
            outString = "SMTP Configuration:\n"
                + "MailServiceName: " + MailServiceName + "\n"
                + "MailServiceName: " + MailServiceName + "\n"
                + "SmtpServerAddress: " + SmtpServerAddress + "\n"
                + "SmtpServerPort: " + SmtpServerPort.ToString() + "\n"
                + "UserAccount: " + UserAccount + "\n"
                //+ "UserPassword: " + UserPassword + "\n"
                + "EncryptionType: " + EncryptionType.ToString() + "\n"
                ;
            return outString;
        }
        public string MailServiceName { get; set; }
        public string SmtpServerAddress { get; set; }
        public int SmtpServerPort { get; set; }
        public string UserAccount { get; set; }
        public string UserPassword { get; set; }
        public int EncryptionType { get; set; }
    }

    class EMailInfo
    {
        public EMailInfo()
        {
            From = "";
            To = "";
            CC = "";
            BCC = "";
            EmlFilePath = "";
            TextFilePath = "";
        }
        public string From { get; set; }
        public string To { get; set; }
        public string CC { get; set; }
        public string BCC { get; set; }
        public string EmlFilePath { get; set; }
        public string TextFilePath { get; set; }
    }

    class Program
    {
        // Handle division by zero, and other unhandled exceptions
        static void MyHandler(object sender, UnhandledExceptionEventArgs args)
        {
            //Exception e = (Exception)args.ExceptionObject;
            //Console.WriteLine("MyHandler caught : " + e.Message);
            //Console.WriteLine("Runtime terminating: {0}", args.IsTerminating);
            // Server socket may linger ??
            System.Environment.Exit(ExitCodes.ExitUnhandledException);
        }

        public static class ExitCodes
        {
            public const int ExitOk = 0;
            public const int ExitCmdArguments = -11;
            public const int ExitMailAddress = -12;
            public const int ExitTcpPort = -13;
            public const int ExitTcpServer = -14;
            public const int ExitMimeMessage = -15;
            public const int ExitConnect = -16;
            public const int ExitAuthenticate = -17;
            public const int ExitSend = -18;
            public const int ExitUnhandledException = -19;
        }

        public static void Main(string[] args)
        {
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.UnhandledException += new UnhandledExceptionEventHandler(MyHandler);

            string dataPath = GetLocalAppDataPath();
            string MBoxViewerPath = Path.Combine(dataPath, "UMBoxViewer");
            string MailServicePath = Path.Combine(MBoxViewerPath, "MailService");
            string TempPath = Path.Combine(MailServicePath, "Temp");

            string okFilePath = MailServicePath + "\\ForwardMailSuccess.txt";
            string errorFilePath = MailServicePath + "\\ForwardMailError.txt";
            string errorFilePathOldInstance = MailServicePath + "\\ForwardMailError2.txt";

            System.IO.DirectoryInfo info = Directory.CreateDirectory(TempPath);

            string loggerFilePath = FindKeyinArgs(args, "--logger-file");
            var logger = new FileLogger();

            logger.Open(loggerFilePath);
            logger.Log("Logger Open");

#if DEBUG
            //System.Threading.Thread.Sleep(15000);
            //for (int i = 0; i < 10; i++)  Console.Beep();

            System.Diagnostics.Debugger.Launch();
            System.Diagnostics.Debugger.Break();

            logger.Log("Debugger.Launch()");
#endif

            try
            {
                // File.Delete doesn't seem to generate exceptions if file doesn't exist
                //if (File.Exists(okFilePath)
                File.Delete(okFilePath);
                //if (File.Exists(errorFilePath)
                File.Delete(errorFilePath);
                File.Delete(errorFilePathOldInstance);
            }
            catch (Exception ex)
            {
                string txt = String.Format("Delete Critical Files Failed\n{0}\n{1}\n{2}\n\n{3}",
                    okFilePath, errorFilePath, errorFilePathOldInstance, ex.Message);
                bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, txt);
                logger.Log("Exception in Delete Critical Files: ", ex.ToString());
                System.Environment.Exit(ExitCodes.ExitMailAddress);
            }

            int numArgs = args.GetLength(0);
            if ((numArgs <= 0) || ((numArgs % 2) != 0))
            {
                string errorText = String.Format("Invalid command argument list: {0} .", String.Join(" ", args));
                bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, errorText + "\n");
                logger.Log(errorText);
                System.Environment.Exit(ExitCodes.ExitCmdArguments);
            }
            /*
            if (numArgs <= 0)
            {
                logger.Log(@"Usage: --from addr --to addr1,addr2,.. --cc addr1,addr2,.. -bcc addr1,addr2,.. 
                    --user login-user-name --passwd --login-user-password --smtp smtp-server-name", "");
                Debug.Assert(true == false);
                System.Environment.Exit(1);
            }
            */
            string instance = "";
            IniFile smtpIni = null;
            EMailInfo mailInfo = new EMailInfo();
            SMTPConfig smtpConfig = new SMTPConfig();
            string smtpConfigFilePath = "";
            string UserPassword = "";
            string protocolLoggerFilePath = "";

            int tcpListenPort = 0;
            logger.Log("Command line argument list:");
            for (int j = 0, i = 0; j < numArgs; j = j + 2, i++)
            {
                string key = args[j];
                string val = args[j + 1];

                if (!key.StartsWith("--"))
                {
                    string errorText = String.Format("Invalid key: {0} ", key);
                    bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, errorText + "\n");
                    logger.Log(errorText);
                    System.Environment.Exit(ExitCodes.ExitCmdArguments);
                }
                if ((j + 1) >= numArgs)
                {
                    string errorText = String.Format("Found key: {0} without value.", key);
                    bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, errorText + "\n");
                    logger.Log(errorText);
                    System.Environment.Exit(ExitCodes.ExitCmdArguments);
                }
                if (key.CompareTo("--instance-id") == 0)
                {
                    instance = val;
                }
                else if (key.CompareTo("--smtp-protocol-logger") == 0)
                {
                    protocolLoggerFilePath = val;
                }
                else if (key.CompareTo("--from") == 0)
                {
                    mailInfo.From = val;
                }
                else if (key.CompareTo("--to") == 0)
                {
                    mailInfo.To = val;
                }
                else if (key.CompareTo("--cc") == 0)
                {
                    mailInfo.CC = val;
                }
                else if (key.CompareTo("--bcc") == 0)
                {
                    mailInfo.BCC = val;
                }
                else if (key.CompareTo("--user") == 0)
                {
                    mailInfo.To = val;
                }
                else if (key.CompareTo("--passwd") == 0)
                {
                    UserPassword = val;
                }
                else if (key.CompareTo("--smtp-cnf") == 0)
                {
                    smtpConfigFilePath = val;
                }
                else if (key.CompareTo("--tcp-port") == 0)
                {
                    tcpListenPort = int.Parse(val);
                }
                else if (key.CompareTo("--eml-file") == 0)
                {
                    mailInfo.EmlFilePath = val;
                }
                else if (key.CompareTo("--mail-text-file") == 0)
                {
                    mailInfo.TextFilePath = val;
                }
                else if (key.CompareTo("--logger-file") == 0)
                {
                    ; // see FindKeyinArgs(args, "--logger-file");
                }
                else
                {
                    logger.Log(String.Format("    Unknown Key: {0} {1}", args[j], args[j + 1]));
                }
                logger.Log(String.Format("    {0} {1}", args[j], args[j + 1]));
            }

            if (smtpConfigFilePath.Length == 0)
            {
                string errorText = String.Format("required --smtp-cnf command line argument missing.");
                bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, errorText + "\n");
                logger.Log(errorText);
                System.Environment.Exit(ExitCodes.ExitCmdArguments);
            }

            if (!File.Exists(smtpConfigFilePath))
            {
                string errorText = String.Format("SMTP configuration file {0} doesn't exist.", smtpConfigFilePath);
                bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, errorText + "\n");
                logger.Log(errorText);
                System.Environment.Exit(ExitCodes.ExitCmdArguments);
            }
            try
            {
                if (protocolLoggerFilePath.Length > 0)
                    File.Delete(protocolLoggerFilePath);
            }
            catch (Exception /*ex*/) {; } // ignore

            smtpIni = new IniFile(smtpConfigFilePath);

            string ActiveMailService = smtpIni.IniReadValue("MailService", "ActiveMailService");

            smtpConfig.MailServiceName = smtpIni.IniReadValue(ActiveMailService, "MailServiceName");
            smtpConfig.SmtpServerAddress = smtpIni.IniReadValue(ActiveMailService, "SmtpServerAddress");
            smtpConfig.SmtpServerPort = int.Parse(smtpIni.IniReadValue(ActiveMailService, "SmtpServerPort"));
            smtpConfig.UserAccount = smtpIni.IniReadValue(ActiveMailService, "UserAccount");
            if (UserPassword.Length > 0)
                smtpConfig.UserPassword = UserPassword;
            else
                smtpConfig.UserPassword = smtpIni.IniReadValue(ActiveMailService, "UserPassword");
            smtpConfig.EncryptionType = int.Parse(smtpIni.IniReadValue(ActiveMailService, "EncryptionType"));

            logger.Log(smtpConfig.ToString());

            // Uncomment when you exec this application from MBoxViewer
            //smtpConfig.UserPassword = "";
            if (smtpConfig.UserPassword.Length == 0)
            {
                logger.Log("Waiting to receive password");
                smtpConfig.UserPassword = WaitForPassword(tcpListenPort, logger, errorFilePath);

                if (smtpConfig.UserPassword.Length > 0)
                {
                    logger.Log("Received non empty password");
                    //logger.Log("Received non empty password: ", smtpConfig.UserPassword);
                }
                else
                    logger.Log("Received empty password");

                int found = smtpConfig.UserPassword.IndexOf(":");
                if (found <= 0)
                {
                    // Old instance , log to differnt file
                    bool rval = FileUtils.CreateWriteCloseFile(errorFilePathOldInstance, "Received invalid id:password. Exitting\n");
                    System.Environment.Exit(ExitCodes.ExitCmdArguments);
                }
                string id = smtpConfig.UserPassword.Substring(0, found);
                string passwd = smtpConfig.UserPassword.Substring(found + 1);
                smtpConfig.UserPassword = passwd;

                logger.Log("Command line instance id: ", instance);
                logger.Log("Received instance id: ", id);
                //logger.Log("Received password: ", smtpConfig.UserPassword);
                logger.Log("Received password: ", "xxxxxxxxxxxx");

                if (id.CompareTo(instance) != 0)
                {
                    // Old instance , log to different file
                    bool rval = FileUtils.CreateWriteCloseFile(errorFilePathOldInstance, "This is old instance. Exitting\n");
                    System.Environment.Exit(ExitCodes.ExitCmdArguments);
                }
            }

            MimeKit.ParserOptions opt = new MimeKit.ParserOptions();


            string FromString = mailInfo.From;
            if (mailInfo.From.Length == 0)
                FromString = smtpConfig.UserAccount;

            var From = new MailboxAddress("", FromString);
            //
            InternetAddressList CCList = new InternetAddressList();
            InternetAddressList BCCList = new InternetAddressList();
            InternetAddressList ToList = new InternetAddressList();

            try
            {
                if (mailInfo.To.Length > 0)
                    ToList = MimeKit.InternetAddressList.Parse(opt, mailInfo.To);
                if (mailInfo.CC.Length > 0)
                    CCList = MimeKit.InternetAddressList.Parse(opt, mailInfo.CC);
                if (mailInfo.BCC.Length > 0)
                    BCCList = MimeKit.InternetAddressList.Parse(opt, mailInfo.BCC);
            }
            catch (Exception ex)
            {
                bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, "Parsing Internet Address list Failed\n", ex.Message);
                logger.Log("Exception in InternetAddressList.Parse: ", ex.ToString());
                System.Environment.Exit(ExitCodes.ExitMailAddress);
            }

            //
            string emlFilePath = mailInfo.EmlFilePath;

            // create the main textual body of the message
            var text = new TextPart("plain");
            try
            {
                text.Text = File.ReadAllText(mailInfo.TextFilePath, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, "Building Mime Mesage Failed\n", ex.Message);
                logger.Log("Exception in Building Mime Message: ", ex.ToString());
                System.Environment.Exit(ExitCodes.ExitMimeMessage);
            }

            logger.Log("Forwarding Eml file:", emlFilePath);
            MimeMessage msg = null;
            try
            {
                var message = new MimeMessage();
                message = MimeKit.MimeMessage.Load(emlFilePath);

                List<MimeMessage> mimeMessages = new List<MimeMessage>();
                mimeMessages.Add(message);

                msg = BuildMimeMessageWithEmlAsRFC822Attachment(text, mimeMessages, From, ToList, CCList, BCCList);
            }
            catch (Exception ex)
            {
                bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, "Building Mime Mesage Failed\n", ex.Message);
                logger.Log("Exception in Building Mime Message: ", ex.ToString());

                System.Environment.Exit(ExitCodes.ExitMimeMessage);
            }

            logger.Log("BuildMimeMessageWithEmlAsRFC822Attachment Done");

            //string msgAsString = MailkitConvert.ToString(msg);
            //string msgAsString = msg.ToString();
            //logger.Log("\n", msgAsString);

            // OAUTH2 works on Google but requires verification by Google and it seems to be chargable option if number of users > 100
            // Another problem is that ClientSecret can't be hardcoded in the application
            // For now we will just rely on User Account and User Password for authentication
            SaslMechanism oauth2 = null; ;
            bool useOAUTH2 = false;
            if (useOAUTH2)
            {
                string appClientId = "xxxxxxxxxxxxxxxxxx.apps.googleusercontent.com";
                string appClientSecret = "yyyyyyyyyyyyyyyyyyyyyyyyyyy";

                var accessScopes = new[]
                {
                    // that is the only scope that works per info from jstedfast
                    "https://mail.google.com/",
                };

                var clientSecrets = new ClientSecrets
                {
                    ClientId = appClientId,
                    ClientSecret = appClientSecret
                };
                oauth2 = GetAuth2Token(smtpConfig.UserAccount, clientSecrets, accessScopes);
            }

            IProtocolLogger smtpProtocolLogger = null;
            if (protocolLoggerFilePath.Length > 0)
                smtpProtocolLogger = new ProtocolLogger(protocolLoggerFilePath);
            else
                smtpProtocolLogger = new NullProtocolLogger();

            using (var client = new SmtpClient(smtpProtocolLogger))
            {
                try
                {
                    client.Connect(smtpConfig.SmtpServerAddress, smtpConfig.SmtpServerPort, (SecureSocketOptions)smtpConfig.EncryptionType);
                }
                catch (Exception ex)
                {
                    bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, "Connect to SMTP Server Failed\n", ex.Message);
                    logger.Log("Exception in Connect to SMPT Server: ", ex.ToString());
                    System.Environment.Exit(ExitCodes.ExitConnect);
                }

                logger.Log(String.Format("Connected to {0} mail service", smtpConfig.MailServiceName));

#if DEBUG
                try
                {
                    PrintCapabilities(client, logger);
                }
                catch (Exception)
                {
                    ; //ignore exception
                }
#endif

                if (client.Capabilities.HasFlag(SmtpCapabilities.Authentication))
                {
                    try
                    {
                        client.Authenticate(smtpConfig.UserAccount, smtpConfig.UserPassword);
                    }
                    catch (AuthenticationException ex)
                    {
                        bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, "Send Failed. SMTP Authentication Failed.\n", ex.Message);
                        logger.Log("Send Failed. SMTP Authentication Failed: ", ex.ToString());

                        System.Environment.Exit(ExitCodes.ExitAuthenticate); ;
                    }
                    catch (SmtpCommandException ex)
                    {
                        string statusCodeStr = ex.StatusCode.ToString();
                        string errorCodeString = ex.ErrorCode.ToString();
                        string exPlus = ex.Message + " StatusCode: " + statusCodeStr + " ErrorCode:  " + errorCodeString;

                        bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, "Send Failed. Error trying to authenticate.\n", exPlus);

                        string exPlusPlus = exPlus + "\n" + ex.ToString();
                        logger.Log("SendFailed. Error trying to authenticate: ", exPlusPlus); ;

                        System.Environment.Exit(ExitCodes.ExitAuthenticate); ;
                    }
                    catch (SmtpProtocolException ex)
                    {
                        bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, "Send Failed. Protocol error while trying to authenticate.\n", ex.Message);
                        logger.Log("SendFailed. Protocol error while trying to authenticate: ", ex.ToString());

                        System.Environment.Exit(ExitCodes.ExitAuthenticate); ;
                    }
                    catch (Exception ex)
                    {
                        bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, "Send Failed. SMTP Authentication Failed\n", ex.Message);
                        logger.Log("Send Failed. Exception in SMTP Authentication: ", ex.ToString());

                        System.Environment.Exit(ExitCodes.ExitAuthenticate);
                    }

                    logger.Log("SMTP Authentication Succeeded");
                }
                else
                {
                    logger.Log("SMTP Authentication not supported by SMTP Server");
                }

#if OLD
                try
                {
                    if (useOAUTH2)
                        client.Authenticate(oauth2);
                    else
                        client.Authenticate(smtpConfig.UserAccount, smtpConfig.UserPassword);
                }
                catch (Exception ex)
                {
                    bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, "SMTP Authentication Failed\n", ex.Message);
                    logger.Log("Exception in SMTP Authentication: ", ex.ToString());
                    System.Environment.Exit(ExitCodes.ExitAuthenticate);
                }

                logger.Log("SMTP Authentication Succeeded");
#endif


#if DEBUG
                try
                {
                    PrintCapabilities(client, logger);
                }
                catch (Exception)
                {
                    ; //ignore exception
                }
#endif

                // Clear smtpConfig.UserPassword in case it cores
                smtpConfig.UserPassword = "";

                try
                {
                    client.Send(msg);
                }
                catch (SmtpCommandException ex)
                {
                    string statusCodeStr = ex.StatusCode.ToString();
                    string errorCodeString = ex.ErrorCode.ToString();
                    string exPlus = ex.Message + " StatusCode: " + statusCodeStr + " ErrorCode:  " + errorCodeString;

                    switch (ex.ErrorCode)
                    {
                        case SmtpErrorCode.RecipientNotAccepted:
                            exPlus += "\nRecipient not accepted:  " + ex.Mailbox.Address;
                            break;
                        case SmtpErrorCode.SenderNotAccepted:
                            exPlus += "\nSender not accepted:  " + ex.Mailbox.Address;
                            break;
                        case SmtpErrorCode.MessageNotAccepted:
                            exPlus += "\nMessage not accepted";
                            break;
                    }

                    bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, "Send Failed\n", exPlus);

                    string exPlusPlus = exPlus + "\n" + ex.ToString();
                    logger.Log("Send Failed. Exception in Send to SMTP Server: ", exPlusPlus);

                    System.Environment.Exit(ExitCodes.ExitSend);
                }
                catch (SmtpProtocolException ex)
                {
                    string exPlus = ex.Message;

                    bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, "Send Failed. Protocol error while sending message:\n", exPlus);

                    logger.Log("Send Failed. Protocol error while sending message: ", ex.ToString());

                    System.Environment.Exit(ExitCodes.ExitSend);
                }
                catch (Exception ex)
                {
                    string exPlus = ex.Message;

                    bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, "Send Failed. Error while sending message:\n", exPlus);

                    logger.Log("Send Failed. Error while sending message: ", ex.ToString());

                    System.Environment.Exit(ExitCodes.ExitSend);
                }
#if OLD
                //
                catch (Exception ex)
                {
                    string exPlus = ex.Message;
                    bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, "Send Failed\n", exPlus);
                    logger.Log("Send Failed. Exception in Send to SMTP Server: ", ex.ToString());
                    //logger.Log("Maximum message length reported by SMTP Server: ", client.MaxSize.ToString());

                    //string msgString = MailkitConvert.ToString(msg);
                    //string msgAsString = msg.ToString();

                    // To help to investigate Warning place at the begining of the serialized MimeMessage
                    // X - MimeKit - Warning: Do NOT use ToString() to serialize messages! Use one of the WriteTo() methods instead! 
                    //logger.Log("\n", msgString);

                    System.Environment.Exit(ExitCodes.ExitSend);
                }
#endif
                string txt = "Mail Sending Succeeded";
                logger.Log(txt);
                bool retval = FileUtils.CreateWriteCloseFile(okFilePath, txt);

                try
                {
                    client.Disconnect(true);
                }
                catch (Exception ex)
                {
                    // Ignore, not a fatal error
                    //bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, "Send Failed\n", ex.Message);
                    logger.Log("Exception in Disconnect to SMTP Server: ", ex.ToString());
                }
                logger.Log("SMTP Client Disconnected. All done.");
            }
            System.Environment.Exit(ExitCodes.ExitOk);
        }

        /// <summary>
        /// Build MimeMessage, inject user text and Eml file as message/rfc822 attachment
        /// </summary>
        /// <PARAM name="text"></PARAM>
        /// <PARAM name="emlMessageList"></PARAM>
        /// <PARAM name="from"></PARAM>
        /// <PARAM name="toList"></PARAM>
        /// <PARAM name="ccList"></PARAM>
        /// <PARAM name="bccList"></PARAM>
        /// <returns></returns>
        public static MimeMessage BuildMimeMessageWithEmlAsRFC822Attachment(TextPart text, List<MimeMessage> emlMessageList,
            MailboxAddress from, InternetAddressList toList, InternetAddressList ccList, InternetAddressList bccList)
        {
            var message = new MimeMessage();
            message.From.Add(from);
            message.To.AddRange(toList);
            if (ccList.Count > 0)
                message.Cc.AddRange(ccList);
            if (bccList.Count > 0)
                message.Bcc.AddRange(bccList);

            MimeMessage firstMessage = emlMessageList[0];
            // set the forwarded subject
            if (!firstMessage.Subject.StartsWith("FW:", StringComparison.OrdinalIgnoreCase))
                message.Subject = "FW: " + firstMessage.Subject;
            else
                message.Subject = firstMessage.Subject;

            // create a multipart/mixed container for the text body and the forwarded message
            var multipart = new Multipart("mixed");
            multipart.Add(text);
            foreach (var emlMessage in emlMessageList)
            {
                // create the message/rfc822 attachment for the original message
                var rfc822 = new MessagePart { Message = emlMessage };
                multipart.Add(rfc822);
            }

            // set the multipart as the body of the message
            message.Body = multipart;

            return message;
        }

        /// <summary>
        /// Get OAUTH2 token needed to Authenticate SMTP client
        /// </summary>
        /// <PARAM name="userAccount"></PARAM>
        /// <PARAM name="clientSecrets"></PARAM>
        /// <PARAM name="accessScopes"></PARAM>
        /// <returns></returns>
        public static SaslMechanism GetAuth2Token(string userAccount, ClientSecrets clientSecrets, string[] accessScopes)
        {
            var AppFileDataStore = new FileDataStore("CredentialCacheFolder", false);

            var codeFlow = new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
            {
                DataStore = AppFileDataStore,
                Scopes = accessScopes,
                ClientSecrets = clientSecrets
            });

            var codeReceiver = new LocalServerCodeReceiver();
            var authCode = new AuthorizationCodeInstalledApp(codeFlow, codeReceiver);

            UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                   clientSecrets
                    , accessScopes
                    , userAccount
                    , CancellationToken.None
                    , AppFileDataStore).Result;

            if (credential.Token.IsExpired(SystemClock.Default))
                credential.RefreshTokenAsync(CancellationToken.None);

            // Note: We use credential.UserId here instead of GMail account because the user *may* have chosen a
            // different GMail account when presented with the browser window during the authentication process.
            SaslMechanism oauth2;

            oauth2 = new SaslMechanismOAuth2(credential.UserId, credential.Token.AccessToken);
            return oauth2;
        }

        public static string WaitForPassword(int TcpListerPort, FileLogger logger, string errorFilePath)
        {
            TcpListener server = null;
            string password = "";
            bool exceptionFired = false;
            int ErrorCode = 0;
            SocketError SocketErrorCode = 0;

            try
            {
                IPAddress localAddr = IPAddress.Parse("127.0.0.1");

                logger.Log("Creating new  TcpListener ... ");
                server = new TcpListener(localAddr, TcpListerPort);
                logger.Log("TcpListener created");

                // Start listening for client requests.
                logger.Log("Waiting to start TcpListener ... ");
                server.Start();
                logger.Log("TcpListener started");
            }
            catch (SocketException ex)
            {
                ErrorCode = ex.ErrorCode;
                SocketErrorCode = ex.SocketErrorCode;

                string errorText = String.Format("WaitForPassword: TcpListener Start Failed\nErrorCode: {0} ({1})\nMessage: {2}\n",
                    Enum.GetName(typeof(SocketError), SocketErrorCode), SocketErrorCode, ex.Message);
                bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, errorText);

                logger.Log("SocketException: ", ex.ToString());
                exceptionFired = true;
                //System.Environment.Exit(1);  // exit if finally
            }
            finally
            {
                // Stop listening for new clients.
                logger.Log("WaitForPassword: Executing finally", exceptionFired.ToString());
                //server.Stop();

                if (exceptionFired)
                {
                    logger.Log("WaitForPassword: calling exit(XX)");
                    if (ErrorCode == (int)SocketError.AddressAlreadyInUse)
                        System.Environment.Exit(ExitCodes.ExitTcpPort);
                    else
                        System.Environment.Exit(ExitCodes.ExitTcpServer);
                }
            }
            try
            {
                // Buffer for reading data
                Byte[] bytes = new Byte[1024];
                String data = null;

                logger.Log("Waiting for a connection... ");

                // Perform a blocking call to accept requests.
                // You could also use server.AcceptSocket() here.

                TcpClient client = server.AcceptTcpClient();
                logger.Log("Connected!");

                data = null;

                // Get a stream object for reading and writing
                NetworkStream stream = client.GetStream();

                logger.Log("Reading");
                int i;
                while ((i = stream.Read(bytes, 0, bytes.Length)) != 0)
                {
                    // Translate data bytes to a ASCII string.
                    data = System.Text.Encoding.ASCII.GetString(bytes, 0, i);
                    password = data;
                    logger.Log("Received: ", data);

                    // Process the data sent by the client.
                    byte[] msg = System.Text.Encoding.ASCII.GetBytes(data);

                    // Send back a response.
                    //stream.Write(null, 0, msg.Length);
                    stream.Write(msg, 0, msg.Length);
                    logger.Log("Sent: ", data);
                    break;
                }

                // Shutdown and end connection
                client.Close();
            }
            catch (Exception ex)
            {
                bool rval = FileUtils.CreateWriteCloseFile(errorFilePath, "WaitForPassword Failed\n", ex.Message);
                logger.Log("SocketException: ", ex.ToString());
                exceptionFired = true;
                //System.Environment.Exit(1);  // exit if finally
            }
            finally
            {
                // Stop listening for new clients.
                logger.Log("WaitForPassword: Executing finally", exceptionFired.ToString());
                server.Stop();

                if (exceptionFired)
                {
                    logger.Log("WaitForPassword: calling exit(1)");
                    System.Environment.Exit(ExitCodes.ExitTcpServer);
                }
            }
            return password;
        }
        public static string GetLocalAppDataPath()
        {
            string tempPath = Path.GetTempPath();
            string tempPathAsFilePath = tempPath.TrimEnd('\\');
            string dataPath = Path.GetDirectoryName(tempPathAsFilePath);
            //string datapath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return dataPath;
        }
        public static string FindKeyinArgs(string[] args, string keyword)
        {
            int numArgs = args.GetLength(0);

            for (int j = 0; j < numArgs; j++)
            {
                if ((j + 1) >= numArgs)
                {
                    return "";
                }

                string key = args[j];
                string val = args[j + 1];

                if (key.CompareTo(keyword) == 0)
                {
                    return val;
                }
            }
            return "";
        }

        public static void PrintCapabilities(SmtpClient client, FileLogger logger)
        {
            if (client.Capabilities.HasFlag(SmtpCapabilities.Authentication))
            {
                var mechanisms = string.Join(", ", client.AuthenticationMechanisms);
                logger.Log("The SMTP server supports the following SASL mechanisms: ", mechanisms);
            }

            if (client.Capabilities.HasFlag(SmtpCapabilities.Size))
                logger.Log("The SMTP server has a size restriction on messages: .", client.MaxSize.ToString());

            if (client.Capabilities.HasFlag(SmtpCapabilities.Dsn))
                logger.Log("The SMTP server supports delivery-status notifications.");

            if (client.Capabilities.HasFlag(SmtpCapabilities.EightBitMime))
                logger.Log("The SMTP server supports Content-Transfer-Encoding: 8bit");

            if (client.Capabilities.HasFlag(SmtpCapabilities.BinaryMime))
                logger.Log("The SMTP server supports Content-Transfer-Encoding: binary");

            if (client.Capabilities.HasFlag(SmtpCapabilities.UTF8))
                logger.Log("The SMTP server supports UTF-8 in message headers.");
        }
    }
}


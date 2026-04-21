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

// This is my second c# code. Need better Exception handling, etc everything needs more work -:)

using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.Encodings;
using MimeKit.IO;
using MimeKit.Text;
using MimeKit.Tnef;
using MimeKit.Utils;
using Org.BouncyCastle.Asn1.X509;

using RTF2HTMLConversion;
using RtfPipe;
using RtfPipe.Model;
using RtfPipe.Tokens;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Reflection.Metadata;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Web;
using static Org.BouncyCastle.Crypto.Engines.SM2Engine;
using static System.Net.Mime.MediaTypeNames;
using static System.Net.WebRequestMethods;
using static System.Runtime.InteropServices.JavaScript.JSType;
using static Winmail2EmlFile.Program;

#pragma warning disable 219

/*
		MessageBox not supported unless project type is changed
		"<TargetFramework>net8.0</TargetFramework> to <TargetFramework>net8.0-windows</TargetFramework>, and add <UseWindowsForms>true</UseWindowsForms>"
using System.Windows.Forms;
*/

namespace Winmail2EmlFile
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

	class Program
	{
		// Handle division by zero, and other unhandled exceptions
		static void MyHandler(object sender, UnhandledExceptionEventArgs args)
		{
			Exception e = (Exception)args.ExceptionObject;
			string ex = e.Message;
			//Console.WriteLine("MyHandler caught : " + e.Message);
			//Console.WriteLine("Runtime terminating: {0}", args.IsTerminating);
			// Server socket may linger ??
#if C
			MessageBox not supported unless project type is changed
			"<TargetFramework>net8.0</TargetFramework> to <TargetFramework>net8.0-windows</TargetFramework>, and add <UseWindowsForms>true</UseWindowsForms>"
			string title = "Execution Failure Info";
			MessageBoxButtons buttons = MessageBoxButtons.Ok;
			MessageBoxIcon icon = MessageBoxIcon.Error;
			DialogResult result = MessageBox.Show(ex, title, buttons, icon);
#endif
			System.Environment.Exit(ExitCodes.ExitUnhandledException);
		}

		static void ErrorExit(int ExitCode, string errorText)
		{
			Console.Beep();
			System.Environment.Exit(ExitCode);
		}

		static void ErrorExit(ref string title, int exitCode, ref Exception ex, ref string targetHtmlLanguageCode, int beepCnt = 1)
		{
			for (int i = 0; i < beepCnt; i++)
				Console.Beep();

			string exText = ex.ToString();
			string errorText = title + " Exit Code: " + exitCode.ToString() + "\n" + exText;
			string contextText = title + " Exit Code: " + exitCode.ToString() + "\n";
			//logger.Log(errorText);

			bool htmlOk = CreateTranslationHtml(ex, contextText, targetHtmlLanguageCode);

			System.Environment.Exit(exitCode);
		}

		public static class ExitCodes
		{
			public const int ExitOk = 0;
			public const int ExitCmdArguments = -11;
			public const int ExitMimeMessage = -12;
			public const int ExitRtf2HtmlFailure = -13;
			public const int ExitWriteUnmodifiedEmlFileFailure = -14;
			public const int ExitWriteEmlFileFailure = -15;
			public const int ExitWinmailProcessingFailure = -16;
			public const int ExitUnhandledException = -17;
		}

		public static class RtfToHtmlConverter
		{
			public static string ConvertRtfToHtml(ref string rtf)
			{
				if (string.IsNullOrEmpty(rtf))
					return string.Empty;

				var html = RtfPipe.Rtf.ToHtml(rtf.Trim('\0'));
				return html;
			}

			public static string ConvertRtfPathToHtml(ref string path)
			{
				using (StreamReader stream = new StreamReader(path, Encoding.UTF8))
				{
					var actual = Rtf.ToHtml(stream);

					return actual;
				}
			}
		}

		private static int _countDown = 30;

		public static void TimerFired(object state)
		{
			// Will do all cleanup including timer thread ??
			System.Environment.Exit(0);
		}

		public static void OnTimedEvent(object source, ElapsedEventArgs e)
		{
			
			if (_countDown >= 10)
				Console.Write("{0}\b\b", _countDown);
			else
				Console.Write("{0} \b\b", _countDown);

			_countDown--;
			if (_countDown < 0)
				// Will  do all cleanup including timer thread ??
				System.Environment.Exit(0);
		}

		public static void Usage()
		{
			Console.WriteLine("\n");
			Console.WriteLine("Usage:\n");
			Console.WriteLine("Winmai2Eml.exe [--target-html-language-code en|es|it|pt|pt-PT|pt-BR|fr|de|pl|ro|ja|ru|zh-CN|ar|uk|hu] winmail.dat");
			Console.Write("\nEnter Return to Exit otherwise application will be closed automatically after {0} seconds: ", _countDown);

			// Both solutions work
#if C
			var timer = new System.Threading.Timer(TimerFired, null, 10000, 0);

			Console.ReadLine();
			timer.Dispose();
#else
			// Create a timer with a 1 second interval.
			var aTimer = new System.Timers.Timer(1000);
			// Hook up the Elapsed event for the timer. 
			aTimer.Elapsed += OnTimedEvent;
			aTimer.AutoReset = true;
			aTimer.Enabled = true;

			Console.ReadLine();
			aTimer.Close();
			aTimer.Dispose();
#endif

			System.Environment.Exit(0);
		}

		public static void Main(string[] args)
		{
			int numArgs = args.GetLength(0);

			if (numArgs == 0)
				Usage();

			AppDomain currentDomain = AppDomain.CurrentDomain;
			currentDomain.UnhandledException += new UnhandledExceptionEventHandler(MyHandler);

			string dataPath = GetLocalAppDataPath();
			string MBoxViewerPath = Path.Combine(dataPath, "UMBoxViewer");
			string MailServicePath = Path.Combine(MBoxViewerPath, "MailService");
			string TempPath = Path.Combine(MailServicePath, "Temp");

			System.IO.DirectoryInfo info = Directory.CreateDirectory(TempPath);

			string loggerFilePath = FindKeyinArgs(args, "--logger-file");
			var logger = new FileLogger();

			logger.Open(loggerFilePath);
			logger.Log("Logger Open");

#if DEBUG
			//System.Diagnostics.Debugger.Launch();
			//System.Diagnostics.Debugger.Break();

			logger.Log("Debugger.Launch()");
#endif

			string inputWinmailFilePath = "";
			string outputEmlFilePath = "";

			// Development team options
			string mboxviewExe = "";
			string inputRtfFilePath = "";
			string outputRtf2HtmlFilePath = "";
			string outputRtfPipe2HtmlFilePath = "";
			string outputUnmodifiedEmlFilePath = "";
			string targetHtmlLanguageCode = "en";

			string Params = string.Join(" ", args);

			logger.Log("Command line argument list:");
			for (int j = 0; j < numArgs; j++)
			{
				string key = args[j];
				string val = "";
				if ((j + 1) < numArgs)
				{
					val = args[j + 1];
					j++;
				}
				else
				{
					inputWinmailFilePath = key;
					break;
				}

				if (!key.StartsWith("--"))
				{
					string errorText = System.String.Format("Invalid key: {0} ", key);
					logger.Log(errorText);
					ErrorExit(ExitCodes.ExitCmdArguments, errorText);
				}

				if (key.CompareTo("--mboxview-exe-path") == 0)
				{
					mboxviewExe = val;
				}
				else if (key.CompareTo("--input-winmail-file") == 0)
				{
					inputWinmailFilePath = val;
				}
				else if (key.CompareTo("--output-eml-file") == 0)
				{
					outputEmlFilePath = val;
				}
				else if (key.CompareTo("--output-unmodified-eml-file") == 0)
				{
					outputUnmodifiedEmlFilePath = val;
				}
				else if (key.CompareTo("--input-rtf-file") == 0)
				{
					inputRtfFilePath = val;
				}
				else if (key.CompareTo("--output-rtf-html-file") == 0)
				{
					outputRtf2HtmlFilePath = val;
				}
				else if (key.CompareTo("--output-rtfpipe-html-file") == 0)
				{
					outputRtfPipe2HtmlFilePath = val;
				}
				else if (key.CompareTo("--logger-file") == 0)
				{
					; // see FindKeyinArgs(args, "--logger-file");
				}
				else if (key.CompareTo("--target-html-language-code") == 0)
				{
					targetHtmlLanguageCode = val;
				}
				else
				{
					logger.Log(System.String.Format("    Unknown Key: {0} {1}", args[j], args[j + 1]));
				}
				logger.Log(System.String.Format("    {0} {1}", key, val));
			}

			// Register the CodePages encoding provider
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

			if (!System.IO.File.Exists(inputWinmailFilePath) && !System.IO.File.Exists(inputRtfFilePath))
			{
				string errorText = System.String.Format("Input file {0} doesn't exist.", inputWinmailFilePath);
				logger.Log(errorText);

				ErrorExit(ExitCodes.ExitCmdArguments, errorText);
			}

			// to support .rtf files
			if (inputRtfFilePath.Length == 0)
			{
				string ext = Path.GetExtension(inputWinmailFilePath);
				if (ext == ".rtf")
				{
					inputRtfFilePath = inputWinmailFilePath;
				}
			}

			if ((inputRtfFilePath.Length > 0) && System.IO.File.Exists(inputRtfFilePath))
			{
				try
				{
					string? rtfText = null;
					if (!IsValidRtfFile(inputRtfFilePath))
					{
						string errorText = "";
						byte[] rtfSrc = System.IO.File.ReadAllBytes(inputRtfFilePath);
						if (IsValidCompressedRtfFile(rtfSrc, ref errorText))
						{
							byte[] byteRtf = MsgReader.Outlook.RtfDecompressor.DecompressRtf(rtfSrc);
							rtfText = Encoding.ASCII.GetString(byteRtf);
						}
						else
							throw new Exception("Invalid Rtf File: expected \"{\\rtf\" at the beginning of the file");
					}
					else
						rtfText = System.IO.File.ReadAllText(inputRtfFilePath);

					//byte[] bytesArr = GetBytes(rtfText);

					bool isHtmlText = rtfText.Contains("fromhtml1");
					Encoding encodingUTF8 = Encoding.UTF8;

					string result;
					if (isHtmlText)
					{
						RTF2HTMLConverter converter = new();
						StringBuilder resultSB = new();

						string res = converter.rtf2html(ref rtfText, ref resultSB);
						result = resultSB.ToString();

						string outRtf2HtmlFilePath = outputRtf2HtmlFilePath;
						if (outputRtf2HtmlFilePath.Length <= 0)
						{
							outRtf2HtmlFilePath = inputRtfFilePath + ".Rtf2Html.html";
						}

						System.IO.File.WriteAllText(outRtf2HtmlFilePath, result, encodingUTF8);
					}

					//const string rtfInlineObject = "[*[RTFINLINEOBJECT]*]";
					//rtfText = rtfText.Replace("\\objattph", rtfInlineObject);

					result = RtfToHtmlConverter.ConvertRtfToHtml(ref rtfText);

					string outRtfPipe2HtmlFilePath = outputRtfPipe2HtmlFilePath;
					if (outRtfPipe2HtmlFilePath.Length <= 0)
					{
						if (isHtmlText)
							outRtfPipe2HtmlFilePath = inputRtfFilePath + ".RtfPipe.html";
						else
							outRtfPipe2HtmlFilePath += inputRtfFilePath + ".RtfPipe.text.html";

					}
					System.IO.File.WriteAllText(outRtfPipe2HtmlFilePath, result, encodingUTF8);

					Process.Start(new ProcessStartInfo(outRtfPipe2HtmlFilePath) { UseShellExecute = true });

					logger.Log("Processing of {0} file completed.", inputRtfFilePath);

					System.Environment.Exit(ExitCodes.ExitOk);;
				}
				catch (Exception ex)
				{
					string title = System.String.Format("Processing of {0} file failed:", inputRtfFilePath);
					ErrorExit(ref title, ExitCodes.ExitRtf2HtmlFailure, ref ex, ref targetHtmlLanguageCode);
				}
			}

			if ((inputWinmailFilePath.Length > 0) && System.IO.File.Exists(inputWinmailFilePath))
			{
				try
				{
					if (!IsValidTnefSignature(inputWinmailFilePath))
					{
							throw new Exception("Invalid Tnef File ID Signature");
					}

					using (var inputStream = new FileInfo(inputWinmailFilePath).OpenRead())
					{
						var tnef = new TnefPart { Content = new MimeKit.MimeContent(inputStream) };

						MimeMessage message = tnef.ConvertToMessage();

						var from = message.Headers["from"];
						if ((from != null) && (from.Length == 0))
						{
							message.Headers["from"] = "unknown@unknown.com";
						}

						if (outputUnmodifiedEmlFilePath.Length > 0)
						{
							try
							{
								message.WriteTo(outputUnmodifiedEmlFilePath);
							}
							catch (Exception ex)
							{
								string title = System.String.Format("Write to {0} file failed:", outputUnmodifiedEmlFilePath);
								ErrorExit(ref title, ExitCodes.ExitWriteUnmodifiedEmlFileFailure, ref ex, ref targetHtmlLanguageCode);
							}
						}

						//string msgstr = message.ToString(); // debug

						// Check if "text/html" and "text/rtf" parts exist
						bool htmlPartExists = false;
						bool rtfPartExists = false;
						foreach (var bodyPart in message.BodyParts)
						{
							if (!bodyPart.IsAttachment)
							{
								var part = (MimePart)bodyPart;
								if ((part.ContentType.MediaType == "text") && (part.ContentType.MediaSubtype == "html"))
								{
									htmlPartExists = true;
								}
								if ((part.ContentType.MediaType == "text") && (part.ContentType.MediaSubtype == "rtf"))
								{
									rtfPartExists = true;
								}
							}
						}

						foreach (var bodyPart in message.BodyParts)
						{
							bool done = false;
							if (!bodyPart.IsAttachment)
							{
								var part = (MimePart)bodyPart;
								//string partstr = part.ToString();  // debug

								if ((part.ContentType.MediaType == "text") && ((part.ContentType.MediaSubtype == "plain") || (part.ContentType.MediaSubtype == "html")))
								{
									var ContentTransferEncoding = part.ContentTransferEncoding;
									if (part.ContentTransferEncoding == ContentEncoding.Default)
									{
										part.ContentTransferEncoding = ContentEncoding.QuotedPrintable;
									}
								}
								else if ((part.ContentType.MediaType == "text") && (part.ContentType.MediaSubtype == "rtf"))
								{
									var ContentTransferEncoding = part.ContentTransferEncoding;

									RTF2HTMLConverter converter = new();
									StringBuilder resultSB = new();
									string result = "";
									bool isHtmlText = false;

									try
									{
										using (var memory = new MemoryStream())
										{
											bool contentOnly = true;
											part.Content.DecodeTo(memory);

											var buffer = memory.ToArray();
											var text = Encoding.UTF8.GetString(buffer);

											isHtmlText = text.Contains("fromhtml1");

											string res;
											// RtfToHtmlConverter.ConvertRtfToHtml(ref text); supports both RTF Text and RTF HML encoding
											// Need to test if converter.rtf2html(ref text, ref resultSB); can be commented out
											if (isHtmlText)
											{
												res = converter.rtf2html(ref text, ref resultSB);
												result = resultSB.ToString();
											}
											else
												result = RtfToHtmlConverter.ConvertRtfToHtml(ref text);

											var textPart = (TextPart)part;
											textPart.Text = result;
											part.ContentType.MediaSubtype = "html";
											part.ContentTransferEncoding = ContentEncoding.QuotedPrintable;
										}
									}
									catch (Exception ex)
									{
										string title = System.String.Format("Conversion from Rtf to Html failed:");
										ErrorExit(ref title, ExitCodes.ExitRtf2HtmlFailure, ref ex, ref targetHtmlLanguageCode);
									}

									done = true;
								}
							}
						}

						string emlFilePath = outputEmlFilePath;
						try
						{
							if (outputEmlFilePath.Length == 0)
							{
								emlFilePath = inputWinmailFilePath + ".eml";
							}
							message.WriteTo(emlFilePath);
						}
						catch (Exception ex)
						{
							string title = System.String.Format("Write to {0} file failed:", emlFilePath);
							ErrorExit(ref title, ExitCodes.ExitWriteEmlFileFailure, ref ex, ref targetHtmlLanguageCode);
						}

						//  Code to update names
#if C
						try
						{
							string rtfText = System.IO.File.ReadAllText(emlFilePath);

							byte[] rawBytesArr = System.IO.File.ReadAllBytes(emlFilePath);
							byte[] bytesArr = GetBytes(rtfText);

							string fromStr = "zzzzzzz";
							string toStr = "inspector";
							string correctedString = Regex.Replace(rtfText, fromStr, toStr, RegexOptions.IgnoreCase);
							if (correctedString.Count() <= 0)
								Console.Beep();
							System.IO.File.WriteAllText(emlFilePath, correctedString);
						}
						catch (Exception ex)
						{
							string title = "Write to file failed:";
							ErrorExit(ref title, ExitCodes.ExitWriteEmlFileFailure, ref ex, ref targetHtmlLanguageCode);
						}
#endif
						int deb = 1;
					}
				}
				catch (Exception ex)
				{
					string title = System.String.Format("Processing of {0} file failed:", inputWinmailFilePath);
					ErrorExit(ref title, ExitCodes.ExitWinmailProcessingFailure, ref ex, ref targetHtmlLanguageCode);
				}
				if (outputEmlFilePath.Length > 0)
				{
					string cargs = outputEmlFilePath;
					Process.Start(mboxviewExe, cargs);
				}

				logger.Log("Processing of {0} file completed.", inputWinmailFilePath);
			}

			System.Environment.Exit(ExitCodes.ExitOk);
		}

		public static string GetLocalAppDataPath()
		{
			string tempPath = Path.GetTempPath();
			string tempPathAsFilePath = tempPath.TrimEnd('\\');
			string? dataPath = Path.GetDirectoryName(tempPathAsFilePath);
			//string datapath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
			if (dataPath == null)
				dataPath = "";
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

		public static byte[] GetBytes(string str)
		{
			byte[] bytes = new byte[str.Length * sizeof(char)];
			System.Buffer.BlockCopy(str.ToCharArray(), 0, bytes, 0, bytes.Length);
			return bytes;
		}

		public static bool IsValidTnefSignature(string filePath)
		{
			byte[] data = new byte[4];
			using (BinaryReader reader = new BinaryReader(new FileStream(filePath, FileMode.Open)))
			{
				//reader.BaseStream.Seek(0, SeekOrigin.Begin);
				reader.Read(data, 0, 4);
				reader.Close();
				reader.Dispose();
			}
			byte signature0 = data[0];
			byte signature1 = data[1];
			byte signature2 = data[2];
			byte signature3 = data[3];
			if ((signature0 != 0x78) ||
				(signature1 != 0x9F) ||
				(signature2 != 0x3E) ||
				(signature3 != 0x22))
			{
				return false;
			}
			else
				return true;
		}

		public static bool IsValidRtfFile(string filePath)
		{
			byte[] data = new byte[5];
			using (BinaryReader reader = new BinaryReader(new FileStream(filePath, FileMode.Open)))
			{
				//reader.BaseStream.Seek(0, SeekOrigin.Begin);
				reader.Read(data, 0, 5);
				reader.Close();
				reader.Dispose();
			}
			if ((data[0] == '{') &&
				(data[1] == '\\') &&
				((data[2] == 'r') || (data[2] == 'R') ) &&
				((data[3] == 't') || (data[3] == 'T') ) &&
				((data[4] == 'f') || (data[4] == 'F'))
				)
			{
				return true;
			}

			return false;
		}

		public static bool IsValidCompressedRtfFile(byte[] rtfSrc, ref string errorText)
		{
			bool ret = MsgReader.Outlook.RtfDecompressor.IsValidCompressedRtf(rtfSrc, ref errorText);
			return ret;
		}

		public static bool CreateTranslationHtml(Exception ex, string contextText, string targetHtmlLanguageCode)
		{
			// en to en translation would not work, removed
			// function googleTranslateElementInit()
			// new google.translate.TranslateElement({ pageLanguage: 'en',

			string outputHtmlFile;

			string messageText = contextText + "\r\n" + ex.Message;

			string htmlEncodedText = HttpUtility.HtmlEncode(messageText);

			string htmlEncodedExceptionText = "<pre>\r\n" + htmlEncodedText + "\r\n</pre>";

			string htmlEncodeStackTrace = HttpUtility.HtmlEncode(ex.StackTrace);
			string htmlEncodeStackTraceText = "<pre>\r\n" + htmlEncodeStackTrace + "\r\n</pre>";

			int d = 1;

			string htmlDocHdrPlusStyle = """
<!DOCTYPE html>
<html lang="en-US">
<head>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8">

<script type="text/javascript">

function googleTranslateElementInit()
{
	new google.translate.TranslateElement({
	includedLanguages: 'en,es,it,pt,pt-PT,pt-BR,fr,de,pl,ro,ja,ru,zh-CN,ar,uk,hu'},
    'google_translate_element');
}

function myTranslate(lang)
{
	var selectElement = document.querySelector('#google_translate_element select');
	selectElement.value = lang;
	selectElement.dispatchEvent(new Event('change'));
}

</script>

<script type="text/javascript" src="https://translate.google.com/translate_a/element.js?cb=googleTranslateElementInit"></script>
</head>

<style>

.goog-te-gadget {
	font-size: 19px !important;
}

body > .skiptranslate {
	//display: none;
}

.goog-te-banner-frame.skiptranslate {
	display: none !important;
}

body {
top: 48px !important;
	font-size: 18px !important;
}

@media print {
# google_translate_element {display: none;}
}
</style>
""";


			string htmlBodyDivTranslate = """
<!--Text that WILL be translated -->
<div id="google_translate_element" ></div>
""";

			string htmlBodyAllsDone = """		
<br>
</body>
</html>
""";

			string htmlBodyDicNoTranslate = """
<!--Text that will NOT be translated-->
<div class="notranslate">

<p>
""";

			htmlBodyDicNoTranslate += htmlEncodeStackTraceText;

			htmlBodyDicNoTranslate += """
</p>

</div>
""";

			string CRLF = "\r\n\r\n"; ;

			try
			{
				using (var stream = new MemoryStream())
				{
					using (var writer = new StreamWriter(stream, Encoding.UTF8))
					{
						writer.Write(htmlDocHdrPlusStyle);
						writer.Write(CRLF);

						string bodyTranslate = System.String.Format("""<body onload="setTimeout(myTranslate, 1000, '{0}')">""", targetHtmlLanguageCode);

						writer.Write(bodyTranslate);
						writer.Write(CRLF);

						writer.Write(htmlBodyDivTranslate);
						writer.Write(CRLF);

						string titleFile = """			
<br><pre>---  Winmail2Eml Runtime Failure Exception Details  ---</pre><br>
""";

						writer.Write(titleFile);
						writer.Write(CRLF);

						writer.Write(htmlEncodedExceptionText);
						writer.Write(CRLF);

						writer.Write(htmlBodyDicNoTranslate);
						writer.Write(CRLF);

						writer.Write(htmlBodyAllsDone);
						writer.Write(CRLF);

						writer.Flush();

						var buffer = stream.ToArray();
						string text = Encoding.UTF8.GetString(buffer);

						outputHtmlFile = GetLocalAppDataPath();

						outputHtmlFile += "\\Temp\\UMBoxViewer\\MboxHelp";

						if (!Directory.Exists(outputHtmlFile))
						{
							// Create the directory
							Directory.CreateDirectory(outputHtmlFile);
						}

						outputHtmlFile += "\\exceptionTextWinmail2Eml.html";

						System.IO.File.WriteAllText(outputHtmlFile, text);

						Process.Start(new ProcessStartInfo(outputHtmlFile) { UseShellExecute = true });
					}
				}
			}
			catch (Exception except)
			{
				string text = except.ToString();
				; // ignore
			}
			return true;
		}
	}
}


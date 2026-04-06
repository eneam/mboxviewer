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
///using System.Text.Encoding.CodePages;
///
using RTF2HTMLConversion;
using RtfPipe;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

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

		public static void Main(string[] args)
		{
			int deb = 1;

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

			int numArgs = args.GetLength(0);

			string inputWinmailFilePath = "";
			string outputEmlFilePath = "";

			// Development team options
			string mboxviewExe = "";
			string inputRtfFilePath = "";
			string outputRtf2HtmlFilePath = "";
			string outputRtfPipe2HtmlFilePath = "";
			string outputUnmodifiedEmlFilePath = "";

			string Params = string.Join(" ", args);

			logger.Log("Command line argument list:");
			for (int j = 0;  j < numArgs; j++)
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
					string errorText = String.Format("Invalid key: {0} ", key);
					logger.Log(errorText);
					System.Environment.Exit(ExitCodes.ExitCmdArguments);
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
				else
				{
					logger.Log(String.Format("    Unknown Key: {0} {1}", args[j], args[j + 1]));
				}
				logger.Log(String.Format("    {0} {1}", key, val));
			}

			// Register the CodePages encoding provider
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

			if (!File.Exists(inputWinmailFilePath) && !File.Exists(inputRtfFilePath))
			{
				string errorText = String.Format("Input file {0} doesn't exist.", inputWinmailFilePath);
				logger.Log(errorText);
				System.Environment.Exit(ExitCodes.ExitCmdArguments);
			}

			// Used by development team
			if ((inputRtfFilePath.Length > 0) && File.Exists(inputRtfFilePath))
			{
				try
				{
					StreamReader reader = new StreamReader(inputRtfFilePath, Encoding.UTF8);
					string rtfText = reader.ReadToEnd();
					byte[] bytesArr = GetBytes(rtfText);

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

						File.WriteAllText(outRtf2HtmlFilePath, result, encodingUTF8);
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
					File.WriteAllText(outRtfPipe2HtmlFilePath, result, encodingUTF8);

					deb = 1;
				}
				catch (Exception ex)
				{
					string exText = ex.ToString();
					logger.Log("Processing rtf file failed: ", exText);

					System.Environment.Exit(ExitCodes.ExitRtf2HtmlFailure);
				}
			}

			if ((inputWinmailFilePath.Length > 0) && File.Exists(inputWinmailFilePath))
			{
				try
				{
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
								string exText = ex.ToString();
								logger.Log("Write to file failed: ", exText);

								System.Environment.Exit(ExitCodes.ExitWriteUnmodifiedEmlFileFailure);
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
										string exText = ex.ToString();
										logger.Log("Write to file failed: ", exText);

										System.Environment.Exit(ExitCodes.ExitRtf2HtmlFailure);
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
							string exText = ex.ToString();
							logger.Log("Write to file failed: ", exText);

							System.Environment.Exit(ExitCodes.ExitWriteEmlFileFailure);
						}

						deb = 1;
					}
				}
				catch (Exception ex)
				{
					string exstr = ex.ToString();
					logger.Log("Processing of winmail.dat file failed: ", ex.ToString());
					System.Environment.Exit(ExitCodes.ExitWinmailProcessingFailure);
				}
				if (outputEmlFilePath.Length > 0)
				{
					string cargs = outputEmlFilePath;
					Process.Start(mboxviewExe, cargs);
				}

				logger.Log("Processing of winmail.dat file completed.");
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

	}
}


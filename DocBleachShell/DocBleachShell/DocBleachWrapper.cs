// License: MIT
// Copyright: Joe Security
// Dependencies: - DocBleach https://github.com/docbleach
//				 - Log4Net https://logging.apache.org/log4net/
//				 - Ntfs Streams https://github.com/RichardD2/NTFS-Streams

using System;
using System.Diagnostics;
using System.Configuration;
using System.IO;
using log4net;

namespace DocBleachShell
{
	/// <summary>
	/// Wrapper around DocBleach.
	/// </summary>
	public class DocBleachWrapper
	{
		private static readonly ILog Logger = LogManager.GetLogger(typeof(DocBleachWrapper));
		
		/// <summary>
		/// Bleach the document.
		/// </summary>
		/// <param name="FilePath"></param>
		/// <param name="TargetDirectory"></param>
		public void Bleach(String FilePath, String TargetDirectory)
		{
			Logger.Debug("Try to bleach: " + FilePath);
			
			String CurrentDir = Directory.GetCurrentDirectory();
					
			Directory.SetCurrentDirectory(Directory.GetParent(FilePath).FullName);
		
			try
			{
				File.Copy(FilePath, TargetDirectory + "\\backup\\docs\\" + Path.GetFileName(FilePath), true);
			} catch(Exception e)
			{
				Logger.Error("Unable to make a backup of the document", e);
			}
			
			// Original doc -> bleach. .... do
			
			String TmpDoc = Path.GetDirectoryName(FilePath) + "\\bleach." + Path.GetFileName(FilePath);
			
			try
			{
				File.Move(FilePath, TmpDoc);
			} catch(Exception e)
			{
				Logger.Error("Unable to rename document: " + FilePath, e);
			}

			// Call docbleach which will generate doc
			String Output = "initialized";
			String bleachBin = "\"" + ConfigurationManager.AppSettings["PathToDocBleach"] + "\"";
			if (File.Exists(ConfigurationManager.AppSettings["PathToDocBleach"]))
			{
				Process p = new Process();
				p.StartInfo.FileName = ConfigurationManager.AppSettings["PathToDocBleach"];
				p.StartInfo.Arguments = "-in \"" + TmpDoc + "\" -out \"" + Path.GetFileName(FilePath) + "\"";
				p.StartInfo.UseShellExecute = false;
				p.StartInfo.CreateNoWindow = true;
				p.StartInfo.RedirectStandardOutput = true;
				p.Start();

				Output = p.StandardOutput.ReadToEnd();
				p.WaitForExit();

			}
			else
			{
				Logger.Error("Cant find DocBleach at: " + ConfigurationManager.AppSettings["PathToDocBleach"]);
			}
			
			Logger.Debug("DocBleach output: " + Output);

			// If the document was bleach "contains potential malicious elements" analyze the file with Joe Sandbox Cloud.
			// If no Cloud API key has been configured, nothing is send to the cloud
			String APIKey = ConfigurationManager.AppSettings["JoeSandboxCloudAPIKey"];
			if (APIKey.Length != 0)
			{
				if (!Output.Contains("file was already safe"))
				{
					new JoeSandboxClient().Analyze(TmpDoc, APIKey);
					Logger.Debug("Doc sent to the cloud");
				}
			}
			else
			{
				Logger.Debug("Doc not sent to cloud : no API key configured");
			}

			// Cleanup & recovery
			if (File.Exists(FilePath))
			{
				try
				{
					File.Delete(TmpDoc);
				} catch(Exception e)
				{
					Logger.Error("Unable to delete original file: " + TmpDoc, e);
				}
				
				Logger.Debug("Successfully bleached: " + FilePath);
			} else
			{
				Logger.Debug("Unable to bleach: " + FilePath + ", file not found");
				
				// No doc, move back.
				try
				{
					File.Move(TmpDoc, FilePath);
					
				} catch(Exception e)
				{
					Logger.Error("Unable to rename document: " + TmpDoc, e);
				}
			}
			
			Directory.SetCurrentDirectory(CurrentDir);

		}
	}
}

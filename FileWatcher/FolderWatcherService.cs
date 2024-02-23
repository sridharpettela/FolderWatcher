using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace FileWatcher
{
	public partial class FolderWatcherService : ServiceBase
	{
		private FileSystemWatcher _watcher;
		private string _folderToWatch = ConfigurationSettings.AppSettings["WatchFolderPath"];
		private string _powershellScriptPath = ConfigurationSettings.AppSettings["PowershellScriptPath"];
		private string _fileExtension = ConfigurationSettings.AppSettings["FileExtension"];
		public FolderWatcherService()
		{
			InitializeComponent();
		}

		protected override void OnStart(string[] args)
		{
			InitializeWatcher();
		}

		protected override void OnStop()
		{
			_watcher.EnableRaisingEvents = false;
			_watcher.Dispose();
		}

		private void InitializeWatcher()
		{
			_watcher = new FileSystemWatcher();
			_watcher.Path = _folderToWatch;
			_watcher.NotifyFilter = NotifyFilters.FileName;
			_watcher.Filter = _fileExtension;
			_watcher.Created += OnFileCreated;
			_watcher.EnableRaisingEvents = true;
		}

		private void OnFileCreated(object sender, FileSystemEventArgs e)
		{
			// Call PowerShell script passing the newly created file path
			string filePath = e.FullPath;
			RunPowerShellScript(filePath);
		}

		private void RunPowerShellScript(string filePath)
		{
			string command = $"powershell.exe -ExecutionPolicy Bypass -File \"{_powershellScriptPath}\" \"{filePath}\"";
			System.Diagnostics.ProcessStartInfo processInfo = new System.Diagnostics.ProcessStartInfo("cmd.exe", "/c " + command);
			processInfo.CreateNoWindow = true;
			processInfo.UseShellExecute = false;
			System.Diagnostics.Process process = System.Diagnostics.Process.Start(processInfo);
			process.WaitForExit();
		}
	}
}

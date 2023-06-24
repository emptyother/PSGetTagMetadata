using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using Microsoft.PowerShell.Commands;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace OnlyHuman.PSGetTagMetadata
{
	[Cmdlet(VerbsCommon.Set, "Shortcut", DefaultParameterSetName = ParamSetPath, SupportsShouldProcess = true)]
	public class Set_Shortcut : PSCmdlet
	{
		private const string Noun = "FileMetadata";

		private const string ParamSetLiteral = "Literal";

		private const string ParamSetPath = "Path";

		private string[] _paths;

		private bool _shouldExpandWildcards;

		[Parameter(
			Mandatory = true,
			ValueFromPipeline = false,
			ValueFromPipelineByPropertyName = true,
			ParameterSetName = ParamSetLiteral)
		]
		[Alias("PSPath")]
		[ValidateNotNullOrEmpty]
		public string[] LiteralPath
		{
			get { return _paths; }
			set { _paths = value; }
		}

		[Parameter(
			Position = 0,
			Mandatory = true,
			ValueFromPipeline = true,
			ValueFromPipelineByPropertyName = true,
			ParameterSetName = ParamSetPath)
		]
		[ValidateNotNullOrEmpty]
		public string[] Path
		{
			get { return _paths; }
			set
			{
				_shouldExpandWildcards = true;
				_paths = value;
			}
		}

		[Parameter(
			Position = 1,
			Mandatory = false,
			ValueFromPipeline = false,
			ValueFromPipelineByPropertyName = true
		)]
		public string OutputPath { get; set; } = Directory.GetCurrentDirectory();

		protected override void BeginProcessing()
		{
			if (!Directory.Exists(OutputPath))
			{
				throw new DirectoryNotFoundException($"Directory {OutputPath} does not exist.");
			}
		}

		protected override void ProcessRecord()
		{
			foreach (string path in _paths)
			{
				// This will hold information about the provider containing
				// the items that this path string might resolve to.
				ProviderInfo provider;
				// This will be used by the method that processes literal paths
				PSDriveInfo drive;
				// this contains the paths to process for this iteration of the
				// loop to resolve and optionally expand wildcards.
				var filePaths = new List<string>();
				if (_shouldExpandWildcards)
				{
					// Turn *.txt into foo.txt,foo2.txt etc.
					// if path is just "foo.txt," it will return unchanged.
					filePaths.AddRange(this.GetResolvedProviderPathFromPSPath(path, out provider));
				}
				else
				{
					// no wildcards, so don't try to expand any * or ? symbols.
					filePaths.Add(this.SessionState.Path.GetUnresolvedProviderPathFromPSPath(path, out provider, out drive));
				}
				// ensure that this path (or set of paths after wildcard expansion)
				// is on the filesystem. A wildcard can never expand to span multiple
				// providers.
				if (IsFileSystemPath(provider, path) == false)
				{
					// no, so skip to next path in _paths.
					continue;
				}
				// at this point, we have a list of paths on the filesystem.
				foreach (string filePath in filePaths)
				{
					var custom = CreateShortcut(new FileInfo(filePath));
					if (custom != null)
					{
						WriteObject(custom);
					}
				}
			}
		}

		private FileInfo CreateShortcut(FileInfo file)
		{
			WriteVerbose($"Processing {file.FullName}");
			try
			{
				if (ShouldProcess(file.FullName, "Create shortcut to file.") == false)
				{
					return null;
				}
				IShellLink link = (IShellLink)new ShellLink();

				// setup shortcut information
				link.SetDescription(file.Name);
				link.SetPath(file.FullName);

				// save it
				IPersistFile newfile = (IPersistFile)link;
				var newpath = System.IO.Path.Combine(OutputPath, $"{file.Name}.lnk");
				newfile.Save(newpath, false);

				return new FileInfo(newpath);
			}
			catch (Exception ex)
			{
				WriteError(new ErrorRecord(ex, "Error", ErrorCategory.InvalidOperation, file));
				return null;
			}
		}

		private bool IsFileSystemPath(ProviderInfo provider, string path)
		{
			var isFileSystem = true;
			// check that this provider is the filesystem
			if (provider.ImplementingType != typeof(FileSystemProvider))
			{
				// create a .NET exception wrapping our error text
				ArgumentException ex = new ArgumentException($"{path} does not resolve to a path on the FileSystem provider.");
				// wrap this in a powershell errorrecord
				ErrorRecord error = new ErrorRecord(ex, "InvalidProvider", ErrorCategory.InvalidArgument, path);
				// write a non-terminating error to pipeline
				this.WriteError(error);
				// tell our caller that the item was not on the filesystem
				isFileSystem = false;
			}
			return isFileSystem;
		}
	}

	[ComImport]
	[Guid("00021401-0000-0000-C000-000000000046")]
	internal class ShellLink
	{
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("000214F9-0000-0000-C000-000000000046")]
	internal interface IShellLink
	{
		void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, out IntPtr pfd, int fFlags);
		void GetIDList(out IntPtr ppidl);
		void SetIDList(IntPtr pidl);
		void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);
		void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
		void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
		void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
		void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);
		void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
		void GetHotkey(out short pwHotkey);
		void SetHotkey(short wHotkey);
		void GetShowCmd(out int piShowCmd);
		void SetShowCmd(int iShowCmd);
		void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath, out int piIcon);
		void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
		void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, int dwReserved);
		void Resolve(IntPtr hwnd, int fFlags);
		void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
	}
}

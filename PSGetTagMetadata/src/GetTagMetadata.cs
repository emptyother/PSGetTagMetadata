using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using Microsoft.PowerShell.Commands;

namespace OnlyHuman.PSGetTagMetadata
{
	[Cmdlet(VerbsCommon.Get, "TagMetadata", DefaultParameterSetName = ParamSetPath, SupportsShouldProcess = true)]
	public class Get_TagMetadata : PSCmdlet
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

		private string[] supportedFormats = new string[] { "bmp", "gif", "jpeg", "pbm", "pgm", "ppm", "pnm", "pcx", "png", "tiff", "dng", "svg", "jpg" };

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
					var custom = GetFileCustomObject(new FileInfo(filePath));
					if (custom != null)
					{
						WriteObject(custom);
					}
				}
			}
		}

		private PSObject GetFileCustomObject(FileInfo file)
		{
			WriteVerbose($"Processing {file.FullName}");
			WriteVerbose($"Extension: {file.Extension.ToLower()}");
			WriteVerbose($"Index: {Array.IndexOf(supportedFormats, file.Extension.Substring(1).ToLower())}");
			if (file.Extension == null || Array.IndexOf(supportedFormats, file.Extension.Substring(1).ToLower()) == -1)
			{
				WriteVerbose($"File format is not supported");
				WriteVerbose($"Supported formats: {string.Join(", ", supportedFormats)}");
				return null;
			}
			WriteVerbose($"File format is supported");
			try
			{
				var tfile = TagLib.File.Create(file.FullName);
				var tag = tfile.Tag as TagLib.Image.CombinedImageTag;
				var keywords = tag.Keywords;
				WriteVerbose($"Keywords: {keywords}");

				// New custom PSObject
				var custom = new PSObject();
				custom.Properties.Add(new PSNoteProperty("File", file));
				custom.Properties.Add(new PSNoteProperty("Keywords", keywords));
				WriteVerbose($"Custom object: {custom}");
				return custom;
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
}

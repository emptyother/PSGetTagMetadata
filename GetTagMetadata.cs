using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Internal;
using System.Management.Automation.Provider;
using System.Management.Automation.Runspaces;

namespace OnlyHuman.PSGetTagMetadata
{
	[Cmdlet(VerbsCommon.Get, "TagMetadata")]
	public class Get_TagMetadata : PSCmdlet
	{
		/// <summary>
		/// Gets or sets the path parameter to the command.
		/// </summary>
		[Parameter(
			Position = 0,
			ParameterSetName = "Path",
			Mandatory = true,
			ValueFromPipelineByPropertyName = true
		)]
		public string[] Path { get; set; }

		/// <summary>
		/// Gets or sets the literal path parameter to the command.
		/// </summary>
		[Parameter(
			ParameterSetName = "LiteralPath",
			Mandatory = true,
			ValueFromPipeline = false,
			ValueFromPipelineByPropertyName = true
		)]
		public string[] LiteralPath
		{
			get
			{
				return Path;
			}

			set
			{
				Path = value;
			}
		}

		/// <summary>
		/// This bool is used to determine if the path
		/// parameter was specified on the command line or via the pipeline.
		/// </summary>
		private bool _pipingPaths;

		/// <summary>
		/// Determines if the paths are specified on the command line
		/// or being piped in.
		/// </summary>
		protected override void BeginProcessing()
		{
			if (Path != null && Path.Length > 0)
			{
				_pipingPaths = false;
			}
			else
			{
				_pipingPaths = true;
			}
		}

		protected override void ProcessRecord()
		{
			var paths = GetAcceptedPaths(Path);
			WriteObject(paths);
		}

		protected override void EndProcessing()
		{
		}

		/// <summary>
		/// Resolves the specified paths to PathInfo objects.
		/// </summary>
		/// <param name="pathsToResolve">
		/// The paths to be resolved. Each path may contain glob characters.
		/// </param>
		/// <param name="allowNonexistingPaths">
		/// If true, resolves the path even if it doesn't exist.
		/// </param>
		/// <param name="allowEmptyResult">
		/// If true, allows a wildcard that returns no results.
		/// </param>
		/// <returns>
		/// An array of PathInfo objects that are the resolved paths for the
		/// <paramref name="pathsToResolve"/> parameter.
		/// </returns>
		internal Collection<PathInfo> ResolvePaths(
			string[] pathsToResolve,
			bool allowNonexistingPaths,
			bool allowEmptyResult
		)
		{
			Collection<PathInfo> results = new();

			foreach (string path in pathsToResolve)
			{
				var pathNotFound = false;
				ErrorRecord pathNotFoundErrorRecord = null;
				try
				{
					// First resolve each of the paths
					var pathInfos = SessionState.Path.GetResolvedPSPathFromPSPath( path);
					if (pathInfos.Count == 0)
					{
						pathNotFound = true;
					}
					foreach (PathInfo pathInfo in pathInfos)
					{
						results.Add(pathInfo);
					}
				}
				catch (PSNotSupportedException notSupported)
				{
					WriteError( new ErrorRecord( notSupported.ErrorRecord, notSupported));
				}
				catch (DriveNotFoundException driveNotFound)
				{
					WriteError( new ErrorRecord( driveNotFound.ErrorRecord, driveNotFound));
				}
				catch (ProviderNotFoundException providerNotFound)
				{
					WriteError( new ErrorRecord( providerNotFound.ErrorRecord, providerNotFound));
				}
				catch (ItemNotFoundException pathNotFoundException)
				{
					pathNotFound = true;
					pathNotFoundErrorRecord = new ErrorRecord(pathNotFoundException.ErrorRecord, pathNotFoundException);
				}
				if (pathNotFound)
				{
					if (pathNotFoundErrorRecord == null)
					{
						// Detect if the path resolution failed to resolve to a file.
						var error = $"Item not found: {path}";
						Exception e = new(error);
						pathNotFoundErrorRecord = new ErrorRecord(e, "ItemNotFound", ErrorCategory.ObjectNotFound, Path);
						WriteError(pathNotFoundErrorRecord);
					}
				}
			}
			return results;
		}

		/// <summary>
		/// Gets the list of paths accepted by the user.
		/// </summary>
		/// <param name="unfilteredPaths">The list of unfiltered paths.</param>
		/// <param name="currentContext">The current context.</param>
		/// <returns>The list of paths accepted by the user.</returns>
		private string[] GetAcceptedPaths(string[] unfilteredPaths)
		{
			var pathInfos = ResolvePaths(unfilteredPaths, true, false);
			var paths = new List<string>();
			foreach (PathInfo pathInfo in pathInfos)
			{
				paths.Add(pathInfo.Path);
			}
			return paths.ToArray();
		}
	}
}

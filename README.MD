# README

PS module with various functions.

- **Get-TagMetadata:** Retrieve the Keyword tag from supported image files.
  Returns an object containing `@{ FileInfo File, string[] Keywords }` for each
  file.
- **Set-Shortcut:** Takes an url and creates a shortcut to it in the current
  folder or in the _OutputPath_ folder. Returns a `FileInfo` object pointing to
  the created shortcut(s).

## Example of use

```ps1
Import-Module .\PSGetTagMetadata.dll
# Find all jpg and png images with just one tag, and create a shortcut to them.
$myfiles = gci -Path "C:\users\username\Pictures\" -Recurse -Include "*.jpg","*.png" | Get-TagMetadata | ? { $_.Keywords.Count -eq 1 }
$myfiles | % { $_.File } | Set-Shortcut -OutputPath .\MyFolder\ -Verbose
```

## Debugging

I have no idea how to debug a powershell module with the dotnet tool. And
`dotnet build` doesnt copy over dependency dlls to the bin folder. So for the
moment I use `dotnet publish` then import the dll. And use good old
`WriteVerbose()` (aka `console.log`) to debug this app.

## TO DO

- Turn it into a proper module with a manifest, and publish it.

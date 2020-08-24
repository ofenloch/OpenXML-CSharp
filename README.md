# Project OpenXML - Use OpenXML to maniupulate Visio Files


## Project Setup

Run `dotnet new console --name OpenXML` to create a new console application:

```bash
ofenloch@teben:~$ cd workspaces/dotnet/
ofenloch@teben:~/workspaces/dotnet$ dotnet new console --name OpenXML
The template "Console Application" was created successfully.

Processing post-creation actions...
Running 'dotnet restore' on OpenXML/OpenXML.csproj...
  Restore completed in 90.63 ms for /home/ofenloch/workspaces/dotnet/OpenXML/OpenXML.csproj.

Restore succeeded.

ofenloch@teben:~/workspaces/dotnet$ cd OpenXML/
ofenloch@teben:~/workspaces/dotnet/OpenXML$ git init
Initialized empty Git repository in /home/ofenloch/workspaces/dotnet/OpenXML/.git/
ofenloch@teben:~/workspaces/dotnet/OpenXML$ git add .
ofenloch@teben:~/workspaces/dotnet/OpenXML$ git commit -a -m"initial commit after 'dotnet new console --name OpenXML'"
[master (root-commit) 8513d1b] initial commit after 'dotnet new console --name OpenXML'
 7 files changed, 174 insertions(+)
 create mode 100644 OpenXML.csproj
 create mode 100644 Program.cs
 create mode 100644 obj/OpenXML.csproj.nuget.cache
 create mode 100644 obj/OpenXML.csproj.nuget.dgspec.json
 create mode 100644 obj/OpenXML.csproj.nuget.g.props
 create mode 100644 obj/OpenXML.csproj.nuget.g.targets
 create mode 100644 obj/project.assets.json
ofenloch@teben:~/workspaces/dotnet/OpenXML$ 
```

Run `dotnet run` to **build and run the new application**:

```bash
ofenloch@teben:~/workspaces/dotnet/OpenXML$ dotnet run
Hello World!
ofenloch@teben:~/workspaces/dotnet/OpenXML$ 
```

Run `dotnet add package DocumentFormat.OpenXml` to **add the OpenXML package as a dependency to the project**:

```bash
ofenloch@teben:~/workspaces/dotnet/OpenXML$ dotnet add package DocumentFormat.OpenXml
  Writing /tmp/tmp76HFeK.tmp
info : Adding PackageReference for package 'DocumentFormat.OpenXml' into project '/home/ofenloch/workspaces/dotnet/OpenXML/OpenXML.csproj'.
info : Restoring packages for /home/ofenloch/workspaces/dotnet/OpenXML/OpenXML.csproj...
info :   GET https://api.nuget.org/v3-flatcontainer/documentformat.openxml/index.json
info :   OK https://api.nuget.org/v3-flatcontainer/documentformat.openxml/index.json 138ms
info :   GET https://api.nuget.org/v3-flatcontainer/documentformat.openxml/2.11.3/documentformat.openxml.2.11.3.nupkg
info :   OK https://api.nuget.org/v3-flatcontainer/documentformat.openxml/2.11.3/documentformat.openxml.2.11.3.nupkg 30ms
info : Installing DocumentFormat.OpenXml 2.11.3.
info : Package 'DocumentFormat.OpenXml' is compatible with all the specified frameworks in project '/home/ofenloch/workspaces/dotnet/OpenXML/OpenXML.csproj'.
info : PackageReference for package 'DocumentFormat.OpenXml' version '2.11.3' added to file '/home/ofenloch/workspaces/dotnet/OpenXML/OpenXML.csproj'.
info : Committing restore...
info : Writing assets file to disk. Path: /home/ofenloch/workspaces/dotnet/OpenXML/obj/project.assets.json
log  : Restore completed in 2.74 sec for /home/ofenloch/workspaces/dotnet/OpenXML/OpenXML.csproj.
ofenloch@teben:~/workspaces/dotnet/OpenXML$ 
```


Run `dotnet --help` to get more info:

```bash
ofenloch@teben:~/workspaces/dotnet/OpenXML$ dotnet --help
.NET Core SDK (3.1.101)
Usage: dotnet [runtime-options] [path-to-application] [arguments]

Execute a .NET Core application.

runtime-options:
  --additionalprobingpath <path>   Path containing probing policy and assemblies to probe for.
  --additional-deps <path>         Path to additional deps.json file.
  --fx-version <version>           Version of the installed Shared Framework to use to run the application.
  --roll-forward <setting>         Roll forward to framework version  (LatestPatch, Minor, LatestMinor, Major, LatestMajor, Disable).

path-to-application:
  The path to an application .dll file to execute.

Usage: dotnet [sdk-options] [command] [command-options] [arguments]

Execute a .NET Core SDK command.

sdk-options:
  -d|--diagnostics  Enable diagnostic output.
  -h|--help         Show command line help.
  --info            Display .NET Core information.
  --list-runtimes   Display the installed runtimes.
  --list-sdks       Display the installed SDKs.
  --version         Display .NET Core SDK version in use.

SDK commands:
  add               Add a package or reference to a .NET project.
  build             Build a .NET project.
  build-server      Interact with servers started by a build.
  clean             Clean build outputs of a .NET project.
  help              Show command line help.
  list              List project references of a .NET project.
  msbuild           Run Microsoft Build Engine (MSBuild) commands.
  new               Create a new .NET project or file.
  nuget             Provides additional NuGet commands.
  pack              Create a NuGet package.
  publish           Publish a .NET project for deployment.
  remove            Remove a package or reference from a .NET project.
  restore           Restore dependencies specified in a .NET project.
  run               Build and run a .NET project output.
  sln               Modify Visual Studio solution files.
  store             Store the specified assemblies in the runtime package store.
  test              Run unit tests using the test runner specified in a .NET project.
  tool              Install or manage tools that extend the .NET experience.
  vstest            Run Microsoft Test Engine (VSTest) commands.

Additional commands from bundled tools:
  dev-certs         Create and manage development certificates.
  fsi               Start F# Interactive / execute F# scripts.
  sql-cache         SQL Server cache command-line tools.
  user-secrets      Manage development user secrets.
  watch             Start a file watcher that runs a command when files change.

Run 'dotnet [command] --help' for more information on a command.
ofenloch@teben:~/workspaces/dotnet/OpenXML$ 
```

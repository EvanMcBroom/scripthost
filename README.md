# Scripthost

[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg?style=flat)](LICENSE.txt)

Scripthost is a utility for parsing files with an ActiveX scripting engine.
The main goal of this project is to  provide a reference for implementing and using a COM server.

The scripting engines that are available on your machine may be enumerated by using [OleView .NET](https://github.com/tyranid/oleviewdotnet) and searching for COM objects that implement the `IActiveScript` interface.
The engines that are installed by default have already been documented by Geoff Chappell and are included here for convenience<sup>1</sup>:

- ECMAScript, JavaScript, JavaScript1.1, JavaScript1.2, JavaScript1.3, JScript, or LiveScript
- JScript.Compact
- JScript.Encode	
- VBS or VBScript
- VBScript.Encode
- XML

> :pencil2: Equivalent engine names are grouped on the same bullet.

## Building

Scripthost uses [CMake](https://cmake.org/) to generate and run the build system files for your platform.

```
git clone https://github.com/EvanMcBroom/scripthost.git && cd scripthost
mkdir builds && cd builds
cmake ..
cmake --build .
```

Scripthost will link against the static version of the runtime library which allows the tool to run as a standalone program on other hosts.

## Running

Scripthost is simple to use only and requires the name of an engine and one or more files to run.

```
scripthost.exe <engine name> <script path>...
```

Below is an example of running a VBScript file.

```
echo MsgBox "Hello World!" > example.vbs
scripthost.exe vbscript example.vbs
```

## References

1. [Script Languages in Internet Explorer](https://www.geoffchappell.com/notes/windows/ie/mshtml/scriptlanguage.htm)
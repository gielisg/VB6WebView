# 📘 VB6 + WPF WebView2 Integration

**Build & Deployment Guide**

This repository contains two coordinated projects:

-   **`WpfBrowserLib`** -- a WPF COM-visible WebView2 control\
-   **`VB6CallWPF`** -- a VB6 application that loads the WPF control

------------------------------------------------------------------------

## 📁 Repository Structure

    /
    ├── VB6CallWPF/            → VB6 project (forms, modules, VBP)
    ├── WpfBrowserLib/         → WPF Class Library (COM-visible WebView2)
    │   ├── BrowserControl.xaml
    │   ├── BrowserControl.xaml.cs
    │   ├── BrowserWindow.xaml
    │   ├── BrowserWindow.xaml.cs
    │   ├── BrowserHost.cs
    │   ├── WpfBrowserLib.csproj
    │   └── packages/
    └── README.md

------------------------------------------------------------------------

## 🚀 1. Prerequisites

### 🛠 Required Development Tools

  Component                               Version         Purpose
  --------------------------------------- --------------- -----------------------
  Visual Studio 2019/2022+ (2026)         Latest          Build WPF DLL
  .NET Framework Developer Pack 4.8       Required        Build WPF project
  VB6 IDE                                 Any version     Build VB6 EXE
  WebView2 Evergreen Runtime              109 or latest   Required for WebView2
  VC++ 2015--2022 Redistributable (x86)   Latest          Dependency

------------------------------------------------------------------------

## 🏗 2. Building the WPF COM DLL (`WpfBrowserLib.dll`)

1.  Open `WpfBrowserLib.sln`\
2.  Set **Release / x86**\
3.  Restore NuGet packages\
4.  Build the project

Output:

    WpfBrowserLib/bin/Release/WpfBrowserLib.dll

------------------------------------------------------------------------

## 🔐 3. Registering the DLL (COM)

Run PowerShell as Administrator:

    regasm WpfBrowserLib.dll /codebase

------------------------------------------------------------------------

## 🏗 4. Building VB6 EXE

1.  Open `VB6_CallWpf.vbp`\
2.  Add reference to **WpfBrowserLib**\
3.  Compile:\

```{=html}
<!-- -->
```
    Make VB6_CallWpf.exe

------------------------------------------------------------------------

## 📦 5. Deployment

Place all files in the same folder:

    VB6_CallWpf.exe
    WpfBrowserLib.dll
    WebView2Loader.dll
    Microsoft.WebView2.Core.dll
    (other WebView2 runtime files)

Install required runtimes:

-   .NET Framework 4.8\
-   WebView2 Runtime\
-   VC++ 2015--2022 (x86)

Register DLL:

    regasm WpfBrowserLib.dll /codebase

------------------------------------------------------------------------

## 🧪 6. Troubleshooting

### ActiveX component cannot create object

→ Register DLL again.

### Blank WebView

→ Ensure WebView2 runtime + VC++ redistributable installed.

### Missing WebView2Loader.dll

→ Copy from NuGet package "x86" folder.

------------------------------------------------------------------------

## 🎉 7. Completion

Fully working:

✔ WPF WebView2 browser\
✔ COM interface\
✔ VB6 hosting\
✔ Win7-compatible\
✔ Deployment-ready

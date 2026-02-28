# Extractor SubSetUp

A Windows Forms application for extracting and processing SubSetUp configuration data.

## Description

This is a desktop application built with C# and Windows Forms that provides tools for working with SubSetUp configuration files. The application supports Excel file operations through the EPPlus library for data import/export and manipulation.

## Features

- Windows Forms-based user interface
- Excel file support (import/export)
- Native folder picker dialog
- Data extraction and processing capabilities
- Cross-platform .NET 10.0 Windows runtime support

## Requirements

- .NET 10.0 Windows or later
- Windows operating system
- Visual Studio or any C# compatible IDE (VS Code, Visual Studio, Rider, etc.)

## Installation

1. Clone or download the repository
2. Open the `C#.sln` solution file in Visual Studio or your preferred IDE
3. Restore NuGet packages:
   ```
   dotnet restore
   ```

## Building

Build the project using Visual Studio or the .NET CLI:

```bash
# Debug build
dotnet build

# Release build
dotnet build --configuration Release
```

## Running

To run the application:

```bash
# From the project directory
dotnet run
```

Or double-click the compiled executable:
- Debug: `bin/Debug/net10.0-windows/C#.exe`
- Release: `bin/Release/net10.0-windows/win-x64/C#.exe`

## Project Structure

```
├── Form1.cs          # Main application form
├── Program.cs        # Application entry point
├── C#.csproj         # Project file
├── C#.sln            # Solution file
├── bin/              # Compiled binaries
└── obj/              # Build artifacts
```

## Dependencies

- **EPPlus** (v7.0.0) - Excel file manipulation library

## Technologies

- Language: C#
- Framework: .NET 10.0 Windows
- UI: Windows Forms
- File Format Support: Excel (.xlsx)


## Notes

- The application requires Windows to run (WinForms specific)
- Make sure EPPlus is properly installed via NuGet for Excel functionality
- This project targets .NET 10.0 or later

## Support

For issues or questions, please contact the development team.


# ConsoleExcel
A simple console application to read and write Excel files using the ClosedXML library.
It use Log4Net for logging to the subfolder : logging

# Requirements to build the project
Visual Studio 2022 or later with .NET 8 SDK installed.
Visual Studio Code with .NET 8 SDK installed.

# Usage
Open the solution file `ConsoleExcel.sln` in Visual Studio or Visual Studio Code. Then, build and run the project.

# Building the Project
To build the project, you can use the following command in the terminal:
```bash
dotnet build ConsoleExcel.sln
```

# Running the Project
To run the project, you can use the following command in the terminal with the following parameters:
```bash
dotnet run --project ConsoleExcel/ConsoleExcel.csproj --file <Excel_file> --option <sheet_name>
```

Or if you have built the project, you can run the executable directly:

```bash
ConsoleExcel.exe --file Book1.xlsx --option test1

```


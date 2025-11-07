# RunGivenVBCodeInVB

A simple Visual Basic utility that allows you to execute VB code dynamically at runtime.

## Features

- Execute Visual Basic code from a string
- Pass variables to the executed code via a dictionary
- Uses Roslyn compiler for safe, compiled execution
- Simple and straightforward API

## Usage

### Build and Run

```bash
dotnet build
dotnet run
```

### Using the RunVBCode Function

```vb
' Simple execution without variables
Dim code As String = "Console.WriteLine(""Hello World"")"
RunVBCode(code)

' Execution with variables
Dim variables As New Dictionary(Of String, Object) From {
    {"name", "Alice"},
    {"age", 25},
    {"score", 95.5}
}

Dim codeWithVars As String = "
Console.WriteLine(""Name: "" & name)
Console.WriteLine(""Age: "" & age)
Console.WriteLine(""Score: "" & score)
"

RunVBCode(codeWithVars, variables)
```

## How It Works

The `RunVBCode` function:
1. Takes your VB code as a string
2. Wraps it in a proper class structure
3. Makes dictionary variables available as local variables in the code
4. Compiles it using the Roslyn VB compiler
5. Executes the compiled code in memory

## Requirements

- .NET 6.0 or higher
- Microsoft.CodeAnalysis.VisualBasic package (installed via NuGet)
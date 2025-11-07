Imports System
Imports System.Collections.Generic
Imports System.Reflection
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.Emit
Imports System.IO
Imports System.Linq
Imports System.Runtime.Loader

Module VBCodeRunner
    ' Main function to execute Visual Basic code dynamically
    ' codeToRun: The VB code to execute as a string
    ' variables: Optional dictionary where keys become variable names and values become their values
    Public Function RunVBCode(codeToRun As String, Optional variables As Dictionary(Of String, Object) = Nothing) As Object
        Try
            ' Build the complete code with a wrapper class
            Dim variableDeclarations As String = ""
            Dim variableAssignments As String = ""

            ' Create variable declarations and assignments if dictionary is provided
            If variables IsNot Nothing Then
                For Each kvp In variables
                    Dim varType = If(kvp.Value Is Nothing, "Object", kvp.Value.GetType().Name)
                    variableDeclarations &= $"        Dim {kvp.Key} As Object = Nothing{Environment.NewLine}"
                Next
            End If

            ' Create the complete code with proper structure
            Dim completeCode As String = $"
Imports System
Imports System.Collections.Generic

Public Class DynamicCode
    Public Shared Function Execute(variables As Dictionary(Of String, Object)) As Object
{variableDeclarations}

        ' Assign variables from dictionary
        If variables IsNot Nothing Then
            For Each kvp In variables
                Select Case kvp.Key
{String.Join(Environment.NewLine, If(variables?.Select(Function(v) $"                    Case ""{v.Key}""" & Environment.NewLine & $"                        {v.Key} = kvp.Value"), Array.Empty(Of String)()))}
                End Select
            Next
        End If

        ' User code starts here
        {codeToRun}
        ' User code ends here

        Return Nothing
    End Function
End Class
"

            ' Parse the code
            Dim syntaxTree = VisualBasicSyntaxTree.ParseText(completeCode)

            ' Get references to required assemblies
            Dim references As New List(Of MetadataReference) From {
                MetadataReference.CreateFromFile(GetType(Object).Assembly.Location),
                MetadataReference.CreateFromFile(GetType(Console).Assembly.Location),
                MetadataReference.CreateFromFile(GetType(Dictionary(Of ,)).Assembly.Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Runtime").Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Collections").Location)
            }

            ' Create compilation
            Dim assemblyName As String = $"DynamicAssembly_{Guid.NewGuid().ToString("N")}"
            Dim compilation As VisualBasicCompilation = VisualBasicCompilation.Create(
                assemblyName,
                {syntaxTree},
                references,
                New VisualBasicCompilationOptions(OutputKind.DynamicallyLinkedLibrary)
            )

            ' Compile to memory
            Using ms As New MemoryStream()
                Dim result As EmitResult = compilation.Emit(ms)

                If Not result.Success Then
                    ' Compilation failed - collect errors
                    Dim errors = result.Diagnostics.Where(Function(d) d.Severity = DiagnosticSeverity.Error)
                    Dim errorMessage As String = "Compilation errors:" & Environment.NewLine
                    For Each err In errors
                        errorMessage &= err.ToString() & Environment.NewLine
                    Next
                    Throw New Exception(errorMessage)
                End If

                ' Load and execute the compiled assembly
                ms.Seek(0, SeekOrigin.Begin)
                Dim assembly = AssemblyLoadContext.Default.LoadFromStream(ms)
                Dim type = assembly.GetType("DynamicCode")
                Dim method = type.GetMethod("Execute", BindingFlags.Public Or BindingFlags.Static)

                ' Execute with variables
                Return method.Invoke(Nothing, {variables})
            End Using

        Catch ex As Exception
            Console.WriteLine($"Error executing VB code: {ex.Message}")
            Throw
        End Try
    End Function

    ' Example usage in Main
    Sub Main(args As String())
        Console.WriteLine("VB Code Runner - Example Usage")
        Console.WriteLine("=" & New String("="c, 40))

        ' Example 1: Simple code without variables
        Console.WriteLine(Environment.NewLine & "Example 1: Simple Hello World")
        Dim simpleCode As String = "Console.WriteLine(""Hello from dynamic VB code!"")"
        RunVBCode(simpleCode)

        ' Example 2: Code with variables
        Console.WriteLine(Environment.NewLine & "Example 2: Using variables")
        Dim vars As New Dictionary(Of String, Object) From {
            {"name", "John"},
            {"age", 30},
            {"salary", 50000.5}
        }

        Dim codeWithVars As String = "
Console.WriteLine(""Name: "" & name)
Console.WriteLine(""Age: "" & age)
Console.WriteLine(""Salary: "" & salary)
Dim bonus = salary * 0.1
Console.WriteLine(""Bonus: "" & bonus)
"
        RunVBCode(codeWithVars, vars)

        ' Example 3: Mathematical calculations
        Console.WriteLine(Environment.NewLine & "Example 3: Calculations")
        Dim mathVars As New Dictionary(Of String, Object) From {
            {"x", 10},
            {"y", 20}
        }

        Dim mathCode As String = "
Dim sum = CInt(x) + CInt(y)
Dim product = CInt(x) * CInt(y)
Console.WriteLine($""x = {x}, y = {y}"")
Console.WriteLine($""Sum = {sum}"")
Console.WriteLine($""Product = {product}"")
"
        RunVBCode(mathCode, mathVars)

        Console.WriteLine(Environment.NewLine & "Done!")
    End Sub
End Module

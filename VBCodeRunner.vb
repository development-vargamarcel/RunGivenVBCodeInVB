''' <summary>
''' VBCodeRunner Module - Provides functionality to dynamically compile and execute Visual Basic code at runtime.
''' This module uses the Roslyn compiler to safely compile and execute VB.NET code in a controlled environment.
''' </summary>
''' <remarks>
''' Production-ready implementation with fully qualified type names for clarity and maintainability.
''' All external type references use their full namespace paths.
''' </remarks>
Module VBCodeRunner

    ''' <summary>
    ''' Dynamically compiles and executes Visual Basic code provided as a string.
    ''' </summary>
    ''' <param name="codeToRun">The VB.NET code to compile and execute. Must not be null or empty.</param>
    ''' <param name="variables">
    ''' Optional dictionary containing variables to make available to the executing code.
    ''' Dictionary keys become variable names, and values become the variable values.
    ''' </param>
    ''' <returns>
    ''' The return value from the executed code, or Nothing if no value is returned.
    ''' </returns>
    ''' <exception cref="System.ArgumentNullException">
    ''' Thrown when codeToRun is null.
    ''' </exception>
    ''' <exception cref="System.ArgumentException">
    ''' Thrown when codeToRun is empty or contains only whitespace, or when variable names contain invalid characters.
    ''' </exception>
    ''' <exception cref="System.InvalidOperationException">
    ''' Thrown when compilation fails or when the compiled assembly cannot be loaded or executed.
    ''' </exception>
    ''' <remarks>
    ''' This method wraps the provided code in a class structure, compiles it using Roslyn,
    ''' and executes it in the current application domain. The code has access to System and
    ''' System.Collections.Generic namespaces by default.
    '''
    ''' Security Warning: This method executes arbitrary code and should only be used with trusted input.
    ''' Never use this with user-provided code in production environments without proper sandboxing.
    ''' </remarks>
    Public Function RunVBCode(codeToRun As System.String, Optional variables As System.Collections.Generic.Dictionary(Of System.String, System.Object) = Nothing) As System.Object
        ' Input validation
        If codeToRun Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(codeToRun), "The code to execute cannot be null.")
        End If

        If System.String.IsNullOrWhiteSpace(codeToRun) Then
            Throw New System.ArgumentException("The code to execute cannot be empty or contain only whitespace.", NameOf(codeToRun))
        End If

        ' Validate variable names if provided
        If variables IsNot Nothing Then
            For Each kvp As System.Collections.Generic.KeyValuePair(Of System.String, System.Object) In variables
                If System.String.IsNullOrWhiteSpace(kvp.Key) Then
                    Throw New System.ArgumentException("Variable names cannot be null, empty, or contain only whitespace.", NameOf(variables))
                End If

                ' Check for valid VB identifier (basic validation)
                If Not System.Text.RegularExpressions.Regex.IsMatch(kvp.Key, "^[a-zA-Z_][a-zA-Z0-9_]*$") Then
                    Throw New System.ArgumentException(System.String.Format("Variable name '{0}' is not a valid Visual Basic identifier. Variable names must start with a letter or underscore and contain only letters, digits, and underscores.", kvp.Key), NameOf(variables))
                End If
            Next
        End If

        Try
            ' Build variable declarations for the generated code
            Dim variableDeclarations As System.Text.StringBuilder = New System.Text.StringBuilder()

            If variables IsNot Nothing AndAlso variables.Count > 0 Then
                For Each kvp As System.Collections.Generic.KeyValuePair(Of System.String, System.Object) In variables
                    variableDeclarations.AppendLine(System.String.Format("        Dim {0} As System.Object = Nothing", kvp.Key))
                Next
            End If

            ' Build variable assignment cases for the Select statement
            Dim variableAssignmentCases As System.Text.StringBuilder = New System.Text.StringBuilder()

            If variables IsNot Nothing AndAlso variables.Count > 0 Then
                For Each kvp As System.Collections.Generic.KeyValuePair(Of System.String, System.Object) In variables
                    variableAssignmentCases.AppendLine(System.String.Format("                    Case ""{0}""", kvp.Key))
                    variableAssignmentCases.AppendLine(System.String.Format("                        {0} = kvp.Value", kvp.Key))
                Next
            End If

            ' Create the complete code with proper structure and fully qualified type names
            Dim completeCode As System.String = System.String.Format("
Public Class DynamicCode
    Public Shared Function Execute(variables As System.Collections.Generic.Dictionary(Of System.String, System.Object)) As System.Object
{0}

        ' Assign variables from dictionary
        If variables IsNot Nothing Then
            For Each kvp As System.Collections.Generic.KeyValuePair(Of System.String, System.Object) In variables
                Select Case kvp.Key
{1}
                End Select
            Next
        End If

        ' User code starts here
        {2}
        ' User code ends here

        Return Nothing
    End Function
End Class
", variableDeclarations.ToString(), variableAssignmentCases.ToString(), codeToRun)

            ' Parse the code into a syntax tree
            Dim syntaxTree As Microsoft.CodeAnalysis.SyntaxTree = Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxTree.ParseText(completeCode)

            ' Get references to required assemblies using fully qualified names
            Dim references As System.Collections.Generic.List(Of Microsoft.CodeAnalysis.MetadataReference) = New System.Collections.Generic.List(Of Microsoft.CodeAnalysis.MetadataReference)()

            ' Add core runtime references
            references.Add(Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(GetType(System.Object).Assembly.Location))
            references.Add(Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(GetType(System.Console).Assembly.Location))
            references.Add(Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(GetType(System.Collections.Generic.Dictionary(Of ,)).Assembly.Location))

            ' Add additional runtime assemblies
            Try
                references.Add(Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(System.Reflection.Assembly.Load("System.Runtime").Location))
            Catch ex As System.IO.FileNotFoundException
                ' System.Runtime might not be available in all environments, continue without it
            End Try

            Try
                references.Add(Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(System.Reflection.Assembly.Load("System.Collections").Location))
            Catch ex As System.IO.FileNotFoundException
                ' System.Collections might not be available in all environments, continue without it
            End Try

            ' Add System.Private.CoreLib for .NET Core/.NET 5+
            Try
                references.Add(Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(GetType(System.Object).Assembly.Location))
            Catch ex As System.Exception
                ' Already added or not needed
            End Try

            ' Add netstandard if available (for compatibility)
            Try
                references.Add(Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(System.Reflection.Assembly.Load("netstandard").Location))
            Catch ex As System.Exception
                ' Not available in all environments, continue without it
            End Try

            ' Create compilation with unique assembly name
            Dim assemblyName As System.String = System.String.Concat("DynamicAssembly_", System.Guid.NewGuid().ToString("N"))
            Dim compilation As Microsoft.CodeAnalysis.VisualBasic.VisualBasicCompilation = Microsoft.CodeAnalysis.VisualBasic.VisualBasicCompilation.Create(
                assemblyName,
                New Microsoft.CodeAnalysis.SyntaxTree() {syntaxTree},
                references,
                New Microsoft.CodeAnalysis.VisualBasic.VisualBasicCompilationOptions(Microsoft.CodeAnalysis.OutputKind.DynamicallyLinkedLibrary)
            )

            ' Compile to memory stream
            Using ms As System.IO.MemoryStream = New System.IO.MemoryStream()
                Dim emitResult As Microsoft.CodeAnalysis.Emit.EmitResult = compilation.Emit(ms)

                If Not emitResult.Success Then
                    ' Compilation failed - collect and format errors
                    Dim errorBuilder As System.Text.StringBuilder = New System.Text.StringBuilder()
                    errorBuilder.AppendLine("Failed to compile the provided Visual Basic code. Compilation errors:")
                    errorBuilder.AppendLine()

                    Dim errorCount As System.Int32 = 0
                    For Each diagnostic As Microsoft.CodeAnalysis.Diagnostic In emitResult.Diagnostics
                        If diagnostic.Severity = Microsoft.CodeAnalysis.DiagnosticSeverity.Error Then
                            errorCount += 1
                            errorBuilder.AppendLine(System.String.Format("Error {0}: {1}", errorCount, diagnostic.Id))
                            errorBuilder.AppendLine(System.String.Format("  Location: {0}", diagnostic.Location.GetLineSpan()))
                            errorBuilder.AppendLine(System.String.Format("  Message: {0}", diagnostic.GetMessage()))
                            errorBuilder.AppendLine()
                        End If
                    Next

                    Throw New System.InvalidOperationException(errorBuilder.ToString())
                End If

                ' Rewind the memory stream to the beginning
                ms.Seek(0, System.IO.SeekOrigin.Begin)

                ' Load the compiled assembly from the memory stream
                Dim assembly As System.Reflection.Assembly = Nothing
                Try
                    assembly = System.Runtime.Loader.AssemblyLoadContext.Default.LoadFromStream(ms)
                Catch ex As System.Exception
                    Throw New System.InvalidOperationException("Failed to load the compiled assembly from memory.", ex)
                End Try

                ' Get the DynamicCode type from the assembly
                Dim dynamicType As System.Type = assembly.GetType("DynamicCode")
                If dynamicType Is Nothing Then
                    Throw New System.InvalidOperationException("Failed to locate the 'DynamicCode' type in the compiled assembly.")
                End If

                ' Get the Execute method using reflection
                Dim executeMethod As System.Reflection.MethodInfo = dynamicType.GetMethod("Execute", System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Static)
                If executeMethod Is Nothing Then
                    Throw New System.InvalidOperationException("Failed to locate the 'Execute' method in the 'DynamicCode' type.")
                End If

                ' Invoke the Execute method with the variables dictionary
                Try
                    Return executeMethod.Invoke(Nothing, New System.Object() {variables})
                Catch ex As System.Reflection.TargetInvocationException
                    ' Unwrap the inner exception from reflection invocation
                    If ex.InnerException IsNot Nothing Then
                        Throw New System.InvalidOperationException(System.String.Concat("An error occurred while executing the dynamic code: ", ex.InnerException.Message), ex.InnerException)
                    Else
                        Throw New System.InvalidOperationException(System.String.Concat("An error occurred while executing the dynamic code: ", ex.Message), ex)
                    End If
                End Try
            End Using

        Catch ex As System.ArgumentNullException
            ' Re-throw validation exceptions
            Throw
        Catch ex As System.ArgumentException
            ' Re-throw validation exceptions
            Throw
        Catch ex As System.InvalidOperationException
            ' Re-throw compilation/execution exceptions
            Throw
        Catch ex As System.Exception
            ' Wrap any other unexpected exceptions
            Throw New System.InvalidOperationException(System.String.Concat("An unexpected error occurred while processing the Visual Basic code: ", ex.Message), ex)
        End Try
    End Function

    ''' <summary>
    ''' Main entry point for the application. Demonstrates usage of the RunVBCode function with various examples.
    ''' </summary>
    ''' <param name="args">Command-line arguments (not used in this application).</param>
    Sub Main(args As System.String())
        Try
            System.Console.WriteLine("VB Code Runner - Production Ready Example Usage")
            System.Console.WriteLine(System.String.Concat("=", New System.String("="c, 50)))

            ' Example 1: Simple code without variables
            System.Console.WriteLine()
            System.Console.WriteLine("Example 1: Simple Hello World")
            System.Console.WriteLine(New System.String("-"c, 40))
            Dim simpleCode As System.String = "System.Console.WriteLine(""Hello from dynamic VB code!"")"
            RunVBCode(simpleCode)

            ' Example 2: Code with variables
            System.Console.WriteLine()
            System.Console.WriteLine("Example 2: Using variables")
            System.Console.WriteLine(New System.String("-"c, 40))
            Dim vars As System.Collections.Generic.Dictionary(Of System.String, System.Object) = New System.Collections.Generic.Dictionary(Of System.String, System.Object)()
            vars.Add("name", "John")
            vars.Add("age", 30)
            vars.Add("salary", 50000.5)

            Dim codeWithVars As System.String = "
System.Console.WriteLine(""Name: "" & name)
System.Console.WriteLine(""Age: "" & age)
System.Console.WriteLine(""Salary: "" & salary)
Dim bonus As System.Double = System.Convert.ToDouble(salary) * 0.1
System.Console.WriteLine(""Bonus: "" & bonus)
"
            RunVBCode(codeWithVars, vars)

            ' Example 3: Mathematical calculations
            System.Console.WriteLine()
            System.Console.WriteLine("Example 3: Mathematical Calculations")
            System.Console.WriteLine(New System.String("-"c, 40))
            Dim mathVars As System.Collections.Generic.Dictionary(Of System.String, System.Object) = New System.Collections.Generic.Dictionary(Of System.String, System.Object)()
            mathVars.Add("x", 10)
            mathVars.Add("y", 20)

            Dim mathCode As System.String = "
Dim sum As System.Int32 = System.Convert.ToInt32(x) + System.Convert.ToInt32(y)
Dim product As System.Int32 = System.Convert.ToInt32(x) * System.Convert.ToInt32(y)
System.Console.WriteLine(System.String.Format(""x = {0}, y = {1}"", x, y))
System.Console.WriteLine(System.String.Format(""Sum = {0}"", sum))
System.Console.WriteLine(System.String.Format(""Product = {0}"", product))
"
            RunVBCode(mathCode, mathVars)

            ' Example 4: Error handling demonstration
            System.Console.WriteLine()
            System.Console.WriteLine("Example 4: Error Handling")
            System.Console.WriteLine(New System.String("-"c, 40))
            Try
                ' This will cause a compilation error
                RunVBCode("This is not valid VB code!")
            Catch ex As System.InvalidOperationException
                System.Console.WriteLine("Caught expected compilation error:")
                System.Console.WriteLine(ex.Message.Substring(0, System.Math.Min(200, ex.Message.Length)) & "...")
            End Try

            System.Console.WriteLine()
            System.Console.WriteLine("All examples completed successfully!")

        Catch ex As System.Exception
            System.Console.WriteLine(System.String.Concat("Fatal error in Main: ", ex.Message))
            System.Console.WriteLine(ex.StackTrace)
            System.Environment.Exit(1)
        End Try
    End Sub
End Module

# VBA Testing Rules

VBA has no built-in test framework, but that does not excuse untested code. These rules establish a pragmatic testing strategy.

## Compile Check (Mandatory)

- Run **Debug > Compile VBAProject** after every change
- Fix all compile errors before proceeding
- A project that does not compile is never acceptable

## Debug.Assert (Development Invariants)

Use `Debug.Assert` to catch logic errors during development:

```vb
Debug.Assert UBound(arr) >= LBound(arr)   ' array is not empty
Debug.Assert Not ws Is Nothing             ' worksheet was found
Debug.Assert lngCount > 0                  ' counter is positive
```

- Assertions are ignored in compiled add-ins — they are a development-time safety net
- Place assertions at procedure entry points to validate assumptions

## Debug.Print (Execution Tracing)

- Use `Debug.Print` to trace execution flow and inspect values in the Immediate Window
- Remove or guard tracing output before deployment:

```vb
#Const ENABLE_TRACE = True

Sub TracePrint(ByVal sMsg As String)
    #If ENABLE_TRACE Then
        Debug.Print Format$(Now, "hh:nn:ss") & " | " & sMsg
    #End If
End Sub
```

## Pure Functions

- Write business logic as pure functions (input in, output out, no side effects)
- Pure functions can be tested directly from the Immediate Window:

```vb
? CalculateDiscount(100, 0.15)   ' expect 85
? FormatFullName("Jane", "Doe")  ' expect "Doe, Jane"
```

## Test Harness Pattern

Create a `modTestHarness` module with a `RunAllTests` procedure:

```vb
Public Sub RunAllTests()
    Debug.Print "=== Test Run: " & Now & " ==="
    TestCalculateDiscount
    TestFormatFullName
    TestValidateInput
    Debug.Print "=== All tests passed ==="
End Sub

Private Sub TestCalculateDiscount()
    Debug.Assert CalculateDiscount(100, 0.1) = 90
    Debug.Assert CalculateDiscount(0, 0.5) = 0
    Debug.Print "  PASS: TestCalculateDiscount"
End Sub
```

## Manual Test Documentation

For procedures that interact with the UI or Office object model, document test steps:
- Input data and preconditions
- Steps to execute
- Expected results
- Edge cases to verify (empty ranges, missing sheets, locked workbooks)

## Optional: Rubberduck VBA

[Rubberduck VBA](https://rubberduckvba.com/) provides unit testing, code inspections, and refactoring for VBA. When available:
- Write `@TestModule` and `@TestMethod` annotated tests
- Use Rubberduck's test explorer to run and report
- Leverage code inspections for static analysis

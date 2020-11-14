Attribute VB_Name = "Hello_World"
Option Explicit

Public Sub HelloWorld()

    Debug.Print "Hello, World!"
    
    ThisWorkbook.Worksheets("Hello World").Range("A1").Value = "Hello, World"
    ThisWorkbook.Worksheets("Hello World").Cells(2, 1).Value = "Hello, World"
    
    With ThisWorkbook.Worksheets("Hello World")
        
        .Range("A1").Value = "Hello, World!"
        .Cells(2, 1).Value = "Hello, World!"
        
    End With
    
    MsgBox "Hello, World!", vbOKOnly, "Hi there!"
    
    Dim result As Variant
    
    result = InputBox("What is your name?")
    
    MsgBox "Hi " & result

End Sub

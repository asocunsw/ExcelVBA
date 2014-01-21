' demonstrate common data types and how to use them

Sub demo_variables()
    
    ' type Dim then the variales name
    ' followed by as and then the type of variable
    ' we store text in strings
    Dim name As String
    ' now excel nows we have a variable called name
    ' which stores some text but doesnt know what untill we set it
    name = "John Smith"
       
    ' we can stole whole numbers between -32768 and 32767
    Dim age As Integer
    age = 19
    
    ' double can store very large numbers
    Dim salary As Double
    salary = 172954
    
    ' boolean can store 2 values
    ' either True or False
    Dim alive As Boolean
    alive = True
    
    ' use the double type to store decimal numbers as well
    Dim something As Double
    something = 3.141598765432
      
    ' you can put variables together by using a & between them
    ' put vbNewLine wherever you want the line in the message box dialogue to end
    MsgBox ("Name: " & name & vbNewLine & "Age: " & age & vbNewLine & "Salary: " & salary & vbNewLine & "Is Alive?: " & alive & vbNewLine & "something: " & something)
    
End Sub


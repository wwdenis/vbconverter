# Welcome!

The VB Converter analyses VB6 code and converts it to .NET 2.0 (C# or VB). 
It uses a modified implementation VBParser (of the wondeful work of Paul Vick) and CodeDom.

Feel free to download it, contribute and comment.
For any other issues, please contact us: vbconverter@gmail.com

Migrated from Codeplex: https://vbconverter.codeplex.com/

## Features
* Convert part of Visual Basic 6 code to C# and VB .NET
* Batch converter (Files)

## Roadmap
* Refactor the CodeWriter module.
* Add more VB6 expressions (such as Print)

## Examples

**Methods (Visual Basic 6)**

```vb
Public Sub Main()
    Dim name As String
    name = InputBox("Enter your name: ", "VB6 App")
    If name = "" Then
        MsgBox "You must enter your name !", vbCritical, "Error"
    Else
        MsgBox "Your name is: " & name, vbInformation, "Info"
    End If
End Sub
```

**Methods (C#)**
```csharp
public void Main() {
    string name;
    name = Interaction.InputBox("Enter your name: ", "VB6 App");
    if ((name == "")) {
        Interaction.MsgBox("You must enter your name !", Constants.vbCritical, "Error");
    }
    else {
        Interaction.MsgBox(("Your name is: " + name), Constants.vbInformation, "Info");
    }
}
```

## Project Structure
- __VBConverter.CodeParser:__ Responsible for parsing the VB code into a tree.
- __VBConverter.CodeWriter:__ Responsible for converting the tree structure to .NET.
- __VBConverter.UI:__ User interface (Windows Forms) responsible for testing the convertion.

## Examples

**Declarations (Visual Basic 6)**
```vb
Dim name As String
Private list(5) As Integer
Protected rs As ADODB.Recordset
Friend birthDate As Date
Public Const CUSTOMER As Integer = &H16
```

**Declarations (C#)**
```csharp
string name;
private short[] list = new short[6];
protected ADODB.Recordset rs;
internal DateTime birthDate;
public const short CUSTOMER = 22;
```

**Methods (Visual Basic 6)**
```vb
Private Function Sum(ByVal num1 As Integer, ByVal num2 As Integer) As Integer
    Sum = num1 + num2
End Function
Sub Load()
    db.Load
End Sub
```

**Methods (C#)**
```csharp
private short Sum(short num1, short num2) {
    short _result_Sum = default(short);
    _result_Sum = (num1 + num2);
    return _result_Sum;
}
private void Load() {
    db.Load();
}
```

**Properties (Visual Basic 6)**
```vb
Public Property Get Name() As String
    Name = m_Name
End Function
Public Property Let Name(newName As String)
    m_Name = newName
End Function
```

**Properties (C#)**
```csharp
public string Name {
    get {
        return m_Name;
    }
    set {
        m_Name = value;
    }
}
```

**Events (Visual Basic 6)**
```vb
Friend Event OnLoad(ByVal Status As Integer)
```

**Properties (C#)**
```csharp
internal event OnLoadHandler OnLoad;
internal delegate void OnLoadHandler(short Status);
```

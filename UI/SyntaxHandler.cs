using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace VBConverter.UI
{
    // https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/keywords/
    // https://docs.microsoft.com/en-us/dotnet/csharp/language-reference/keywords/
    public static class SyntaxHandler
    {
        public static readonly string[] VisualBasicKeywords =
        {
            "AddHandler",
            "AddressOf",
            "Alias",
            "And",
            "AndAlso",
            "As",
            "Boolean",
            "ByRef",
            "Byte",
            "ByVal",
            "Call",
            "Case",
            "Catch",
            "CBool",
            "CByte",
            "CChar",
            "CDate",
            "CDbl",
            "CDec",
            "Char",
            "CInt",
            "Class",
            "CLng",
            "CObj",
            "Const",
            "Continue",
            "CSByte",
            "CShort",
            "CSng",
            "CStr",
            "CType",
            "CUInt",
            "CULng",
            "CUShort",
            "Date",
            "Decimal",
            "Declare",
            "Default",
            "Delegate",
            "Dim",
            "DirectCast",
            "Do",
            "Double",
            "Each",
            "Else",
            "ElseIf",
            "End Statement",
            "End",
            "EndIf",
            "Enum",
            "Erase",
            "Error",
            "Event",
            "Exit",
            "False",
            "Finally",
            "For",
            "For",
            "Friend",
            "Function",
            "Get",
            "GetType",
            "GetXMLNamespace",
            "Global",
            "GoSub",
            "GoTo",
            "Handles",
            "If",
            "Implements",
            "Statement",
            "Imports",
            "In",
            "Inherits",
            "Integer",
            "Interface",
            "Is",
            "IsNot",
            "Let",
            "Lib",
            "Like",
            "Long",
            "Loop",
            "Me",
            "Mod",
            "Module",
            "Module Statement",
            "MustInherit",
            "MustOverride",
            "MyBase",
            "MyClass",
            "Namespace",
            "Narrowing",
            "New Constraint",
            "New Operator",
            "Next",
            "Not",
            "Nothing",
            "NotInheritable",
            "NotOverridable",
            "Object",
            "Of",
            "On",
            "Operator",
            "Option",
            "Optional",
            "Or",
            "OrElse",
            "Out",
            "Overloads",
            "Overridable",
            "Overrides",
            "ParamArray",
            "Partial",
            "Private",
            "Property",
            "Protected",
            "Public",
            "RaiseEvent",
            "ReadOnly",
            "ReDim",
            "REM",
            "RemoveHandler",
            "Resume",
            "Return",
            "SByte",
            "Select",
            "Set",
            "Shadows",
            "Shared",
            "Short",
            "Single",
            "Static",
            "Step",
            "Stop",
            "String",
            "Structure",
            "Sub",
            "SyncLock",
            "Then",
            "Throw",
            "To",
            "True",
            "Try",
            "TryCast",
            "TypeOf",
            "UInteger",
            "ULong",
            "UShort",
            "Using",
            "Variant",
            "Wend",
            "When",
            "While",
            "Widening",
            "With",
            "WithEvents",
            "WriteOnly",
            "Xor",
            "Aggregate",
            "Ansi",
            "Assembly",
            "Async",
            "Auto",
            "Await",
            "Binary",
            "Compare",
            "Custom",
            "Distinct",
            "Equals",
            "Explicit",
            "From",
            "Group",
            "Join",
            "Into",
            "IsFalse",
            "IsTrue",
            "Iterator",
            "Join",
            "Key",
            "Mid",
            "Off",
            "Order",
            "By",
            "Preserve",
            "Skip",
            "Strict",
            "Take",
            "Take While",
            "Text",
            "Unicode",
            "Until",
            "Where",
            "Yield"
        };

        public static readonly string[] CSharpKeywords =
        {
            "abstract",
            "as",
            "base",
            "bool",
            "break",
            "byte",
            "case",
            "catch",
            "char",
            "checked",
            "class",
            "const",
            "continue",
            "decimal",
            "default",
            "delegate",
            "do",
            "double",
            "else",
            "enum",
            "event",
            "explicit",
            "extern",
            "false",
            "finally",
            "fixed",
            "float",
            "for",
            "foreach",
            "goto",
            "if",
            "implicit",
            "in",
            "int",
            "interface",
            "internal",
            "is",
            "lock",
            "long",
            "namespace",
            "new",
            "null",
            "object",
            "operator",
            "out",
            "override",
            "params",
            "private",
            "protected",
            "public",
            "readonly",
            "ref",
            "return",
            "sbyte",
            "sealed",
            "short",
            "sizeof",
            "stackalloc",
            "static",
            "string",
            "struct",
            "switch",
            "this",
            "throw",
            "true",
            "try",
            "typeof",
            "uint",
            "ulong",
            "unchecked",
            "unsafe",
            "ushort",
            "using",
            "using static",
            "virtual",
            "void",
            "volatile",
            "while",
            "add",
            "alias",
            "ascending",
            "async",
            "await",
            "by",
            "descending",
            "dynamic",
            "equals",
            "from",
            "get",
            "global",
            "group",
            "into",
            "join",
            "let",
            "nameof",
            "on",
            "orderby",
            "partial",
            "remove",
            "select",
            "set",
            "value",
            "var",
            "when",
            "where",
            "yield",
        };

        public static readonly Dictionary<Color, string> CSharpStyles = new Dictionary<Color, string>
        {
            { Color.Blue, @"\b("  + string.Join("|", CSharpKeywords) + @")\b" }, // Keywords
            { Color.Green, @"(\/\/.+?$|\/\*.+?\*\/)" }, // Comments
            { Color.Brown, "\".+?\"" }, // Strings
            { Color.DarkCyan, @"\b(Console|DateTime)\b" }, // Types
        };

        public static readonly Dictionary<Color, string> VisualBasicStyles = new Dictionary<Color, string>
        {
            { Color.Blue, @"\b("  + string.Join("|", VisualBasicKeywords) + @")\b" }, // Keywords
            { Color.Green, "(^'|'').*$" }, // Comments
            { Color.Brown, "\".+?\"" }, // Strings
            { Color.DarkCyan, @"\b(Console|DateTime)\b" }, // Types
        };

        public static void SetVisualBasic(RichTextBox editor)
        {
            SetSyntax(editor, VisualBasicStyles, true);
        }

        public static void SetCSharp(RichTextBox editor)
        {
            SetSyntax(editor, CSharpStyles, false);
        }

        private static void SetSyntax(RichTextBox editor, Dictionary<Color, string> mappings, bool ignoreCase)
        {
            var options = ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None;

            var start = editor.SelectionStart;
            var length = editor.SelectionLength;
            var color = Color.Black;

            // Prevents blinking
            editor.SelectionStart = 0;
            editor.SelectionLength = editor.Text.Length;
            editor.SelectionColor = color;

            foreach (var map in mappings)
            {
                var matches = Regex.Matches(editor.Text, map.Value, options);
                foreach (Match m in matches)
                {
                    editor.SelectionStart = m.Index;
                    editor.SelectionLength = m.Length;
                    editor.SelectionColor = map.Key;
                    editor.SelectionFont = new Font(editor.SelectionFont, FontStyle.Bold);
                }
            }

            // restoring the original colors, for further writing
            editor.SelectionStart = start;
            editor.SelectionLength = length;
            editor.SelectionColor = color;
            editor.SelectionFont = new Font(editor.SelectionFont, FontStyle.Regular);

            // giving back the focus
            editor.Focus();
        }
    }
}

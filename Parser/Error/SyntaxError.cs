using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
//
// Visual Basic .NET Parser
//
// Copyright (C) 2005, Microsoft Corporation. All rights reserved.
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
// EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
// MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
//
namespace VBConverter.CodeParser.Error
{
	/// <summary>
	/// A syntax error.
	/// </summary>
	public sealed class SyntaxError
	{
        static string[] ErrorMessages = null;

        private readonly SyntaxErrorType _Type;
		private readonly Span _Span;

		/// <summary>
		/// The type of the syntax error.
		/// </summary>
		public SyntaxErrorType Type {
			get { return _Type; }
		}

		/// <summary>
		/// The location of the syntax error.
		/// </summary>
		public Span Span {
			get { return _Span; }
		}

		/// <summary>
		/// Constructs a new syntax error.
		/// </summary>
		/// <param name="type">The type of the syntax error.</param>
		/// <param name="span">The location of the syntax error.</param>
		public SyntaxError(SyntaxErrorType type, Span span)
		{
			Debug.Assert(System.Enum.IsDefined(typeof(SyntaxErrorType), type));
			_Type = type;
			_Span = span;
		}

		public override string ToString()
		{
            int typeValue = (int)this.Type;
            string description = GetErrorMessage(this.Type);
            string message = string.Format("error {0} {1}: {2}", typeValue, Span.ToString(), description);
            return message;
		}

        static SyntaxError()
        {
            BuildErrorMessages();
        }

        static public string GetErrorMessage(SyntaxErrorType type)
        {
            int index = (int)type;
            string message = "Unknown error.";
            if (index <= ErrorMessages.GetUpperBound(0))
                message = ErrorMessages[index];

            return message;
        }

        static private void BuildErrorMessages()
        {
            string[] Strings = new string[] {
                "", 
                "Invalid escaped identifier.", 
                "Invalid identifier.", 
                "Cannot put a type character on a keyword.", 
                "Invalid character.", 
                "Invalid character literal.", 
                "Invalid string literal.", 
                "Invalid date literal.", 
                "Invalid floating point literal.", 
                "Invalid integer literal.", 
			    "Invalid decimal literal.", 
                "Syntax error.", 
                "Expected ','.", 
                "Expected '('.", 
                "Expected ')'.", 
                "Expected '='.", 
                "Expected 'As'.", 
                "Expected '}'.", 
                "Expected '.'.", 
                "Expected '-'.", 
			    "Expected 'Is'.", 
                "Expected '>'.", 
                "Type expected.", 
                "Expected identifier.", 
                "Invalid use of keyword.", 
                "Bounds can be specified only for the top-level array when initializing an array of arrays.", 
                "Array bounds cannot appear in type specifiers.", 
                "Expected expression.", 
                "Comma, ')', or a valid expression continuation expected.", 
                "Expected named argument.", 
			    "MyBase must be followed by a '.'.", 
                "MyClass must be followed by a '.'.", 
                "Exit must be followed by Do, For, While, Select, Sub, Function, Property or Try.", 
                "Expected 'Next'.", 
                "Expected 'Resume' or 'GoTo'.", 
                "Expected 'Error'.", 
                "Type character does not match declared data type String.", 
                "Comma, '}', or a valid expression continuation expected.", 
                "Expected one of 'Dim', 'Const', 'Public', 'Private', 'Protected', 'Friend', 'Shadows', 'ReadOnly' or 'Shared'.", 
                "End of statement expected.", 
			    "'Do' must end with a matching 'Loop'.", 
                "'While' must end with a matching 'End While'.", 
                "'Select' must end with a matching 'End Select'.", 
                "'SyncLock' must end with a matching 'End SyncLock'.", 
                "'With' must end with a matching 'End With'.", 
                "'If' must end with a matching 'End If'.", 
                "'Try' must end with a matching 'End Try'.", 
                "'Sub' must end with a matching 'End Sub'.", 
                "'Function' must end with a matching 'End Function'.", 
                "'Property' must end with a matching 'End Property'.", 
			    "'Get' must end with a matching 'End Get'.", 
                "'Set' must end with a matching 'End Set'.", 
                "'Class' must end with a matching 'End Class'.", 
                "'Structure' must end with a matching 'End Structure'.", 
                "'Module' must end with a matching 'End Module'.", 
                "'Interface' must end with a matching 'End Interface'.", 
                "'Enum' must end with a matching 'End Enum'.", 
                "'Namespace' must end with a matching 'End Namespace'.", 
                "'Loop' cannot have a condition if matching 'Do' has one.", 
                "'Loop' must be preceded by a matching 'Do'.", 
			    "'Next' must be preceded by a matching 'For' or 'For Each'.", 
                "'End While' must be preceded by a matching 'While'.", 
                "'End Select' must be preceded by a matching 'Select'.", 
                "'End SyncLock' must be preceded by a matching 'SyncLock'.", 
                "'End If' must be preceded by a matching 'If'.", 
                "'End Try' must be preceded by a matching 'Try'.", 
                "'End With' must be preceded by a matching 'With'.", 
                "'Catch' cannot appear outside a 'Try' statement.", 
                "'Finally' cannot appear outside a 'Try' statement.", 
                "'Catch' cannot appear after 'Finally' within a 'Try' statement.", 
			    "'Finally' can only appear once in a 'Try' statement.", 
                "'Case' must be preceded by a matching 'Select'.", 
                "'Case' cannot appear after 'Case Else' within a 'Select' statement.", 
                "'Case Else' can only appear once in a 'Select' statement.", 
                "'Case Else' must be preceded by a matching 'Select'.", 
                "'End Sub' must be preceded by a matching 'Sub'.", 
                "'End Function' must be preceded by a matching 'Function'.", 
                "'End Property' must be preceded by a matching 'Property'.", 
                "'End Get' must be preceded by a matching 'Get'.", 
                "'End Set' must be preceded by a matching 'Set'.", 
			    "'End Class' must be preceded by a matching 'Class'.", 
                "'End Structure' must be preceded by a matching 'Structure'.", 
                "'End Module' must be preceded by a matching 'Module'.", 
                "'End Interface' must be preceded by a matching 'Interface'.", 
                "'End Enum' must be preceded by a matching 'Enum'.", 
                "'End Namespace' must be preceded by a matching 'Namespace'.", 
                "Statements and labels are not valid between 'Select Case' and first 'Case'.", 
                "'ElseIf' cannot appear after 'Else' within an 'If' statement.", 
                "'ElseIf' must be preceded by a matching 'If'.", 
                "'Else' can only appear once in an 'If' statement.", 
			    "'Else' must be preceded by a matching 'If'.", 
                "Statement cannot end a block outside of a line 'If' statement.", 
                "Attribute of this type is not allowed here.", 
                "Modifier cannot be specified twice.", 
                "Modifier is not valid on this declaration type.", 
                "Can only specify one of 'Dim', 'Static' or 'Const'.", 
                "Events cannot have a return type.", 
                "Comma or ')' expected.", 
                "Method declaration statements must be the first statement on a logical line.", 
                "First statement of a method body cannot be on the same line as the method declaration.", 
			    "'End Sub' must be the first statement on a line.", 
                "'End Function' must be the first statement on a line.", 
                "'End Get' must be the first statement on a line.", 
                "'End Set' must be the first statement on a line.", 
                "'Sub' or 'Function' expected.", 
                "String constant expected.", 
                "'Lib' expected.", 
                "Declaration cannot appear within a Property declaration.", 
                "Declaration cannot appear within a Class declaration.", 
                "Declaration cannot appear within a Structure declaration.", 
			    "Declaration cannot appear within a Module declaration.", 
                "Declaration cannot appear within an Interface declaration.", 
                "Declaration cannot appear within an Enum declaration.", 
                "Declaration cannot appear within a Namespace declaration.", 
                "Specifiers and attributes are not valid on this statement.", 
                "Specifiers and attributes are not valid on a Namespace declaration.", 
                "Specifiers and attributes are not valid on an Imports declaration.", 
                "Specifiers and attributes are not valid on an Option declaration.", 
                "Inherits' can only specify one class.", 
                "'Inherits' statements must precede all declarations.", 
			    "'Implements' statements must follow any 'Inherits' statement and precede all declarations in a class.", 
                "Enum must contain at least one member.", 
                "'Option Explicit' can be followed only by 'On' or 'Off'.", 
                "'Option Strict' can be followed only by 'On' or 'Off'.", 
                "'Option Compare' must be followed by 'Text' or 'Binary'.", 
                "'Option' must be followed by 'Compare', 'Explicit', or 'Strict'.", 
                "'Option' statements must precede any declarations or 'Imports' statements.", 
                "'Imports' statements must precede any declarations.", 
                "Assembly or Module attribute statements must precede any declarations in a file.", 
                "'End' statement not valid.", 
			    "Expected relational operator.", 
                "'If', 'ElseIf', 'Else', 'End If', or 'Const' expected.", 
                "Expected integer literal.", 
                "'#ExternalSource' statements cannot be nested.", 
                "'ExternalSource', 'Region' or 'If' expected.", 
                "'#End ExternalSource' must be preceded by a matching '#ExternalSource'.", 
                "'#ExternalSource' must end with a matching '#End ExternalSource'.", 
                "'#End Region' must be preceded by a matching '#Region'.", 
                "'#Region' must end with a matching '#End Region'.", 
                "'#Region' and '#End Region' statements are not valid within method bodies.", 
			    "Conversions to and from 'String' cannot occur in a conditional compilation expression.", 
                "Conversion is not valid in a conditional compilation expression.", 
                "Expression is not valid in a conditional compilation expression.", 
                "Operator is not valid for these types in a conditional compilation expression.", 
                "'#If' must end with a matching '#End If'.", 
                "'#End If' must be preceded by a matching '#If'.", 
                "'#ElseIf' cannot appear after '#Else' within an '#If' statement.", 
                "'#ElseIf' must be preceded by a matching '#If'.", 
                "'#Else' can only appear once in an '#If' statement.", 
                "'#Else' must be preceded by a matching '#If'.", 
			    "'Global' not allowed in this context; identifier expected.", 
                "Modules cannot be generic.", 
                "Expected 'Of'.", 
                "Operator declaration must be one of:  +, -, *, \\, /, ^, &, Like, Mod, And, Or, Xor, Not, <<, >>, =, <>, <, <=, >, >=, CType, IsTrue, IsFalse.",
                "'Operator' must end with a matching 'End Operator'.", 
                "'End Operator' must be preceded by a matching 'Operator'.", 
                "'End Operator' must be the first statement on a line.", 
                "Properties cannot be generic.", 
                "Constructors cannot be generic.", 
                "Operators cannot be generic.", 
			    "Global must be followed by a '.'.", 
                "Continue must be followed by Do, For, or While.", 
                "'Using' must end with a matching 'End Using'.", 
                "Custom 'Event' must end with a matching 'End Event'.", 
                "'AddHandler' must end with a matching 'End AddHandler'.", 
                "'RemoveHandler' must end with a matching 'End RemoveHandler'.", 
                "'RaiseEvent' must end with a matching 'End RaiseEvent'.", 
                "'End Using' must be preceded by a matching 'Using'.", 
                "'End Event' must be preceded by a matching custom 'Event'.", 
                "'End AddHandler' must be preceded by a matching 'AddHandler'.", 
			    "'End RemoveHandler' must be preceded by a matching 'RemoveHandler'.", 
                "'End RaiseEvent' must be preceded by a matching 'RaiseEvent'.", 
                "'End AddHandler' must be the first statement on a line.", 
                "'End RemoveHandler' must be the first statement on a line.", 
                "'End RaiseEvent' must be the first statement on a line.",
                "Expected: Get or Let or Set.",
                "Argument required for Property Let or Property Set",
                "Statement invalid inside Type Block",
                "Expected: 0 or 1",
                "Only comments may appear End Sub, End Function, or End Property"
            };

            ErrorMessages = Strings;
        }
	}
}

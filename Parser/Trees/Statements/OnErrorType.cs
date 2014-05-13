using System;
using System.Collections.Generic;
using System.Text;
//
// Visual Basic .NET Parser
//
// Copyright (C) 2005, Microsoft Corporation. All rights reserved.
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
// EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
// MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
//
namespace VBConverter.CodeParser.Trees.Statements
{
	/// <summary>
	/// The type of an On Error statement.
	/// </summary>
	public enum OnErrorType
	{
		Bad,
		Next,
		// On Error Resume Next
		Zero,
		// On Error GoTo 0
		MinusOne,
		// On Error GoTo -1
		Label
		// On Error GoTo <label>
	}
}

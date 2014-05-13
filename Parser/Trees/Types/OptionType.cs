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
namespace VBConverter.CodeParser.Trees.Types
{
	/// <summary>
	/// The type of an Option declaration.
	/// </summary>
	public enum OptionType
	{
		SyntaxError,
		Explicit,
		ExplicitOn,
		ExplicitOff,
		Strict,
		StrictOn,
		StrictOff,
		CompareBinary,
		CompareText,
        BaseZero,
        BaseOne
	}
}

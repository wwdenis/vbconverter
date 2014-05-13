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
namespace VBConverter.CodeParser.Trees.Attributes
{
	/// <summary>
	/// The type of an attribute usage.
	/// </summary>
	[Flags()]
	public enum AttributeTypes
	{
		/// <summary>Regular application.</summary>
		Regular = 1,

		/// <summary>Applied to the netmodule.</summary>
		Module = 2,

		/// <summary>Applied to the assembly.</summary>
		Assembly = 4
	}
}

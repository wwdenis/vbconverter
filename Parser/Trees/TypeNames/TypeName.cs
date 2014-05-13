using System;
using System.Collections.Generic;
using System.Diagnostics;
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
namespace VBConverter.CodeParser.Trees.TypeNames
{
	/// <summary>
	/// A parse tree for a type name.
	/// </summary>
	public abstract class TypeName : Tree
	{

		protected TypeName(TreeType type, Span span) : base(type, span)
		{

			Debug.Assert(type >= TreeType.IntrinsicType && type <= TreeType.ArrayType);
		}
	}
}

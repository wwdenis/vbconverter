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
namespace VBConverter.CodeParser.Trees.Imports
{
	/// <summary>
	/// A parse tree for an Imports statement.
	/// </summary>
	public abstract class Import : Tree
	{

		protected Import(TreeType type, Span span) : base(type, span)
		{

			Debug.Assert(type == TreeType.AliasImport || type == TreeType.NameImport);
		}
	}
}

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
namespace VBConverter.CodeParser.Trees.Names
{
	/// <summary>
	/// A parse tree for a name.
	/// </summary>
	public abstract class Name : Tree
	{
		protected Name(TreeType type, Span span) : base(type, span)
		{
			Debug.Assert(type >= TreeType.SimpleName && type <= TreeType.MyBaseName);
		}

        public abstract string FullName
        {
            get;
        }
	}
}

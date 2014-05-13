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
namespace VBConverter.CodeParser.Trees.Expressions
{
	/// <summary>
	/// A parse tree for Nothing.
	/// </summary>
	public sealed class GlobalExpression : Expression
	{

		/// <summary>
		/// Constructs a new parse tree for Global.
		/// </summary>
		/// <param name="span">The location of the parse tree.</param>
		public GlobalExpression(Span span) : base(TreeType.GlobalExpression, span)
		{
		}
	}
}

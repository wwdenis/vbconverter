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
namespace VBConverter.CodeParser.Trees.Expressions
{
	/// <summary>
	/// A parse tree for an expression.
	/// </summary>
	public class Expression : Tree
	{

		/// <summary>
		/// Creates a bad expression.
		/// </summary>
		/// <param name="span">The location of the parse tree.</param>
		/// <returns>A bad expression.</returns>
		public static Expression GetBadExpression(Span span)
		{
			return new Expression(span);
		}

		/// <summary>
		/// Whether the expression is constant or not.
		/// </summary>
		public virtual bool IsConstant {
			get { return false; }
		}

		protected Expression(TreeType type, Span span) : base(type, span)
		{

			Debug.Assert(type >= TreeType.SimpleNameExpression && type <= TreeType.GetTypeExpression);
		}

		private Expression(Span span) : base(TreeType.SyntaxError, span)
		{
		}

		public override bool IsBad {
			get { return Type == TreeType.SyntaxError; }
		}
	}
}

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
	/// A parse tree for a parenthesized expression.
	/// </summary>
	public sealed class ParentheticalExpression : UnaryExpression
	{

		private readonly Location _RightParenthesisLocation;

		/// <summary>
		/// The location of the ')'.
		/// </summary>
		public Location RightParenthesisLocation {
			get { return _RightParenthesisLocation; }
		}

		/// <summary>
		/// Constructs a new parenthesized expression parse tree.
		/// </summary>
		/// <param name="operand">The operand of the expression.</param>
		/// <param name="rightParenthesisLocation">The location of the ')'.</param>
		/// <param name="span">The location of the parse tree.</param>
		public ParentheticalExpression(Expression operand, Location rightParenthesisLocation, Span span) : base(TreeType.ParentheticalExpression, operand, span)
		{

			_RightParenthesisLocation = rightParenthesisLocation;
		}
	}
}

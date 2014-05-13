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
using VBConverter.CodeParser.Trees.TypeNames;

namespace VBConverter.CodeParser.Trees.Expressions
{
	/// <summary>
	/// A parse tree for a CType expression.
	/// </summary>
	public sealed class CTypeExpression : CastTypeExpression
	{

		/// <summary>
		/// Constructs a new parse tree for a CType expression.
		/// </summary>
		/// <param name="leftParenthesisLocation">The location of the '('.</param>
		/// <param name="operand">The expression to be converted.</param>
		/// <param name="commaLocation">The location of the ','.</param>
		/// <param name="target">The target type of the conversion.</param>
		/// <param name="rightParenthesisLocation">The location of the ')'.</param>
		/// <param name="span">The location of the parse tree.</param>
		public CTypeExpression(Location leftParenthesisLocation, Expression operand, Location commaLocation, TypeName target, Location rightParenthesisLocation, Span span) : base(TreeType.CTypeExpression, leftParenthesisLocation, operand, commaLocation, target, rightParenthesisLocation, span)
		{
		}
	}
}

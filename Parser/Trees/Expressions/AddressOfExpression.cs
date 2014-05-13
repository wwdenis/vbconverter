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
	/// A parse tree for an AddressOf expression.
	/// </summary>
	public sealed class AddressOfExpression : UnaryExpression
	{

		/// <summary>
		/// Constructs a new AddressOf expression parse tree.
		/// </summary>
		/// <param name="operand">The operand of AddressOf.</param>
		/// <param name="span">The location of the parse tree.</param>
		public AddressOfExpression(Expression operand, Span span) : base(TreeType.AddressOfExpression, operand, span)
		{
		}
	}
}

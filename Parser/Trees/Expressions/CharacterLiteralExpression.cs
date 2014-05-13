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
	/// A parse tree for a character literal expression.
	/// </summary>
	public sealed class CharacterLiteralExpression : LiteralExpression<char>
	{
		/// <summary>
		/// Constructs a new parse tree for a character literal expression.
		/// </summary>
		/// <param name="literal">The literal value.</param>
		/// <param name="span">The location of the parse tree.</param>
		public CharacterLiteralExpression(char literal, Span span) : base(literal, TreeType.CharacterLiteralExpression, span)
		{
		}
	}
}

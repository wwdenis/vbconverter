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
namespace VBConverter.CodeParser.Tokens.Literal
{
	/// <summary>
	/// A string literal.
	/// </summary>
	public sealed class StringLiteralToken : LiteralToken<string>
	{
		/// <summary>
		/// Constructs a new string literal token.
		/// </summary>
		/// <param name="literal">The value of the literal.</param>
		/// <param name="span">The location of the literal.</param>
		public StringLiteralToken(string literal, Span span) : base(TokenType.StringLiteral, literal, span)
		{
			if (literal == null)
				throw new ArgumentNullException("literal");
		}
	}
}

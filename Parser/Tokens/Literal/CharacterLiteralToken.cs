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
	/// A character literal.
	/// </summary>
	public sealed class CharacterLiteralToken : LiteralToken<char>
	{
		/// <summary>
		/// Constructs a new character literal token.
		/// </summary>
		/// <param name="literal">The literal value.</param>
		/// <param name="span">The location of the literal.</param>
		public CharacterLiteralToken(char literal, Span span) : base(TokenType.CharacterLiteral, literal, span)
		{
		}
	}
}

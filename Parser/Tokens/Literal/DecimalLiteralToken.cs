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

using VBConverter.CodeParser.Base;

namespace VBConverter.CodeParser.Tokens.Literal
{
	/// <summary>
	/// A decimal literal token.
	/// </summary>
	public sealed class DecimalLiteralToken : LiteralToken<decimal>
	{
        // The type character after the literal, if any
		private readonly TypeCharacter _TypeCharacter;

		/// <summary>
		/// The type character of the literal.
		/// </summary>
		public TypeCharacter TypeCharacter {
			get { return _TypeCharacter; }
		}

		/// <summary>
		/// Constructs a new decimal literal token.
		/// </summary>
		/// <param name="literal">The literal value.</param>
		/// <param name="typeCharacter">The literal's type character.</param>
		/// <param name="span">The location of the literal.</param>
		public DecimalLiteralToken(decimal literal, TypeCharacter typeCharacter, Span span) : base(TokenType.DecimalLiteral, literal, span)
		{

			if (typeCharacter != TypeCharacter.None && typeCharacter != TypeCharacter.DecimalChar && typeCharacter != TypeCharacter.DecimalSymbol)
				throw new ArgumentOutOfRangeException("typeCharacter");
			
			_TypeCharacter = typeCharacter;
		}
	}
}

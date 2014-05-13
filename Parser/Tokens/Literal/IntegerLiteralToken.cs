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
	/// An integer literal.
	/// </summary>
	public sealed class IntegerLiteralToken : LiteralToken<long>
	{
        // The type character after the literal, if any
		private readonly TypeCharacter _TypeCharacter;
		// The base of the literal
		private readonly IntegerBase _IntegerBase;

		/// <summary>
		/// The type character of the literal.
		/// </summary>
		public TypeCharacter TypeCharacter {
			get { return _TypeCharacter; }
		}

		/// <summary>
		/// The integer base of the literal.
		/// </summary>
		public IntegerBase IntegerBase {
			get { return _IntegerBase; }
		}

		/// <summary>
		/// Constructs a new integer literal.
		/// </summary>
		/// <param name="literal">The literal value.</param>
		/// <param name="integerBase">The integer base of the literal.</param>
		/// <param name="typeCharacter">The type character of the literal.</param>
		/// <param name="span">The location of the literal.</param>
		public IntegerLiteralToken(long literal, IntegerBase integerBase, TypeCharacter typeCharacter, Span span) : base(TokenType.IntegerLiteral, literal, span)
		{
            if (!Enum.IsDefined(typeof(IntegerBase), integerBase))
            	throw new ArgumentOutOfRangeException("integerBase");

			if (typeCharacter != TypeCharacter.None && typeCharacter != TypeCharacter.IntegerSymbol && typeCharacter != TypeCharacter.IntegerChar && typeCharacter != TypeCharacter.ShortChar && typeCharacter != TypeCharacter.LongSymbol && typeCharacter != TypeCharacter.LongChar)
				throw new ArgumentOutOfRangeException("typeCharacter");

            _IntegerBase = integerBase;
			_TypeCharacter = typeCharacter;
		}
	}
}

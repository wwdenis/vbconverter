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
using VBConverter.CodeParser.Error;

namespace VBConverter.CodeParser.Tokens
{
	/// <summary>
	/// A lexical error.
	/// </summary>
	public sealed class ErrorToken : Token
	{

		private readonly SyntaxError _SyntaxError;

		/// <summary>
		/// The syntax error that represents the lexical error.
		/// </summary>
		public SyntaxError SyntaxError {
			get { return _SyntaxError; }
		}

		/// <summary>
		/// Creates a new lexical error token.
		/// </summary>
		/// <param name="errorType">The type of the error.</param>
		/// <param name="span">The location of the error.</param>
		public ErrorToken(SyntaxErrorType errorType, Span span) : base(TokenType.LexicalError, span)
		{
            if (errorType < SyntaxErrorType.InvalidEscapedIdentifier || errorType > SyntaxErrorType.InvalidDecimalLiteral)
				throw new ArgumentOutOfRangeException("errorType");
			
			_SyntaxError = new SyntaxError(errorType, span);
		}
	}
}

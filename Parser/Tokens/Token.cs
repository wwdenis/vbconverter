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
namespace VBConverter.CodeParser.Tokens
{
	/// <summary>
	/// The base class for all tokens. Contains line and column information as well as token type.
	/// </summary>
	public abstract class Token
	{
        private readonly TokenType _Type;
		// A span ends at the first character beyond the token
		private readonly Span _Span;

		/// <summary>
		/// The type of the token.
		/// </summary>
		public TokenType Type {
			get { return _Type; }
		}

		/// <summary>
		/// The span of the token in the source text.
		/// </summary>
		public Span Span {
			get { return _Span; }
		}

		protected Token(TokenType type, Span span)
		{
			Debug.Assert(Enum.IsDefined(typeof(TokenType), type));
			_Type = type;
			_Span = span;
		}

		/// <summary>
		/// Returns the unreserved keyword type of an identifier.
		/// </summary>
		/// <returns>The unreserved keyword type of an identifier, the token's type otherwise.</returns>
		public virtual TokenType AsUnreservedKeyword()
		{
			return Type;
		}

		public override string ToString()
		{
			return this.Type.ToString();
		}
	}
}

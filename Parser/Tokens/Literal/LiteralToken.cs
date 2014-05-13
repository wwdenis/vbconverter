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
	/// A generic literal token.
    /// <typeparam name="LiteralType">The literal type</typeparam>
	/// </summary>
	public abstract class LiteralToken<LiteralType> : Token
	{
        private readonly LiteralType _Literal;
        
		/// <summary>
		/// The literal value.
		/// </summary>
        public LiteralType Literal
        {
			get { return _Literal; }
		}

		/// <summary>
		/// Constructs a new literal token.
		/// </summary>
        /// <param name="type">The token type</param>
		/// <param name="literal">The literal value.</param>
		/// <param name="span">The location of the literal.</param>
        public LiteralToken(TokenType type, LiteralType literal, Span span) : base(type, span)
		{
			_Literal = literal;
		}
	}
}

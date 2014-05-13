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
	/// A date/time literal.
	/// </summary>
	public sealed class DateLiteralToken : LiteralToken<DateTime>
	{
		/// <summary>
		/// Constructs a new date literal instance.
		/// </summary>
		/// <param name="literal">The literal value.</param>
		/// <param name="span">The location of the literal.</param>
		public DateLiteralToken(DateTime literal, Span span) : base(TokenType.DateLiteral, literal, span)
		{
		}
	}
}

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
namespace VBConverter.CodeParser.Tokens.Terminator
{
	/// <summary>
	/// A line terminator.
	/// </summary>
	public sealed class LineTerminatorToken : Token
	{
		/// <summary>
		/// Create a new line terminator token.
		/// </summary>
		/// <param name="span">The location of the line terminator.</param>
		public LineTerminatorToken(Span span) : base(TokenType.LineTerminator, span)
		{
		}
	}
}

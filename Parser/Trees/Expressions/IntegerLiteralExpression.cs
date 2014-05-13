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

namespace VBConverter.CodeParser.Trees.Expressions
{
	/// <summary>
	/// A parse tree for an integer literal.
	/// </summary>
	public sealed class IntegerLiteralExpression : NumericLiteralExpression<long>
	{
		private readonly IntegerBase _IntegerBase;

		/// <summary>
		/// The integer base of the literal.
		/// </summary>
		public IntegerBase IntegerBase 
        {
			get { return _IntegerBase; }
		}

		/// <summary>
		/// Constructs a new parse tree for an integer literal.
		/// </summary>
		/// <param name="literal">The literal value.</param>
		/// <param name="integerBase">The integer base of the literal.</param>
		/// <param name="typeCharacter">The type character on the literal.</param>
		/// <param name="span">The location of the parse tree.</param>
        public IntegerLiteralExpression(long literal, IntegerBase integerBase, TypeCharacter typeCharacter, Span span) : base(literal, typeCharacter, TreeType.IntegerLiteralExpression, span)
		{
            if (!Enum.IsDefined(typeof(IntegerBase), integerBase))
				throw new ArgumentOutOfRangeException("integerBase");
			
			_IntegerBase = integerBase;
		}
	}
}

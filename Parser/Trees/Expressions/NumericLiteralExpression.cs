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
	/// A parse tree for a numeric literal expression (integer, decimal or floating).
	/// </summary>
	public abstract class NumericLiteralExpression<NumericType> : LiteralExpression<NumericType>
	{
		private readonly TypeCharacter _TypeCharacter;

		/// <summary>
		/// The type character on the literal value.
		/// </summary>
		public TypeCharacter TypeCharacter 
        {
			get { return _TypeCharacter; }
		}

		/// <summary>
		/// Constructs a new parse tree for a numeric literal.
		/// </summary>
		/// <param name="literal">The literal value.</param>
		/// <param name="typeCharacter">The type character on the literal value.</param>
		/// <param name="span">The location of the parse tree.</param>
        public NumericLiteralExpression(NumericType literal, TypeCharacter typeCharacter, TreeType treeType, Span span) : base(literal, treeType, span)
		{
            ValidateTypeCharacter(typeCharacter);
			
			_TypeCharacter = typeCharacter;
		}

        private void ValidateTypeCharacter(TypeCharacter typeCharacter)
        {
            TypeCharacter[] types = new TypeCharacter[] { TypeCharacter.None };

            if (typeof(NumericType) == typeof(long))
                types = new TypeCharacter[] 
                    { 
                        TypeCharacter.None,
                        TypeCharacter.IntegerSymbol,
                        TypeCharacter.IntegerChar,
                        TypeCharacter.ShortChar,
                        TypeCharacter.LongSymbol,
                        TypeCharacter.LongChar
                    };
            else if (typeof(NumericType) == typeof(decimal))
                types = new TypeCharacter[] 
                    { 
                        TypeCharacter.None, 
                        TypeCharacter.DecimalChar, 
                        TypeCharacter.DecimalSymbol 
                    };
            else if (typeof(NumericType) == typeof(double))
                types = new TypeCharacter[] 
                    { 
                        TypeCharacter.None,
                        TypeCharacter.SingleSymbol,
                        TypeCharacter.SingleChar,
                        TypeCharacter.DoubleSymbol,
                        TypeCharacter.DoubleChar
                    };

            bool found = false;

            foreach (TypeCharacter item in types)
            {
                if (item == typeCharacter)
                {
                    found = true;
                    break;
                }
            }

            if (!found)
                throw new ArgumentOutOfRangeException("typeCharacter");
        }
	}
}

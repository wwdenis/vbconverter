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
namespace VBConverter.CodeParser.Trees.Expressions
{
	/// <summary>
	/// A parse tree for a literal expression.
	/// </summary>
	public abstract class LiteralExpression<LiteralType> : Expression
	{
        private readonly LiteralType _literal;

        /// <summary>
        /// The literal value.
        /// </summary>
        public LiteralType Literal
        {
            get { return _literal; }
        }
		
        public override sealed bool IsConstant 
        {
			get { return true; }
		}

		protected LiteralExpression(LiteralType literal, TreeType type, Span span) : base(type, span)
		{
            Debug.Assert(type >= TreeType.StringLiteralExpression && type <= TreeType.BooleanLiteralExpression);
            _literal = literal;
		}
	}
}

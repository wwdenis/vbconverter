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
	/// A parse tree for an expression that has an operand.
	/// </summary>
	public abstract class UnaryExpression : Expression
	{

		private readonly Expression _Operand;

		/// <summary>
		/// The operand of the expression.
		/// </summary>
		public Expression Operand {
			get { return _Operand; }
		}

		protected UnaryExpression(TreeType type, Expression operand, Span span) : base(type, span)
		{

			Debug.Assert(type == TreeType.ParentheticalExpression || type == TreeType.TypeOfExpression || (type == TreeType.CTypeExpression || type == TreeType.DirectCastExpression || type == TreeType.TryCastExpression) || type == TreeType.IntrinsicCastExpression || (type >= TreeType.UnaryOperatorExpression && type <= TreeType.AddressOfExpression));

			if (operand == null)
			{
				throw new ArgumentNullException("operand");
			}

			SetParent(operand);
			_Operand = operand;
		}

		public override bool IsConstant {
			get { return Operand.IsConstant; }
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Operand);
		}
	}
}

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
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.Expressions
{
	/// <summary>
	/// A parse tree for an intrinsic conversion expression.
	/// </summary>
	public sealed class IntrinsicCastExpression : UnaryExpression
	{

		private readonly IntrinsicType _IntrinsicType;
		private readonly Location _LeftParenthesisLocation;
		private readonly Location _RightParenthesisLocation;

		/// <summary>
		/// The intrinsic type conversion.
		/// </summary>
		public IntrinsicType IntrinsicType {
			get { return _IntrinsicType; }
		}

		/// <summary>
		/// The location of the '('.
		/// </summary>
		public Location LeftParenthesisLocation {
			get { return _LeftParenthesisLocation; }
		}

		/// <summary>
		/// The location of the ')'.
		/// </summary>
		public Location RightParenthesisLocation {
			get { return _RightParenthesisLocation; }
		}

		/// <summary>
		/// Constructs a new parse tree for an intrinsic conversion expression.
		/// </summary>
		/// <param name="intrinsicType">The intrinsic type conversion.</param>
		/// <param name="leftParenthesisLocation">The location of the '('.</param>
		/// <param name="operand">The expression to convert.</param>
		/// <param name="rightParenthesisLocation">The location of the ')'.</param>
		/// <param name="span">The location of the parse tree.</param>
		public IntrinsicCastExpression(IntrinsicType intrinsicType, Location leftParenthesisLocation, Expression operand, Location rightParenthesisLocation, Span span) : base(TreeType.IntrinsicCastExpression, operand, span)
		{

            if (!Enum.IsDefined(typeof(IntrinsicType), intrinsicType))
				throw new ArgumentOutOfRangeException("intrinsicType");
			
			if (operand == null)
				throw new ArgumentNullException("operand");
			
			_IntrinsicType = intrinsicType;
			_LeftParenthesisLocation = leftParenthesisLocation;
			_RightParenthesisLocation = rightParenthesisLocation;
		}
	}
}

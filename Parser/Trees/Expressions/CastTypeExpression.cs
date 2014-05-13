using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
//
// Visual Basic .NET Parser
//
// Copyright (C) 2005, Microsoft Corporation. All rights reserved.
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
// EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
// MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
//
using VBConverter.CodeParser.Trees.TypeNames;

namespace VBConverter.CodeParser.Trees.Expressions
{
	/// <summary>
	/// A parse tree for a CType or DirectCast expression.
	/// </summary>
	public abstract class CastTypeExpression : UnaryExpression
	{

		private readonly Location _LeftParenthesisLocation;
		private readonly Location _CommaLocation;
		private readonly TypeName _Target;
		private readonly Location _RightParenthesisLocation;

		/// <summary>
		/// The location of the '('.
		/// </summary>
		public Location LeftParenthesisLocation {
			get { return _LeftParenthesisLocation; }
		}

		/// <summary>
		/// The location of the ','.
		/// </summary>
		public Location CommaLocation {
			get { return _CommaLocation; }
		}

		/// <summary>
		/// The target type for the operand.
		/// </summary>
		public TypeName Target {
			get { return _Target; }
		}

		/// <summary>
		/// The location of the ')'.
		/// </summary>
		public Location RightParenthesisLocation {
			get { return _RightParenthesisLocation; }
		}

		protected CastTypeExpression(TreeType type, Location leftParenthesisLocation, Expression operand, Location commaLocation, TypeName target, Location rightParenthesisLocation, Span span) : base(type, operand, span)
		{

			Debug.Assert(type == TreeType.CTypeExpression || type == TreeType.DirectCastExpression || type == TreeType.TryCastExpression);

			if (target == null)
			{
				throw new ArgumentNullException("target");
			}

			SetParent(target);

			_Target = target;
			_LeftParenthesisLocation = leftParenthesisLocation;
			_CommaLocation = commaLocation;
			_RightParenthesisLocation = rightParenthesisLocation;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			base.GetChildTrees(childList);
			AddChild(childList, Target);
		}
	}
}

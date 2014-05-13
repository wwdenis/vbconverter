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
using VBConverter.CodeParser.Trees.TypeNames;

namespace VBConverter.CodeParser.Trees.Expressions
{
	/// <summary>
	/// A parse tree for a TypeOf ... Is expression.
	/// </summary>
	public sealed class TypeOfExpression : UnaryExpression
	{

		private readonly Location _IsLocation;
		private readonly TypeName _Target;

		/// <summary>
		/// The location of the 'Is'.
		/// </summary>
		public Location IsLocation {
			get { return _IsLocation; }
		}

		/// <summary>
		/// The target type for the operand.
		/// </summary>
		public TypeName Target {
			get { return _Target; }
		}

		/// <summary>
		/// Constructs a new parse tree for a TypeOf ... Is expression.
		/// </summary>
		/// <param name="operand">The target value.</param>
		/// <param name="isLocation">The location of the 'Is'.</param>
		/// <param name="target">The target type to check against.</param>
		/// <param name="span">The location of the parse tree.</param>
		public TypeOfExpression(Expression operand, Location isLocation, TypeName target, Span span) : base(TreeType.TypeOfExpression, operand, span)
		{

			if (target == null)
			{
				throw new ArgumentNullException("target");
			}

			SetParent(target);

			_Target = target;
			_IsLocation = isLocation;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			base.GetChildTrees(childList);
			AddChild(childList, Target);
		}
	}
}

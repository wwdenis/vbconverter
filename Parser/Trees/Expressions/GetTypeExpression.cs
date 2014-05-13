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
	/// A parse tree for a GetType expression.
	/// </summary>
	public sealed class GetTypeExpression : Expression
	{

		private readonly Location _LeftParenthesisLocation;
		private readonly TypeName _Target;
		private readonly Location _RightParenthesisLocation;

		/// <summary>
		/// The location of the '('.
		/// </summary>
		public Location LeftParenthesisLocation {
			get { return _LeftParenthesisLocation; }
		}

		/// <summary>
		/// The target type of the GetType expression.
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

		/// <summary>
		/// Constructs a new parse tree for a GetType expression.
		/// </summary>
		/// <param name="leftParenthesisLocation">The location of the '('.</param>
		/// <param name="target">The target type of the GetType expression.</param>
		/// <param name="rightParenthesisLocation">The location of the ')'.</param>
		/// <param name="span">The location of the parse tree.</param>
		public GetTypeExpression(Location leftParenthesisLocation, TypeName target, Location rightParenthesisLocation, Span span) : base(TreeType.GetTypeExpression, span)
		{

			if (target == null)
			{
				throw new ArgumentNullException("target");
			}

			SetParent(target);

			_LeftParenthesisLocation = leftParenthesisLocation;
			_Target = target;
			_RightParenthesisLocation = rightParenthesisLocation;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Target);
		}
	}
}

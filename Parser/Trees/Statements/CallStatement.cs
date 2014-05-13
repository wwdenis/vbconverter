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
using VBConverter.CodeParser.Trees.Arguments;
using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Expressions;

namespace VBConverter.CodeParser.Trees.Statements
{
	/// <summary>
	/// A parse tree for a method call statement.
	/// </summary>
	public sealed class CallStatement : Statement
	{

		private readonly Location _CallLocation;
		private readonly Expression _TargetExpression;
		private readonly ArgumentCollection _Arguments;

		/// <summary>
		/// The location of the 'Call', if any.
		/// </summary>
		public Location CallLocation {
			get { return _CallLocation; }
		}

		/// <summary>
		/// The target of the call.
		/// </summary>
		public Expression TargetExpression {
			get { return _TargetExpression; }
		}

		/// <summary>
		/// The arguments to the call.
		/// </summary>
		public ArgumentCollection Arguments {
			get { return _Arguments; }
		}

		/// <summary>
		/// Constructs a new parse tree for a method call statement.
		/// </summary>
		/// <param name="callLocation">The location of the 'Call', if any.</param>
		/// <param name="targetExpression">The target of the call.</param>
		/// <param name="arguments">The arguments to the call.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments of the parse tree.</param>
		public CallStatement(Location callLocation, Expression targetExpression, ArgumentCollection arguments, Span span, IList<Comment> comments) : base(TreeType.CallStatement, span, comments)
		{

			if (targetExpression == null)
			{
				throw new ArgumentNullException("targetExpression");
			}

			SetParent(targetExpression);
			SetParent(arguments);

			_CallLocation = callLocation;
			_TargetExpression = targetExpression;
			_Arguments = arguments;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, TargetExpression);
			AddChild(childList, Arguments);
		}
	}
}

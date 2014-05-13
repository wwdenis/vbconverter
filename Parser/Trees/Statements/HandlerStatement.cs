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
using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Expressions;

namespace VBConverter.CodeParser.Trees.Statements
{
	/// <summary>
	/// A parse tree for an AddHandler or RemoveHandler statement.
	/// </summary>
	public abstract class HandlerStatement : Statement
	{

		private readonly Expression _Name;
		private readonly Location _CommaLocation;
		private readonly Expression _DelegateExpression;

		/// <summary>
		/// The event name.
		/// </summary>
		public Expression Name {
			get { return _Name; }
		}

		/// <summary>
		/// The location of the ','.
		/// </summary>
		public Location CommaLocation {
			get { return _CommaLocation; }
		}

		/// <summary>
		/// The delegate expression.
		/// </summary>
		public Expression DelegateExpression {
			get { return _DelegateExpression; }
		}

		protected HandlerStatement(TreeType type, Expression name, Location commaLocation, Expression delegateExpression, Span span, IList<Comment> comments) : base(type, span, comments)
		{

			Debug.Assert(type == TreeType.AddHandlerStatement || type == TreeType.RemoveHandlerStatement);

			if (name == null)
			{
				throw new ArgumentNullException("name");
			}

			if (delegateExpression == null)
			{
				throw new ArgumentNullException("delegateExpression");
			}

			SetParent(name);
			SetParent(delegateExpression);

			_Name = name;
			_CommaLocation = commaLocation;
			_DelegateExpression = delegateExpression;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Name);
			AddChild(childList, DelegateExpression);
		}
	}
}

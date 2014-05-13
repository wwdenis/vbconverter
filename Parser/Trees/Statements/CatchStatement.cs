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
using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Expressions;
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.TypeNames;

namespace VBConverter.CodeParser.Trees.Statements
{
	/// <summary>
	/// A parse tree for a Catch statement.
	/// </summary>
	public sealed class CatchStatement : Statement
	{

		private readonly SimpleName _Name;
		private readonly Location _AsLocation;
		private readonly TypeName _ExceptionType;
		private readonly Location _WhenLocation;
		private readonly Expression _FilterExpression;

		/// <summary>
		/// The name of the catch variable, if any.
		/// </summary>
		public SimpleName Name {
			get { return _Name; }
		}

		/// <summary>
		/// The location of the 'As', if any.
		/// </summary>
		public Location AsLocation {
			get { return _AsLocation; }
		}

		/// <summary>
		/// The type of the catch variable, if any.
		/// </summary>
		public TypeName ExceptionType {
			get { return _ExceptionType; }
		}

		/// <summary>
		/// The location of the 'When', if any.
		/// </summary>
		public Location WhenLocation {
			get { return _WhenLocation; }
		}

		/// <summary>
		/// The filter expression, if any.
		/// </summary>
		public Expression FilterExpression {
			get { return _FilterExpression; }
		}

		/// <summary>
		/// Constructs a new parse tree for a Catch statement.
		/// </summary>
		/// <param name="name">The name of the catch variable, if any.</param>
		/// <param name="asLocation">The location of the 'As', if any.</param>
		/// <param name="exceptionType">The type of the catch variable, if any.</param>
		/// <param name="whenLocation">The location of the 'When', if any.</param>
		/// <param name="filterExpression">The filter expression, if any.</param>
		/// <param name="span">The location of the parse tree, if any.</param>
		/// <param name="comments">The comments for the parse tree, if any.</param>
		public CatchStatement(SimpleName name, Location asLocation, TypeName exceptionType, Location whenLocation, Expression filterExpression, Span span, IList<Comment> comments) : base(TreeType.CatchStatement, span, comments)
		{

			SetParent(name);
			SetParent(exceptionType);
			SetParent(filterExpression);

			_Name = name;
			_AsLocation = asLocation;
			_ExceptionType = exceptionType;
			_WhenLocation = whenLocation;
			_FilterExpression = filterExpression;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Name);
			AddChild(childList, ExceptionType);
			AddChild(childList, FilterExpression);
		}
	}
}

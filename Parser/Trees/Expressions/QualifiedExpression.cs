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
using VBConverter.CodeParser.Trees.Names;

namespace VBConverter.CodeParser.Trees.Expressions
{
	/// <summary>
	/// A parse tree for a qualified name expression.
	/// </summary>
	public sealed class QualifiedExpression : Expression
	{

		private Expression _Qualifier;
		private readonly Location _DotLocation;
		private readonly SimpleName _Name;

		/// <summary>
		/// The expression qualifying the name.
		/// </summary>
		public Expression Qualifier {
			get { return _Qualifier; }
            set { _Qualifier  =value; }
		}

		/// <summary>
		/// The location of the '.'.
		/// </summary>
		public Location DotLocation {
			get { return _DotLocation; }
		}

		/// <summary>
		/// The qualified name.
		/// </summary>
		public SimpleName Name {
			get { return _Name; }
		}

		/// <summary>
		/// Constructs a new parse tree for a qualified name expression.
		/// </summary>
		/// <param name="qualifier">The expression qualifying the name.</param>
		/// <param name="dotLocation">The location of the '.'.</param>
		/// <param name="name">The qualified name.</param>
		/// <param name="span">The location of the parse tree.</param>
		public QualifiedExpression(Expression qualifier, Location dotLocation, SimpleName name, Span span) : base(TreeType.QualifiedExpression, span)
		{
			if (name == null)
				throw new ArgumentNullException("name");
		
			SetParent(qualifier);
			SetParent(name);

			_Qualifier = qualifier;
			_DotLocation = dotLocation;
			_Name = name;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Qualifier);
			AddChild(childList, Name);
		}
	}
}

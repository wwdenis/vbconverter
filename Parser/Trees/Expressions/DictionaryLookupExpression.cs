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
	/// A parse tree for a dictionary lookup expression.
	/// </summary>
	public sealed class DictionaryLookupExpression : Expression
	{

		private readonly Expression _Qualifier;
		private readonly Location _BangLocation;
		private readonly SimpleName _Name;

		/// <summary>
		/// The dictionary expression.
		/// </summary>
		public Expression Qualifier {
			get { return _Qualifier; }
		}

		/// <summary>
		/// The location of the '!'.
		/// </summary>
		public Location BangLocation {
			get { return _BangLocation; }
		}

		/// <summary>
		/// The name to look up.
		/// </summary>
		public SimpleName Name {
			get { return _Name; }
		}

		/// <summary>
		/// Constructs a new parse tree for a dictionary lookup expression.
		/// </summary>
		/// <param name="qualifier">The dictionary expression.</param>
		/// <param name="bangLocation">The location of the '!'.</param>
		/// <param name="name">The name to look up..</param>
		/// <param name="span">The location of the parse tree.</param>
		public DictionaryLookupExpression(Expression qualifier, Location bangLocation, SimpleName name, Span span) : base(TreeType.DictionaryLookupExpression, span)
		{

			if (name == null)
			{
				throw new ArgumentNullException("name");
			}

			SetParent(qualifier);
			SetParent(name);

			_Qualifier = qualifier;
			_BangLocation = bangLocation;
			_Name = name;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Qualifier);
			AddChild(childList, Name);
		}
	}
}

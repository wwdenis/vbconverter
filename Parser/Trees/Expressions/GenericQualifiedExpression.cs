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
	/// A parse tree for a qualified name expression.
	/// </summary>
	public sealed class GenericQualifiedExpression : Expression
	{

		private readonly Expression _Base;
		private readonly TypeArgumentCollection _TypeArguments;

		/// <summary>
		/// The base expression.
		/// </summary>
		public Expression Base {
			get { return _Base; }
		}

		/// <summary>
		/// The qualifying type arguments.
		/// </summary>
		public TypeArgumentCollection TypeArguments {
			get { return _TypeArguments; }
		}

		/// <summary>
		/// Constructs a new parse tree for a generic qualified expression.
		/// </summary>
		/// <param name="base">The base expression.</param>
		/// <param name="typeArguments">The qualifying type arguments.</param>
		/// <param name="span">The location of the parse tree.</param>
		public GenericQualifiedExpression(Expression @base, TypeArgumentCollection typeArguments, Span span) : base(TreeType.GenericQualifiedExpression, span)
		{

			if (@base == null)
			{
				throw new ArgumentNullException("base");
			}

			if (typeArguments == null)
			{
				throw new ArgumentNullException("typeArguments");
			}

			SetParent(@base);
			SetParent(typeArguments);

			_Base = @base;
			_TypeArguments = typeArguments;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Base);
			AddChild(childList, TypeArguments);
		}
	}
}

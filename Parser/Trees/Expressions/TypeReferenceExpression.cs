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
	/// A parse tree for an expression that refers to a type.
	/// </summary>
	public sealed class TypeReferenceExpression : Expression
	{

		private readonly TypeName _TypeName;

		/// <summary>
		/// The name of the type being referred to.
		/// </summary>
		public TypeName TypeName {
			get { return _TypeName; }
		}

		/// <summary>
		/// Constructs a new parse tree for a type reference.
		/// </summary>
		/// <param name="typeName">The name of the type being referred to.</param>
		/// <param name="span">The location of the parse tree.</param>
		public TypeReferenceExpression(TypeName typeName, Span span) : base(TreeType.TypeReferenceExpression, span)
		{

			if (typeName == null)
			{
				throw new ArgumentNullException("typeName");
			}

			SetParent(typeName);
			_TypeName = typeName;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, TypeName);
		}
	}
}

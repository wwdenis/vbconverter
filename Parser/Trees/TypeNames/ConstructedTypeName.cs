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
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.TypeNames
{
	/// <summary>
	/// A parse tree for a constructed generic type name.
	/// </summary>
	public sealed class ConstructedTypeName : NamedTypeName
	{

		private readonly TypeArgumentCollection _TypeArguments;

		/// <summary>
		/// The type arguments.
		/// </summary>
		public TypeArgumentCollection TypeArguments {
			get { return _TypeArguments; }
		}

		/// <summary>
		/// Constructs a new parse tree for a generic constructed type name.
		/// </summary>
		/// <param name="name">The generic type being constructed.</param>
		/// <param name="typeArguments">The type arguments.</param>
		/// <param name="span">The location of the parse tree.</param>
		public ConstructedTypeName(Name name, TypeArgumentCollection typeArguments, Span span) : base(TreeType.ConstructedType, name, span)
		{

			if (typeArguments == null)
			{
				throw new ArgumentNullException("typeArguments");
			}

			SetParent(typeArguments);
			_TypeArguments = typeArguments;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, TypeArguments);
		}
	}
}

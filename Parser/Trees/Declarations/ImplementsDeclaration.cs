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
using VBConverter.CodeParser.Trees.TypeNames;
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.Declarations
{
	/// <summary>
	/// A parse tree for an Implements declaration.
	/// </summary>
	public sealed class ImplementsDeclaration : Declaration
	{

		private readonly TypeNameCollection _ImplementedTypes;

		/// <summary>
		/// The list of types.
		/// </summary>
		public TypeNameCollection ImplementedTypes {
			get { return _ImplementedTypes; }
		}

		/// <summary>
		/// Constructs a parse tree for an Implements declaration.
		/// </summary>
		/// <param name="implementedTypes">The types inherited or implemented.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public ImplementsDeclaration(TypeNameCollection implementedTypes, Span span, IList<Comment> comments) : base(TreeType.ImplementsDeclaration, span, comments)
		{

			if (implementedTypes == null)
			{
				throw new ArgumentNullException("implementedTypes");
			}

			SetParent(implementedTypes);

			_ImplementedTypes = implementedTypes;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			base.GetChildTrees(childList);

			AddChild(childList, ImplementedTypes);
		}
	}
}

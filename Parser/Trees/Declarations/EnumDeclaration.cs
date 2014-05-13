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
using VBConverter.CodeParser.Trees.Attributes;
using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Modifiers;
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.TypeNames;
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.Declarations
{
	/// <summary>
	/// A parse tree for an Enum declaration.
	/// </summary>
	public sealed class EnumDeclaration : BlockDeclaration
	{

		private readonly Location _AsLocation;
		private readonly TypeName _ElementType;

		/// <summary>
		/// The location of the 'As', if any.
		/// </summary>
		public Location AsLocation {
			get { return _AsLocation; }
		}

		/// <summary>
		/// The element type of the enumerated type, if any.
		/// </summary>
		public TypeName ElementType {
			get { return _ElementType; }
		}

		/// <summary>
		/// Constructs a parse tree for an Enum declaration.
		/// </summary>
		/// <param name="attributes">The attributes for the parse tree.</param>
		/// <param name="modifiers">The modifiers for the parse tree.</param>
		/// <param name="keywordLocation">The location of the keyword.</param>
		/// <param name="name">The name of the declaration.</param>
		/// <param name="asLocation">The location of the 'As', if any.</param>
		/// <param name="elementType">The element type of the enumerated type, if any.</param>
		/// <param name="declarations">The enumerated values.</param>
		/// <param name="endStatement">The end block declaration, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public EnumDeclaration(AttributeBlockCollection attributes, ModifierCollection modifiers, Location keywordLocation, SimpleName name, Location asLocation, TypeName elementType, DeclarationCollection declarations, EndBlockDeclaration endStatement, Span span, IList<Comment> comments
		) : base(TreeType.EnumDeclaration, attributes, modifiers, keywordLocation, name, declarations, endStatement, span, comments)
		{

			SetParent(elementType);

			_AsLocation = asLocation;
			_ElementType = elementType;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			base.GetChildTrees(childList);

			AddChild(childList, ElementType);
		}
	}
}

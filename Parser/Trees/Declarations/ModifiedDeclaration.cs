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
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.Declarations
{
	/// <summary>
	/// A parse tree for a declaration with modifiers.
	/// </summary>
	public abstract class ModifiedDeclaration : Declaration
	{

		private readonly AttributeBlockCollection _Attributes;
		private readonly ModifierCollection _Modifiers;

		/// <summary>
		/// The attributes on the declaration.
		/// </summary>
		public AttributeBlockCollection Attributes {
			get { return _Attributes; }
		}

		/// <summary>
		/// The modifiers on the declaration.
		/// </summary>
		public ModifierCollection Modifiers {
			get { return _Modifiers; }
		}

		protected ModifiedDeclaration(TreeType type, AttributeBlockCollection attributes, ModifierCollection modifiers, Span span, IList<Comment> comments) : base(type, span, comments)
		{

			SetParent(attributes);
			SetParent(modifiers);

			_Attributes = attributes;
			_Modifiers = modifiers;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Attributes);
			AddChild(childList, Modifiers);
		}
	}
}

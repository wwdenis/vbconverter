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

namespace VBConverter.CodeParser.Trees.Declarations
{
	/// <summary>
	/// A parse tree for a namespace VBConverter.CodeParser.declaration.
	/// </summary>
	public sealed class NamespaceDeclaration : ModifiedDeclaration
	{

		private readonly Location _NamespaceLocation;
		private readonly Name _Name;
		private readonly DeclarationCollection _Declarations;
		private readonly EndBlockDeclaration _EndDeclaration;

		/// <summary>
		/// The location of 'Namespace'.
		/// </summary>
		public Location NamespaceLocation {
			get { return _NamespaceLocation; }
		}

		/// <summary>
		/// The name of the namespace.
		/// </summary>
		public Name Name {
			get { return _Name; }
		}

		/// <summary>
		/// The declarations in the namespace.
		/// </summary>
		public DeclarationCollection Declarations {
			get { return _Declarations; }
		}

		/// <summary>
		/// The End Namespace declaration, if any.
		/// </summary>
		public EndBlockDeclaration EndDeclaration {
			get { return _EndDeclaration; }
		}

		/// <summary>
		/// Constructs a parse tree for a namespace VBConverter.CodeParser.declaration.
		/// </summary>
		/// <param name="attributes">The attributes on the declaration.</param>
		/// <param name="modifiers">The modifiers on the declaration.</param>
		/// <param name="namespaceLocation">The location of 'Namespace'.</param>
		/// <param name="name">The name of the namespace.</param>
		/// <param name="declarations">The declarations in the namespace.</param>
		/// <param name="endDeclaration">The End Namespace statement, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public NamespaceDeclaration(AttributeBlockCollection attributes, ModifierCollection modifiers, Location namespaceLocation, Name name, DeclarationCollection declarations, EndBlockDeclaration endDeclaration, Span span, IList<Comment> comments) : base(TreeType.NamespaceDeclaration, attributes, modifiers, span, comments)
		{

			if (name == null)
			{
				throw new ArgumentNullException("name");
			}

			SetParent(name);
			SetParent(declarations);
			SetParent(endDeclaration);

			_NamespaceLocation = namespaceLocation;
			_Name = name;
			_Declarations = declarations;
			_EndDeclaration = endDeclaration;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			base.GetChildTrees(childList);

			AddChild(childList, Name);
			AddChild(childList, Declarations);
			AddChild(childList, EndDeclaration);
		}
	}
}

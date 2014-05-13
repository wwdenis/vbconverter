using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.Declarations
{
	/// <summary>
	/// A parse tree for a block declaration.
	/// </summary>
	public abstract class BlockDeclaration : ModifiedDeclaration
	{

		private readonly Location _KeywordLocation;
		private readonly SimpleName _Name;
		private readonly DeclarationCollection _Declarations;
		private readonly EndBlockDeclaration _EndDeclaration;

		/// <summary>
		/// The location of the keyword.
		/// </summary>
		public Location KeywordLocation {
			get { return _KeywordLocation; }
		}

		/// <summary>
		/// The name of the declaration.
		/// </summary>
		public SimpleName Name {
			get { return _Name; }
		}

		/// <summary>
		/// The declarations in the block.
		/// </summary>
		public DeclarationCollection Declarations {
			get { return _Declarations; }
		}

		/// <summary>
		/// The End statement for the block.
		/// </summary>
		public EndBlockDeclaration EndDeclaration {
			get { return _EndDeclaration; }
		}

		protected BlockDeclaration(TreeType type, AttributeBlockCollection attributes, ModifierCollection modifiers, Location keywordLocation, SimpleName name, DeclarationCollection declarations, EndBlockDeclaration endDeclaration, Span span, IList<Comment> comments) : base(type, attributes, modifiers, span, comments)
		{

			Debug.Assert(type == TreeType.ClassDeclaration || type == TreeType.ModuleDeclaration || type == TreeType.InterfaceDeclaration || type == TreeType.StructureDeclaration || type == TreeType.EnumDeclaration);

			if (name == null)
			{
				throw new ArgumentNullException("name");
			}

			SetParent(name);
			SetParent(declarations);
			SetParent(endDeclaration);

			_KeywordLocation = keywordLocation;
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

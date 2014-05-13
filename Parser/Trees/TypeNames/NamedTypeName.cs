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

namespace VBConverter.CodeParser.Trees.TypeNames
{
	/// <summary>
	/// A parse tree for a named type.
	/// </summary>
	public class NamedTypeName : TypeName
	{

		private readonly Name _Name;

		/// <summary>
		/// The name of the type.
		/// </summary>
		public Name Name {
			get { return _Name; }
		}

		/// <summary>
		/// Creates a new bad named type.
		/// </summary>
		/// <param name="span">The location of the bad named type.</param>
		/// <returns>A bad named type.</returns>
		public static NamedTypeName GetBadNamedType(Span span)
		{
			return new NamedTypeName(SimpleName.GetBadSimpleName(span), span);
		}

		public override bool IsBad {
			get { return Name.IsBad; }
		}

		/// <summary>
		/// Constructs a new named type parse tree.
		/// </summary>
		/// <param name="name">The name of the type.</param>
		/// <param name="span">The location of the parse tree.</param>
		public NamedTypeName(Name name, Span span) : this(TreeType.NamedType, name, span)
		{
		}

		protected NamedTypeName(TreeType treeType, Name name, Span span) : base(treeType, span)
		{

			if (name == null)
			{
				throw new ArgumentNullException("name");
			}

			SetParent(name);

			_Name = name;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Name);
		}
	}
}

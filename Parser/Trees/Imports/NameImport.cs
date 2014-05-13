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

namespace VBConverter.CodeParser.Trees.Imports
{
	/// <summary>
	/// A parse tree for an Imports statement for a name.
	/// </summary>
	public sealed class NameImport : Import
	{

		private readonly TypeName _TypeName;

		/// <summary>
		/// The imported name.
		/// </summary>
		public TypeName TypeName {
			get { return _TypeName; }
		}

		/// <summary>
		/// Constructs a new name import parse tree.
		/// </summary>
		/// <param name="typeName">The name to import.</param>
		/// <param name="span">The location of the parse tree.</param>
		public NameImport(TypeName typeName, Span span) : base(TreeType.NameImport, span)
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

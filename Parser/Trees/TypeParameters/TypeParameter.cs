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

namespace VBConverter.CodeParser.Trees.TypeParameters
{
	/// <summary>
	/// A parse tree for a type parameter.
	/// </summary>
	public sealed class TypeParameter : Tree
	{

		private readonly SimpleName _TypeName;
		private readonly Location _AsLocation;
		private readonly TypeConstraintCollection _TypeConstraints;

		/// <summary>
		/// The name of the type parameter.
		/// </summary>
		public SimpleName TypeName {
			get { return _TypeName; }
		}

		/// <summary>
		/// The location of the 'As', if any.
		/// </summary>
		public Location AsLocation {
			get { return _AsLocation; }
		}

		/// <summary>
		/// The constraints, if any.
		/// </summary>
		public TypeConstraintCollection TypeConstraints {
			get { return _TypeConstraints; }
		}

		/// <summary>
		/// Constructs a new parameter parse tree.
		/// </summary>
		/// <param name="typeName">The name of the type parameter.</param>
		/// <param name="asLocation">The location of the 'As'.</param>
		/// <param name="typeConstraints">The constraints on the type parameter. Can be Nothing.</param>
		/// <param name="span">The location of the parse tree.</param>
		public TypeParameter(SimpleName typeName, Location asLocation, TypeConstraintCollection typeConstraints, Span span) : base(TreeType.TypeParameter, span)
		{

			if (typeName == null)
			{
				throw new ArgumentNullException("typeName");
			}

			SetParent(typeName);
			SetParent(typeConstraints);

			_TypeName = typeName;
			_AsLocation = asLocation;
			_TypeConstraints = typeConstraints;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, TypeName);
			AddChild(childList, TypeConstraints);
		}
	}
}

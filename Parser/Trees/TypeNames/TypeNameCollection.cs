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
using VBConverter.CodeParser.Trees.Collections;

namespace VBConverter.CodeParser.Trees.TypeNames
{
	/// <summary>
	/// A read-only collection of type names.
	/// </summary>
	public sealed class TypeNameCollection : CommaDelimitedTreeCollection<TypeName>
	{

		/// <summary>
		/// Constructs a new type name collection.
		/// </summary>
		/// <param name="typeMembers">The type names in the collection.</param>
		/// <param name="commaLocations">The locations of the commas in the collection.</param>
		/// <param name="span">The location of the parse tree.</param>
		public TypeNameCollection(IList<TypeName> typeMembers, IList<Location> commaLocations, Span span) : base(TreeType.TypeNameCollection, typeMembers, commaLocations, span)
		{

			if (typeMembers == null || typeMembers.Count == 0)
			{
				throw new ArgumentException("TypeNameCollection cannot be empty.");
			}
		}
	}
}

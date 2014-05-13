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
using VBConverter.CodeParser.Trees.TypeNames;

namespace VBConverter.CodeParser.Trees.TypeParameters
{
	/// <summary>
	/// A collection of type constraints.
	/// </summary>
	public sealed class TypeConstraintCollection : CommaDelimitedTreeCollection<TypeName>
	{

		private readonly Location _RightBracketLocation;

		/// <summary>
		/// The location of the '}', if any.
		/// </summary>
		public Location RightBracketLocation {
			get { return _RightBracketLocation; }
		}

		/// <summary>
		/// Constructs a new collection of type constraints.
		/// </summary>
		/// <param name="constraints">The type constraints in the collection</param>
		/// <param name="commaLocations">The locations of the commas.</param>
		/// <param name="rightBracketLocation">The location of the right bracket, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		public TypeConstraintCollection(IList<TypeName> constraints, IList<Location> commaLocations, Location rightBracketLocation, Span span) : base(TreeType.TypeConstraintCollection, constraints, commaLocations, span)
		{

			_RightBracketLocation = rightBracketLocation;
		}
	}
}

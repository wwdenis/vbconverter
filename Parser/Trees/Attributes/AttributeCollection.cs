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
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.Attributes
{
	/// <summary>
	/// A read-only collection of attributes.
	/// </summary>
	public sealed class AttributeCollection : CommaDelimitedTreeCollection<AttributeTree>
	{

		private readonly Location _RightBracketLocation;

		/// <summary>
		/// The location of the '}'.
		/// </summary>
		public Location RightBracketLocation {
			get { return _RightBracketLocation; }
		}

		/// <summary>
		/// Constructs a new collection of attributes.
		/// </summary>
		/// <param name="attributes">The attributes in the collection.</param>
		/// <param name="commaLocations">The location of the commas in the list.</param>
		/// <param name="rightBracketLocation">The location of the right bracket.</param>
		/// <param name="span">The location of the parse tree.</param>
		public AttributeCollection(IList<AttributeTree> attributes, IList<Location> commaLocations, Location rightBracketLocation, Span span) : base(TreeType.AttributeCollection, attributes, commaLocations, span)
		{

			if (attributes == null || attributes.Count == 0)
			{
				throw new ArgumentException("AttributeCollection cannot be empty.");
			}

			_RightBracketLocation = rightBracketLocation;
		}
	}
}

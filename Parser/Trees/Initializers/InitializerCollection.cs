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

namespace VBConverter.CodeParser.Trees.Initializers
{
	/// <summary>
	/// A read-only collection of initializers.
	/// </summary>
	public sealed class InitializerCollection : CommaDelimitedTreeCollection<Initializer>
	{
		private readonly Location _RightCurlyBraceLocation;

		/// <summary>
		/// The location of the '}'.
		/// </summary>
		public Location RightCurlyBraceLocation {
			get { return _RightCurlyBraceLocation; }
		}
		/// <summary>
		/// Constructs a new initializer collection.
		/// </summary>
		/// <param name="initializers">The initializers in the collection.</param>
		/// <param name="commaLocations">The locations of the commas in the collection.</param>
		/// <param name="rightCurlyBraceLocation">The location of the '}'.</param>
		/// <param name="span">The location of the parse tree.</param>
		public InitializerCollection(IList<Initializer> initializers, IList<Location> commaLocations, Location rightCurlyBraceLocation, Span span) : base(TreeType.InitializerCollection, initializers, commaLocations, span)
		{

			_RightCurlyBraceLocation = rightCurlyBraceLocation;
		}
	}
}

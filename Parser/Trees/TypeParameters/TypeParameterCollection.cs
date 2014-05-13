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

namespace VBConverter.CodeParser.Trees.TypeParameters
{
	/// <summary>
	/// A collection of type parameters.
	/// </summary>
	public sealed class TypeParameterCollection : CommaDelimitedTreeCollection<TypeParameter>
	{

		private readonly Location _OfLocation;
		private readonly Location _RightParenthesisLocation;

		/// <summary>
		/// The location of the 'Of'.
		/// </summary>
		public Location OfLocation {
			get { return _OfLocation; }
		}

		/// <summary>
		/// The location of the ')'.
		/// </summary>
		public Location RightParenthesisLocation {
			get { return _RightParenthesisLocation; }
		}

		/// <summary>
		/// Constructs a new collection of type parameters.
		/// </summary>
		/// <param name="ofLocation">The location of the 'Of'.</param>
		/// <param name="parameters">The type parameters in the collection</param>
		/// <param name="commaLocations">The locations of the commas.</param>
		/// <param name="rightParenthesisLocation">The location of the right parenthesis.</param>
		/// <param name="span">The location of the parse tree.</param>
		public TypeParameterCollection(Location ofLocation, IList<TypeParameter> parameters, IList<Location> commaLocations, Location rightParenthesisLocation, Span span) : base(TreeType.TypeParameterCollection, parameters, commaLocations, span)
		{

			_OfLocation = ofLocation;
			_RightParenthesisLocation = rightParenthesisLocation;
		}
	}
}

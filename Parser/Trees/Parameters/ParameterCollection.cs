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

namespace VBConverter.CodeParser.Trees.Parameters
{
	/// <summary>
	/// A collection of parameters.
	/// </summary>
	public sealed class ParameterCollection : CommaDelimitedTreeCollection<Parameter>
	{

		private readonly Location _RightParenthesisLocation;

		/// <summary>
		/// The location of the ')'.
		/// </summary>
		public Location RightParenthesisLocation {
			get { return _RightParenthesisLocation; }
		}

		/// <summary>
		/// Constructs a new collection of parameters.
		/// </summary>
		/// <param name="parameters">The parameters in the collection</param>
		/// <param name="commaLocations">The locations of the commas.</param>
		/// <param name="rightParenthesisLocation">The location of the right parenthesis.</param>
		/// <param name="span">The location of the parse tree.</param>
		public ParameterCollection(IList<Parameter> parameters, IList<Location> commaLocations, Location rightParenthesisLocation, Span span) : base(TreeType.ParameterCollection, parameters, commaLocations, span)
		{

			_RightParenthesisLocation = rightParenthesisLocation;
		}
	}
}

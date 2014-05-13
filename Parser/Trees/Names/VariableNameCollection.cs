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

namespace VBConverter.CodeParser.Trees.Names
{
	/// <summary>
	/// A read-only collection of variable names.
	/// </summary>
	public sealed class VariableNameCollection : CommaDelimitedTreeCollection<VariableName>
	{
		/// <summary>
		/// Constructs a new variable name collection.
		/// </summary>
		/// <param name="variableNames">The variable names in the collection.</param>
		/// <param name="commaLocations">The locations of the commas in the collection.</param>
		/// <param name="span">The location of the parse tree.</param>
		public VariableNameCollection(IList<VariableName> variableNames, IList<Location> commaLocations, Span span) : base(TreeType.VariableNameCollection, variableNames, commaLocations, span)
		{

			if (variableNames == null || variableNames.Count == 0)
			{
				throw new ArgumentException("VariableNameCollection cannot be empty.");
			}
		}
	}
}

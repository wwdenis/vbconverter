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

namespace VBConverter.CodeParser.Trees.Imports
{
	/// <summary>
	/// A read-only collection of imports.
	/// </summary>
	public sealed class ImportCollection : CommaDelimitedTreeCollection<Import>
	{

		/// <summary>
		/// Constructs a collection of imports.
		/// </summary>
		/// <param name="importMembers">The imports in the collection.</param>
		/// <param name="commaLocations">The location of the commas.</param>
		/// <param name="span">The location of the parse tree.</param>
		public ImportCollection(IList<Import> importMembers, IList<Location> commaLocations, Span span) : base(TreeType.ImportCollection, importMembers, commaLocations, span)
		{

			if (importMembers == null || importMembers.Count == 0)
			{
				throw new ArgumentException("ImportCollection cannot be empty.");
			}
		}
	}
}

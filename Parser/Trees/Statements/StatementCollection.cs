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

namespace VBConverter.CodeParser.Trees.Statements
{
	/// <summary>
	/// A read-only collection of statements.
	/// </summary>
	public sealed class StatementCollection : ColonDelimitedTreeCollection<Statement>
	{

		/// <summary>
		/// Constructs a new collection of statements.
		/// </summary>
		/// <param name="statements">The statements in the collection.</param>
		/// <param name="colonLocations">The locations of the colons in the collection.</param>
		/// <param name="span">The location of the parse tree.</param>
		public StatementCollection(IList<Statement> statements, IList<Location> colonLocations, Span span) : base(TreeType.StatementCollection, statements, colonLocations, span)
		{

			if ((statements == null || statements.Count == 0) && (colonLocations == null || colonLocations.Count == 0))
			{
				throw new ArgumentException("StatementCollection cannot be empty.");
			}
		}
	}
}

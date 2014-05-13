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

namespace VBConverter.CodeParser.Trees.CaseClause
{
	/// <summary>
	/// A collection of case clauses.
	/// </summary>
	public sealed class CaseClauseCollection : CommaDelimitedTreeCollection<CaseClause>
	{

		/// <summary>
		/// Constructs a new collection of case clauses.
		/// </summary>
		/// <param name="caseClauses">The case clauses in the collection.</param>
		/// <param name="commaLocations">The locations of the commas in the list.</param>
		/// <param name="span">The location of the parse tree.</param>
		public CaseClauseCollection(IList<CaseClause> caseClauses, IList<Location> commaLocations, Span span) : base(TreeType.CaseClauseCollection, caseClauses, commaLocations, span)
		{

			if (caseClauses == null || caseClauses.Count == 0)
			{
				throw new ArgumentException("CaseClauseCollection cannot be empty.");
			}
		}
	}
}

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using System.Diagnostics;
//
// Visual Basic .NET Parser
//
// Copyright (C) 2005, Microsoft Corporation. All rights reserved.
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
// EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
// MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
//
namespace VBConverter.CodeParser.Trees.Collections
{
	/// <summary>
	/// A collection of trees that are colon delimited.
	/// </summary>
	public abstract class ColonDelimitedTreeCollection<T> : TreeCollection<T> where T : Tree
	{

		private readonly ReadOnlyCollection<Location> _ColonLocations;

		/// <summary>
		/// The locations of the colons in the collection.
		/// </summary>
		public ReadOnlyCollection<Location> ColonLocations {
			get { return _ColonLocations; }
		}

		protected ColonDelimitedTreeCollection(TreeType type, IList<T> trees, IList<Location> colonLocations, Span span) : base(type, trees, span)
		{

			Debug.Assert(type == TreeType.StatementCollection || type == TreeType.DeclarationCollection);

			if (colonLocations != null && colonLocations.Count > 0)
			{
				_ColonLocations = new ReadOnlyCollection<Location>(colonLocations);
			}
		}
	}
}

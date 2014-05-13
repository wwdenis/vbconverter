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
namespace VBConverter.CodeParser.Context
{
	/// <summary>
	/// A region marked in the source code.
	/// </summary>
	public sealed class SourceRegion
	{
		private readonly Location _Start;
		private readonly Location _Finish;
		private readonly string _Description;

		/// <summary>
		/// The start location of the region.
		/// </summary>
		public Location Start {
			get { return _Start; }
		}

		/// <summary>
		/// The end location of the region.
		/// </summary>
		public Location Finish {
			get { return _Finish; }
		}

		/// <summary>
		/// The description of the region.
		/// </summary>
		public string Description {
			get { return _Description; }
		}

		/// <summary>
		/// Constructs a new source region.
		/// </summary>
		/// <param name="start">The start location of the region.</param>
		/// <param name="finish">The end location of the region.</param>
		/// <param name="description">The description of the region.</param>
		public SourceRegion(Location start, Location finish, string description)
		{
			_Start = start;
			_Finish = finish;
			_Description = description;
		}
	}
}

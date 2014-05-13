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
	/// A line mapping from a source span to an external file and line.
	/// </summary>
	public sealed class ExternalLineMapping
	{
		private readonly Location _Start;
		private readonly Location _Finish;
		private readonly string _File;
		private readonly long _Line;

		/// <summary>
		/// The start location of the mapping in the source.
		/// </summary>
		public Location Start {
			get { return _Start; }
		}

		/// <summary>
		/// The end location of the mapping in the source.
		/// </summary>
		public Location Finish {
			get { return _Finish; }
		}

		/// <summary>
		/// The external file the source maps to.
		/// </summary>
		public string File {
			get { return _File; }
		}

		/// <summary>
		/// The external line number the source maps to.
		/// </summary>
		public long Line {
			get { return _Line; }
		}

		/// <summary>
		/// Constructs a new external line mapping.
		/// </summary>
		/// <param name="start">The start location in the source.</param>
		/// <param name="finish">The end location in the source.</param>
		/// <param name="file">The name of the external file.</param>
		/// <param name="line">The line number in the external file.</param>
		public ExternalLineMapping(Location start, Location finish, string file, long line)
		{
			_Start = start;
			_Finish = finish;
			_File = file;
			_Line = line;
		}
	}
}

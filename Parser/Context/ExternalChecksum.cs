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
	/// An external checksum for a file.
	/// </summary>
	public sealed class ExternalChecksum
	{
		private readonly string _Filename;
		private readonly string _Guid;
		private readonly string _Checksum;

		/// <summary>
		/// The filename that the checksum is for.
		/// </summary>
		public string Filename {
			get { return _Filename; }
		}

		/// <summary>
		/// The guid of the file.
		/// </summary>
		public string Guid {
			get { return _Guid; }
		}

		/// <summary>
		/// The checksum for the file.
		/// </summary>
		public string Checksum {
			get { return _Checksum; }
		}

		/// <summary>
		/// Constructs a new external checksum.
		/// </summary>
		/// <param name="filename">The filename that the checksum is for.</param>
		/// <param name="guid">The guid of the file.</param>
		/// <param name="checksum">The checksum for the file.</param>
		public ExternalChecksum(string filename, string guid, string checksum)
		{
			_Filename = filename;
			_Guid = guid;
			_Checksum = checksum;
		}
	}
}

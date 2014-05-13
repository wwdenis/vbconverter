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
using VBConverter.CodeParser.Trees.Arguments;
using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Names;

namespace VBConverter.CodeParser.Trees.Statements
{
	/// <summary>
	/// A parse tree for a RaiseEvent statement.
	/// </summary>
	public sealed class RaiseEventStatement : Statement
	{

		private readonly SimpleName _Name;
		private readonly ArgumentCollection _Arguments;

		/// <summary>
		/// The name of the event to raise.
		/// </summary>
		public SimpleName Name {
			get { return _Name; }
		}

		/// <summary>
		/// The arguments to the event.
		/// </summary>
		public ArgumentCollection Arguments {
			get { return _Arguments; }
		}

		/// <summary>
		/// Constructs a new parse tree for a RaiseEvents statement.
		/// </summary>
		/// <param name="name">The name of the event to raise.</param>
		/// <param name="arguments">The arguments to the event.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public RaiseEventStatement(SimpleName name, ArgumentCollection arguments, Span span, IList<Comment> comments) : base(TreeType.RaiseEventStatement, span, comments)
		{

			if (name == null)
			{
				throw new ArgumentNullException("name");
			}

			SetParent(name);
			SetParent(arguments);

			_Name = name;
			_Arguments = arguments;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Name);
			AddChild(childList, Arguments);
		}
	}
}

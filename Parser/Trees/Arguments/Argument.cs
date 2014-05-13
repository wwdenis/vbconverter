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
using VBConverter.CodeParser.Trees.Expressions;
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.Arguments
{
	/// <summary>
	/// A parse tree for an argument to a call or index.
	/// </summary>
	public sealed class Argument : Tree
	{
        private readonly SimpleName _Name;
		private readonly Location _ColonEqualsLocation;
		private readonly Expression _Expression;
        private readonly Location _ByValLocation = Location.Empty;


		/// <summary>
		/// The name of the argument, if any.
		/// </summary>
		public SimpleName Name {
			get { return _Name; }
		}

		/// <summary>
		/// The location of the ':=', if any.
		/// </summary>
		public Location ColonEqualsLocation {
			get { return _ColonEqualsLocation; }
		}

		/// <summary>
		/// The argument, if any.
		/// </summary>
		public Expression Expression {
			get { return _Expression; }
		}

        /// <summary>
        /// The location of ByVal, if any (VB6).
        /// </summary>
        public Location ByValLocation
        {
            get { return _ByValLocation; }
        }

		/// <summary>
		/// Constructs a new parse tree for an argument.
		/// </summary>
		/// <param name="name">The name of the argument, if any.</param>
		/// <param name="colonEqualsLocation">The location of the ':=', if any.</param>
		/// <param name="expression">The expression, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
        /// <param name="byValLocation">The location of ByVal, if any (VB6).</param>
		public Argument(SimpleName name, Location colonEqualsLocation, Expression expression, Span span, Location byValLocation) : base(TreeType.Argument, span)
		{
            if (expression == null)
				throw new ArgumentNullException("expression");
			
			SetParent(name);
			SetParent(expression);

			_Name = name;
			_ColonEqualsLocation = colonEqualsLocation;
			_Expression = expression;
            _ByValLocation = byValLocation;
		}

        /// <summary>
        /// Constructs a new parse tree for an argument.
        /// </summary>
        /// <param name="name">The name of the argument, if any.</param>
        /// <param name="colonEqualsLocation">The location of the ':=', if any.</param>
        /// <param name="expression">The expression, if any.</param>
        /// <param name="span">The location of the parse tree.</param>
        public Argument(SimpleName name, Location colonEqualsLocation, Expression expression, Span span) : this(name, colonEqualsLocation, expression, span, Location.Empty)
        {
        }

		private Argument() : base(TreeType.Argument, Span.Empty)
		{
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Name);
			AddChild(childList, Expression);
		}
	}
}

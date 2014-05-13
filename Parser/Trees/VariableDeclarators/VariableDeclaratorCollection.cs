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
using VBConverter.CodeParser.Trees.Declarations.Members;
using VBConverter.CodeParser.Trees.Statements;

namespace VBConverter.CodeParser.Trees.VariableDeclarators
{
	/// <summary>
	/// A read-only collection of variable declarators.
	/// </summary>
	public sealed class VariableDeclaratorCollection : CommaDelimitedTreeCollection<VariableDeclarator>
    {
        #region Fields

        public bool IsField
        {
            get 
            {
                bool isField = false;

                if (Parent is VariableListDeclaration)
                    isField = true;
                else if (Parent is LocalDeclarationStatement)
                    isField = false;

                return isField; 
            }
        }

        #endregion

        #region Constructor

        /// <summary>
		/// Constructs a new collection of variable declarators.
		/// </summary>
		/// <param name="variableDeclarators">The variable declarators in the collection.</param>
		/// <param name="commaLocations">The locations of the commas in the list.</param>
		/// <param name="span">The location of the parse tree.</param>
		public VariableDeclaratorCollection(IList<VariableDeclarator> variableDeclarators, IList<Location> commaLocations, Span span) : base(TreeType.VariableDeclaratorCollection, variableDeclarators, commaLocations, span)
		{
			if (variableDeclarators == null || variableDeclarators.Count == 0)
				throw new ArgumentException("VariableDeclaratorCollection cannot be empty.");
        }

        #endregion
    }
}

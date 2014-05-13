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
using VBConverter.CodeParser.Trees.Initializers;
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.TypeNames;


namespace VBConverter.CodeParser.Trees.VariableDeclarators
{
    /// <summary>
    /// A parse tree for a variable declarator (e.g. "x As Integer")
    /// </summary>
    public sealed class VariableDeclarator : Tree
    {
        #region Fields

        private readonly VariableNameCollection _VariableNames;
        private readonly Location _AsLocation;
        private readonly Location _NewLocation;
        private readonly TypeName _VariableType;
        private readonly ArgumentCollection _Arguments;
        private readonly Location _EqualsLocation;
        private readonly Initializer _Initializer;
        
        #endregion

        #region Properties

        /// <summary>
        /// The variable names being declared.
        /// </summary>
        public VariableNameCollection VariableNames
        {
            get { return _VariableNames; }
        }

        /// <summary>
        /// The location of the 'As', if any.
        /// </summary>
        public Location AsLocation
        {
            get { return _AsLocation; }
        }

        /// <summary>
        /// The location of the 'New', if any.
        /// </summary>
        public Location NewLocation
        {
            get { return _NewLocation; }
        }

        /// <summary>
        /// The type of the variables being declared, if any.
        /// </summary>
        public TypeName VariableType
        {
            get { return _VariableType; }
        }

        /// <summary>
        /// The arguments to the constructor, if any.
        /// </summary>
        public ArgumentCollection Arguments
        {
            get { return _Arguments; }
        }

        /// <summary>
        /// The location of the '=', if any.
        /// </summary>
        public Location EqualsLocation
        {
            get { return _EqualsLocation; }
        }

        /// <summary>
        /// The variable initializer, if any.
        /// </summary>
        public Initializer Initializer
        {
            get { return _Initializer; }
        }

        public bool IsField
        {
            get 
            { 
                bool isField = false;

                if (Parent is VariableDeclaratorCollection)
                    isField = ((VariableDeclaratorCollection)Parent).IsField;

                return isField;
            }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Constructs a new parse tree for a variable declarator.
        /// </summary>
        /// <param name="variableNames">The names of the variables being declared.</param>
        /// <param name="asLocation">The location of the 'As', if any.</param>
        /// <param name="newLocation">The location of the 'New', if any.</param>
        /// <param name="variableType">The type of the variables being declared, if any.</param>
        /// <param name="arguments">The arguments of the constructor, if any.</param>
        /// <param name="equalsLocation">The location of the '=', if any.</param>
        /// <param name="initializer">The variable initializer, if any.</param>
        /// <param name="span">The location of the parse tree.</param>
        /// <param name="isField">If the variable declarator is a field or not.</param>
        public VariableDeclarator(VariableNameCollection variableNames, Location asLocation, Location newLocation, TypeName variableType, ArgumentCollection arguments, Location equalsLocation, Initializer initializer, Span span) : base(TreeType.VariableDeclarator, span)
        {
            if (variableNames == null)
                throw new ArgumentNullException("variableNames");
         
            SetParent(variableNames);
            SetParent(variableType);
            SetParent(arguments);
            SetParent(initializer);

            _VariableNames = variableNames;
            _AsLocation = asLocation;
            _NewLocation = newLocation;
            _VariableType = variableType;
            _Arguments = arguments;
            _EqualsLocation = equalsLocation;
            _Initializer = initializer;
        }

        #endregion

        #region Methods

        protected override void GetChildTrees(IList<Tree> childList)
        {
            AddChild(childList, VariableNames);
            AddChild(childList, VariableType);
            AddChild(childList, Arguments);
            AddChild(childList, Initializer);
        }

        #endregion
    }
}

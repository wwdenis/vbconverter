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
using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Modifiers;
using VBConverter.CodeParser.Trees.VariableDeclarators;

namespace VBConverter.CodeParser.Trees.Statements
{
    /// <summary>
    /// A parse tree for a local declaration statement.
    /// </summary>
    public sealed class LocalDeclarationStatement : Statement
    {

        private readonly ModifierCollection _Modifiers;
        private readonly VariableDeclaratorCollection _VariableDeclarators;

        /// <summary>
        /// The statement modifiers.
        /// </summary>
        public ModifierCollection Modifiers
        {
            get { return _Modifiers; }
        }

        /// <summary>
        /// The variable declarators.
        /// </summary>
        public VariableDeclaratorCollection VariableDeclarators
        {
            get { return _VariableDeclarators; }
        }

        /// <summary>
        /// Constructs a new parse tree for a local declaration statement.
        /// </summary>
        /// <param name="modifiers">The statement modifiers.</param>
        /// <param name="variableDeclarators">The variable declarators.</param>
        /// <param name="span">The location of the parse tree.</param>
        /// <param name="comments">The comments for the parse tree.</param>
        public LocalDeclarationStatement(ModifierCollection modifiers, VariableDeclaratorCollection variableDeclarators, Span span, IList<Comment> comments) : base(TreeType.LocalDeclarationStatement, span, comments)
        {
            if (modifiers == null)
                throw new ArgumentNullException("modifers");
         
            if (variableDeclarators == null)
                throw new ArgumentNullException("variableDeclarators");
         
            SetParent(modifiers);
            SetParent(variableDeclarators);

            _Modifiers = modifiers;
            _VariableDeclarators = variableDeclarators;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            AddChild(childList, Modifiers);
            AddChild(childList, VariableDeclarators);
        }
    }
}

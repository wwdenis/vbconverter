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
using VBConverter.CodeParser.Trees.Declarations;
using VBConverter.CodeParser.Trees.Declarations.Members;
using VBConverter.CodeParser.Trees.Statements;
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.VariableDeclarators;

namespace VBConverter.CodeParser.Trees.Expressions
{
    /// <summary>
    /// A parse tree for a call or index expression.
    /// </summary>
    public sealed class CallOrIndexExpression : Expression
    {
        private readonly Expression _TargetExpression;
        private readonly ArgumentCollection _Arguments;

        /// <summary>
        /// The target of the call or index.
        /// </summary>
        public Expression TargetExpression
        {
            get { return _TargetExpression; }
        }

        /// <summary>
        /// The arguments to the call or index.
        /// </summary>
        public ArgumentCollection Arguments
        {
            get { return _Arguments; }
        }

        public bool IsIndex
        {
            get
            {
                return Declarator != null;
            }
        }

        public VariableDeclarator Declarator
        {
            get
            {
                Tree parentTree = Parent;
                SimpleName targetName = null;

                if (_TargetExpression is QualifiedExpression)
                    return null; 
                else if (_TargetExpression is SimpleNameExpression)
                    targetName = ((SimpleNameExpression)_TargetExpression).Name;
                                
                while (parentTree != null)
                {
                    foreach (Tree item in parentTree.Children)
                    {
                        VariableDeclaratorCollection declarators = null;
                        if (item is LocalDeclarationStatement)
                            declarators = ((LocalDeclarationStatement)item).VariableDeclarators;
                        else if (item is VariableListDeclaration)
                            declarators = ((VariableListDeclaration)item).VariableDeclarators;

                        if (declarators != null)
                        {
                            foreach (VariableDeclarator declarator in declarators)
                            {
                                foreach (VariableName variable in declarator.VariableNames)
                                {
                                    if (variable.Name == targetName)
                                        return declarator;
                                }
                            }
                        }
                    }

                    parentTree = parentTree.Parent;
                }

                return null;
            }
        }

        /// <summary>
        /// Constructs a new parse tree for a call or index expression.
        /// </summary>
        /// <param name="targetExpression">The target of the call or index.</param>
        /// <param name="arguments">The arguments to the call or index.</param>
        /// <param name="span">The location of the parse tree.</param>
        public CallOrIndexExpression(Expression targetExpression, ArgumentCollection arguments, Span span) : base(TreeType.CallOrIndexExpression, span)
        {
            if (targetExpression == null)
                throw new ArgumentNullException("targetExpression");

            SetParent(targetExpression);
            SetParent(arguments);

            _TargetExpression = targetExpression;
            _Arguments = arguments;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            AddChild(childList, TargetExpression);
            AddChild(childList, Arguments);
        }
    }
}

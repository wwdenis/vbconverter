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
using VBConverter.CodeParser.Trees.Expressions;
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.VariableDeclarators;
using VBConverter.CodeParser.Trees.Declarations.Members;

namespace VBConverter.CodeParser.Trees.Statements
{
    /// <summary>
    /// A parse tree for a ReDim statement.
    /// </summary>
    public sealed class ReDimStatement : Statement
    {
        private readonly Location _PreserveLocation;
        private readonly ExpressionCollection _Variables;

        /// <summary>
        /// The location of the 'Preserve', if any.
        /// </summary>
        public Location PreserveLocation
        {
            get { return _PreserveLocation; }
        }

        /// <summary>
        /// The variables to redimension (includes bounds).
        /// </summary>
        public ExpressionCollection Variables
        {
            get { return _Variables; }
        }

        /// <summary>
        /// Whether the statement included a Preserve keyword.
        /// </summary>
        public bool IsPreserve
        {
            get { return PreserveLocation.IsValid; }
        }

        public VariableDeclaratorCollection VariableDeclarators
        {
            get
            {
                List<VariableDeclarator> list = new List<VariableDeclarator>();
                VariableDeclaratorCollection result = null;

                Tree parentTree = Parent;

                while (parentTree != null)
                {
                    foreach (Tree item in parentTree.Children)
                    {
                        VariableDeclaratorCollection declarations = null;
                        if (item is VariableListDeclaration)
                            declarations = ((VariableListDeclaration)item).VariableDeclarators;
                        else if (item is LocalDeclarationStatement)
                            declarations = ((LocalDeclarationStatement)item).VariableDeclarators;

                        if (declarations != null)
                        {
                            foreach (VariableDeclarator declarator in declarations)
                            {
                                foreach (VariableName variable in declarator.VariableNames)
                                {
                                    foreach (Expression expression in Variables)
                                    {
                                        foreach (Tree child in expression.Children)
                                        {
                                            if (child is SimpleNameExpression && ((SimpleNameExpression)child).Name == variable.Name)
                                                list.Add(declarator);
                                            if (child is QualifiedExpression && ((QualifiedExpression)child).Name == variable.Name)
                                                return null;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    parentTree = parentTree.Parent;
                }

                if (list.Count > 0)
                    result = new VariableDeclaratorCollection(list, null, Span.Empty);

                return result;
            }
        }


        /// <summary>
        /// Constructs a new parse tree for a ReDim statement.
        /// </summary>
        /// <param name="preserveLocation">The location of the 'Preserve', if any.</param>
        /// <param name="variables">The variables to redimension (includes bounds).</param>
        /// <param name="span">The location of the parse tree.</param>
        /// <param name="comments">The comments for the parse tree.</param>
        public ReDimStatement(Location preserveLocation, ExpressionCollection variables, Span span, IList<Comment> comments) : base(TreeType.ReDimStatement, span, comments)
        {

            if (variables == null)
            {
                throw new ArgumentNullException("variables");
            }

            SetParent(variables);
            _PreserveLocation = preserveLocation;
            _Variables = variables;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            AddChild(childList, Variables);
        }
    }
}

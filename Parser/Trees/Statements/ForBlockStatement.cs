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
using VBConverter.CodeParser.Trees.Declarations.Members;
using VBConverter.CodeParser.Trees.Expressions;
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.VariableDeclarators;



namespace VBConverter.CodeParser.Trees.Statements
{
    /// <summary>
    /// A parse tree for a For statement.
    /// </summary>
    public sealed class ForBlockStatement : BlockStatement
    {
        private readonly Expression _ControlExpression;
        private readonly VariableDeclarator _ControlVariableDeclarator;
        private readonly Location _EqualsLocation;
        private readonly Expression _LowerBoundExpression;
        private readonly Location _ToLocation;
        private readonly Expression _UpperBoundExpression;
        private readonly Location _StepLocation;
        private readonly Expression _StepExpression;
        private readonly NextStatement _NextStatement;

        /// <summary>
        /// The control expression for the loop.
        /// </summary>
        public Expression ControlExpression
        {
            get { return _ControlExpression; }
        }

        /// <summary>
        /// The control variable declarator, if any.
        /// </summary>
        public VariableDeclarator ControlVariableDeclarator
        {
            get
            {
                if (_ControlVariableDeclarator == null)
                {
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
                                        if (ControlExpression is SimpleNameExpression && ((SimpleNameExpression)ControlExpression).Name == variable.Name)
                                            return declarator;
                                        else if (ControlExpression is QualifiedExpression && ((QualifiedExpression)ControlExpression).Name == variable.Name)
                                            return null;
                                    }
                                }
                            }
                        }

                        parentTree = parentTree.Parent;
                    }
                }

                return _ControlVariableDeclarator;
            }
        }

        /// <summary>
        /// The location of the '='.
        /// </summary>
        public Location EqualsLocation
        {
            get { return _EqualsLocation; }
        }

        /// <summary>
        /// The lower bound of the loop.
        /// </summary>
        public Expression LowerBoundExpression
        {
            get { return _LowerBoundExpression; }
        }

        /// <summary>
        /// The location of the 'To'.
        /// </summary>
        public Location ToLocation
        {
            get { return _ToLocation; }
        }

        /// <summary>
        /// The upper bound of the loop.
        /// </summary>
        public Expression UpperBoundExpression
        {
            get { return _UpperBoundExpression; }
        }

        /// <summary>
        /// The location of the 'Step', if any.
        /// </summary>
        public Location StepLocation
        {
            get { return _StepLocation; }
        }

        /// <summary>
        /// The step of the loop, if any.
        /// </summary>
        public Expression StepExpression
        {
            get { return _StepExpression; }
        }

        /// <summary>
        /// The Next statement, if any.
        /// </summary>
        public NextStatement NextStatement
        {
            get { return _NextStatement; }
        }

        /// <summary>
        /// Constructs a new parse tree for a For statement.
        /// </summary>
        /// <param name="controlExpression">The control expression for the loop.</param>
        /// <param name="controlVariableDeclarator">The control variable declarator, if any.</param>
        /// <param name="equalsLocation">The location of the '='.</param>
        /// <param name="lowerBoundExpression">The lower bound of the loop.</param>
        /// <param name="toLocation">The location of the 'To'.</param>
        /// <param name="upperBoundExpression">The upper bound of the loop.</param>
        /// <param name="stepLocation">The location of the 'Step', if any.</param>
        /// <param name="stepExpression">The step of the loop, if any.</param>
        /// <param name="statements">The statements in the For loop.</param>
        /// <param name="nextStatement">The Next statement, if any.</param>
        /// <param name="span">The location of the parse tree.</param>
        /// <param name="comments">The comments for the parse tree.</param>
        public ForBlockStatement(Expression controlExpression, VariableDeclarator controlVariableDeclarator, Location equalsLocation, Expression lowerBoundExpression, Location toLocation, Expression upperBoundExpression, Location stepLocation, Expression stepExpression, StatementCollection statements, NextStatement nextStatement, Span span, IList<Comment> comments) : base(TreeType.ForBlockStatement, statements, span, comments)
        {
            if (controlExpression == null)
                throw new ArgumentNullException("controlExpression");

            SetParent(controlExpression);
            SetParent(controlVariableDeclarator);
            SetParent(lowerBoundExpression);
            SetParent(upperBoundExpression);
            SetParent(stepExpression);
            SetParent(nextStatement);

            _ControlExpression = controlExpression;
            _ControlVariableDeclarator = controlVariableDeclarator;
            _EqualsLocation = equalsLocation;
            _LowerBoundExpression = lowerBoundExpression;
            _ToLocation = toLocation;
            _UpperBoundExpression = upperBoundExpression;
            _StepLocation = stepLocation;
            _StepExpression = stepExpression;
            _NextStatement = nextStatement;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            AddChild(childList, ControlExpression);
            AddChild(childList, ControlVariableDeclarator);
            AddChild(childList, LowerBoundExpression);
            AddChild(childList, UpperBoundExpression);
            AddChild(childList, StepExpression);
            base.GetChildTrees(childList);
            AddChild(childList, NextStatement);
        }
    }
}

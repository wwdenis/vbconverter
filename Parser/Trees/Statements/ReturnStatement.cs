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

namespace VBConverter.CodeParser.Trees.Statements
{
	/// <summary>
	/// A parse tree for an expression statement.
	/// </summary>
	public sealed class ReturnStatement : ExpressionStatement
	{
        private bool _isLabelReference;

        public bool IsLabelReference
        {
            get { return _isLabelReference; }
            set { _isLabelReference = value; }
        }

        public MethodDeclaration MethodReference
        {
            get
            {
                // Look for the parent method declaration
                Tree parentTree = this;
                do
                {
                    parentTree = parentTree.Parent;

                    if (parentTree is MethodDeclaration)
                        return (MethodDeclaration)parentTree;
                }
                while (parentTree != null);

                return null;
            }
        }

        public LabelStatement LabelReference
        {
            get
            {
                if (!IsLabelReference)
                    return null;

                MethodDeclaration method = MethodReference;

                if (method == null)
                    return null;
                
                LabelStatement label = null;
                    
                // Look for the closest label statement
                foreach (Statement statement in method.Statements)
                {
                    if (statement is LabelStatement)
                        label = (LabelStatement)statement;

                    if (object.ReferenceEquals(statement, this))
                        return label;
                }

                return null;
            }
        }

        public GoSubStatement GoSubReference
        {
            get 
            {
                if (!IsLabelReference)
                    return null;

                MethodDeclaration method = MethodReference;

                if (method == null)
                    return null;

                LabelStatement label = LabelReference;

                if (label == null)
                    return null;

                SimpleName labelName = label.Name;
                GoSubStatement gosub = Seek(method, labelName);

                return gosub; 
            }
        }

        private GoSubStatement Seek(Tree root, SimpleName name)
        {
            if (root == null)
                return  null;

            GoSubStatement result = null;

            // Look for a gosub statement that reference the found label name
            foreach (Tree child in root.Children)
            {
                if (child is GoSubStatement && ((GoSubStatement)child).Name == name)
                    result = (GoSubStatement)child;

                if (result == null)
                    result = Seek(child, name);

                if (result != null)
                    break;
            }

            return result;
        }

		/// <summary>
		/// Constructs a new parse tree for a Return statement.
		/// </summary>
		/// <param name="expression">The expression.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public ReturnStatement(Expression expression, Span span, IList<Comment> comments, bool isLabelReference) : base(TreeType.ReturnStatement, expression, span, comments)
		{
            this.IsLabelReference = isLabelReference;
		}
	}
}

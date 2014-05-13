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
using VBConverter.CodeParser.Trees.Types;
using VBConverter.CodeParser.Trees.Declarations.Members;

namespace VBConverter.CodeParser.Trees.Statements
{
    /// <summary>
    /// A parse tree for an Exit statement.
    /// </summary>
    public sealed class ExitStatement : Statement
    {
        #region Fields

        private readonly BlockType _ExitType;
        private readonly Location _ExitArgumentLocation;

        #endregion

        #region Properties

        /// <summary>
        /// The type of tree this statement exits.
        /// </summary>
        public BlockType ExitType
        {
            get { return _ExitType; }
        }

        /// <summary>
        /// The location of the exit statement type.
        /// </summary>
        public Location ExitArgumentLocation
        {
            get { return _ExitArgumentLocation; }
        }

        /// <summary>
        /// The tree that the ExitStatement is related to.
        /// </summary>
        public Tree RelatedTree
        {
            get
            {
                if (ExitType == BlockType.None)
                    return null;

                // Look for the parent method declaration
                Tree parentTree = this;
                do
                {
                    parentTree = parentTree.Parent;

                    if (ExitType == BlockType.Do && parentTree is DoBlockStatement)
                        return parentTree;
                    else if (ExitType == BlockType.For && parentTree is ForBlockStatement)
                        return parentTree;
                    else if (ExitType == BlockType.While && parentTree is WhileBlockStatement)
                        return parentTree;
                    else if (ExitType == BlockType.Sub && parentTree is SubDeclaration)
                        return parentTree;
                    else if (ExitType == BlockType.Function && parentTree is FunctionDeclaration)
                        return parentTree;
                    else if (ExitType == BlockType.Property && (parentTree is GetAccessorDeclaration || parentTree is SetAccessorDeclaration))
                        return parentTree;
                }
                while (parentTree != null);

                return null;
            }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructs a parse tree for an Exit statement.
        /// </summary>
        /// <param name="exitType">The type of tree this statement exits.</param>
        /// <param name="exitArgumentLocation">The location of the exit statement type.</param>
        /// <param name="span">The location of the parse tree.</param>
        /// <param name="comments">The comments for the parse tree.</param>
        public ExitStatement(BlockType exitType, Location exitArgumentLocation, Span span, IList<Comment> comments) : base(TreeType.ExitStatement, span, comments)
        {

            switch (exitType)
            {
                case BlockType.Do:
                case BlockType.For:
                case BlockType.While:
                case BlockType.Select:
                case BlockType.Sub:
                case BlockType.Function:
                case BlockType.Property:
                case BlockType.Try:
                case BlockType.None:
                    break;
                // OK

                default:
                    throw new ArgumentOutOfRangeException("exitType");
            }

            _ExitType = exitType;
            _ExitArgumentLocation = exitArgumentLocation;
        }

        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
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

namespace VBConverter.CodeParser.Trees
{
	/// <summary>
	/// The root class of all trees.
	/// </summary>
	public abstract class Tree
    {
        #region Nested Elements

        public class Types
        {
            public static bool IsCollection(TreeType type)
            {
                return
                    type >= TreeType.ArgumentCollection &&
                    type <= TreeType.DeclarationCollection;
            }

            public static bool IsFile(TreeType type)
            {
                return
                    type == TreeType.FileTree;
            }

            public static bool IsComment(TreeType type)
            {
                return
                    type == TreeType.Comment;
            }

            public static bool IsName(TreeType type)
            {
                return
                    type >= TreeType.SimpleName &&
                    type <= TreeType.QualifiedName;
            }

            public static bool IsType(TreeType type)
            {
                return
                    type >= TreeType.IntrinsicType &&
                    type <= TreeType.ArrayType;
            }

            public static bool IsArgument(TreeType type)
            {
                return
                    type == TreeType.Argument;
            }

            public static bool IsExpression(TreeType type)
            {
                return
                    type >= TreeType.SimpleNameExpression &&
                    type <= TreeType.GetTypeExpression;
            }

            public static bool IsStatement(TreeType type)
            {
                return
                    type >= TreeType.EmptyStatement &&
                    type <= TreeType.EndBlockStatement;
            }

            public static bool IsModifier(TreeType type)
            {
                return
                    type == TreeType.Modifier;
            }

            public static bool IsVariableDeclarator(TreeType type)
            {
                return
                    type == TreeType.VariableDeclarator;
            }

            public static bool IsCaseClause(TreeType type)
            {
                return
                    type >= TreeType.ComparisonCaseClause &&
                    type <= TreeType.RangeCaseClause;
            }

            public static bool IsAttribute(TreeType type)
            {
                return
                    type == TreeType.AttributeTree;
            }

            public static bool IsDeclaration(TreeType type)
            {
                return
                    type >= TreeType.EmptyDeclaration &&
                    type <= TreeType.DelegateFunctionDeclaration;
            }

            public static bool IsParameter(TreeType type)
            {
                return
                    type == TreeType.Parameter;
            }

            public static bool IsTypeParameter(TreeType type)
            {
                return
                    type == TreeType.TypeParameter;
            }

            public static bool IsImport(TreeType type)
            {
                return
                    type >= TreeType.NameImport &&
                    type <= TreeType.AliasImport;
            }

            public static bool IsCollection(Tree tree)
            {
                return tree != null && IsCollection(tree.Type);
            }

            public static bool IsFile(Tree tree)
            {
                return tree != null && IsFile(tree.Type);
            }

            public static bool IsComment(Tree tree)
            {
                return tree != null && IsComment(tree.Type);
            }

            public static bool IsName(Tree tree)
            {
                return tree != null && IsName(tree.Type);
            }

            public static bool IsType(Tree tree)
            {
                return tree != null && IsType(tree.Type);
            }

            public static bool IsArgument(Tree tree)
            {
                return tree != null && IsArgument(tree.Type);
            }

            public static bool IsExpression(Tree tree)
            {
                return tree != null && IsExpression(tree.Type);
            }

            public static bool IsStatement(Tree tree)
            {
                return tree != null && IsStatement(tree.Type);
            }

            public static bool IsModifier(Tree tree)
            {
                return tree != null && IsModifier(tree.Type);
            }

            public static bool IsVariableDeclarator(Tree tree)
            {
                return tree != null && IsVariableDeclarator(tree.Type);
            }

            public static bool IsCaseClause(Tree tree)
            {
                return tree != null && IsCaseClause(tree.Type);
            }

            public static bool IsAttribute(Tree tree)
            {
                return tree != null && IsAttribute(tree.Type);
            }

            public static bool IsDeclaration(Tree tree)
            {
                return tree != null && IsDeclaration(tree.Type);
            }

            public static bool IsParameter(Tree tree)
            {
                return tree != null && IsParameter(tree.Type);
            }

            public static bool IsTypeParameter(Tree tree)
            {
                return tree != null && IsTypeParameter(tree.Type);
            }

            public static bool IsImport(Tree tree)
            {
                return tree != null && IsImport(tree.Type);
            }
        }

        #endregion

        #region Fields

        private readonly TreeType _Type;
		private readonly Span _Span;
		private Tree _Parent;
		private ReadOnlyCollection<Tree> _Children;

        #endregion

        #region Properties

        /// <summary>
		/// The type of the tree.
		/// </summary>
		public TreeType Type 
        {
			get { return _Type; }
		}

		/// <summary>
		/// The location of the tree.
		/// </summary>
		/// <remarks>
		/// The span ends at the first character beyond the tree
		/// </remarks>
		public Span Span 
        {
			get { return _Span; }
		}

		/// <summary>
		/// The parent of the tree. Nothing if the root tree.
		/// </summary>
		public Tree Parent 
        {
			get { return _Parent; }
		}

		/// <summary>
		/// The children of the tree.
		/// </summary>
		public ReadOnlyCollection<Tree> Children 
        {
			get 
            {
				if (_Children == null)
				{
					List<Tree> ChildList = new List<Tree>();

					GetChildTrees(ChildList);
					_Children = new ReadOnlyCollection<Tree>(ChildList);
				}

				return _Children;
			}
		}

		/// <summary>
		/// Whether the tree is 'bad'.
		/// </summary>
		public virtual bool IsBad 
        {
			get { return false; }
        }

        #endregion

        #region Constructor

        protected Tree(TreeType type, Span span)
		{
			Debug.Assert(type >= TreeType.SyntaxError && type <= TreeType.FileTree);
			_Type = type;
			_Span = span;
        }

        #endregion

        #region Protected Methods

        protected void SetParent(Tree child)
		{
			if (child != null)
			{
				child._Parent = this;
			}
		}

		protected void SetParents<T>(IList<T> children) where T : Tree
		{
			if (children != null)
			{
				foreach (Tree Child in children) {
					SetParent(Child);
				}
			}
        }

        #endregion

        #region Static Methods

        protected static void AddChild(IList<Tree> childList, Tree child)
		{
			if (child != null)
				childList.Add(child);
		}

		protected static void AddChildren<T>(IList<Tree> childList, ReadOnlyCollection<T> children) where T : Tree
		{
			if (children != null)
            {
				foreach (Tree Child in children)
					childList.Add(Child);
			}
		}

		protected virtual void GetChildTrees(IList<Tree> childList)
		{
			// By default, trees have no children
		}

        #endregion
    }
}

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
using VBConverter.CodeParser.Trees.Declarations;

namespace VBConverter.CodeParser.Trees
{
	/// <summary>
	/// A parse tree for an entire file.
	/// </summary>
	public sealed class FileTree : Tree
    {
        #region Fields

        private string _name;
        private FileType _fileType;
        private DeclarationCollection _declarations;

        #endregion

        #region Properties

        /// <summary>
        /// The name of the file.
        /// </summary>
        public string Name
        {
            get { return _name; }
            private set { _name = value; }
        }

        /// <summary>
        /// The file type. Can be a class, a module or a form.
        /// </summary>
        public FileType FileType
        {
            get { return _fileType; }
            set { _fileType = value; }
        }

        /// <summary>
		/// The declarations in the file.
		/// </summary>
		public DeclarationCollection Declarations 
        {
			get { return _declarations; }
            private set { _declarations = value; }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructs a new file parse tree.
        /// </summary>
        /// <param name="type">The type of the file.</param>
        /// <param name="name">The name of the file.</param>
        /// <param name="declarations">The declarations in the file.</param>
        /// <param name="span">The location of the tree.</param>
        public FileTree(FileType type, string name, DeclarationCollection declarations, Span span) : base(TreeType.FileTree, span)
        {
            SetParent(declarations);
            
            this.FileType = FileType.None;
            this.Name = name;
            this.Declarations = declarations;
        }

        /// <summary>
		/// Constructs a new file parse tree.
		/// </summary>
		/// <param name="declarations">The declarations in the file.</param>
		/// <param name="span">The location of the tree.</param>
		public FileTree(DeclarationCollection declarations, Span span) : this(FileType.None, string.Empty, declarations, span)
		{
        }

        #endregion

        #region Methods

        protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Declarations);
        }

        #endregion
    }
}

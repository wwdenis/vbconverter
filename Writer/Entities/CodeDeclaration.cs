using System;
using System.Collections.Generic;
using System.Text;

using System.Collections;
using System.CodeDom;

namespace VBConverter.CodeWriter.Entities
{
    internal class CodeDeclaration
    {
        #region Fields

        private CodeDeclarationType _type;
        private CodeTypeMemberCollection _members;
        private CodeStatementCollection _statements;
        private CodeCommentStatementCollection _comments;

        #endregion

        #region Constructors

        public CodeDeclaration(CodeDeclarationType type)
        {
            this.Type = type;
            this.Members = new CodeTypeMemberCollection();
            this.Statements = new CodeStatementCollection();
            this.Comments = new CodeCommentStatementCollection();
        }

        public CodeDeclaration(CodeTypeMemberCollection members) : this(CodeDeclarationType.MemberCollection)
        {
            this.Members.AddRange(members);
        }

        public CodeDeclaration(CodeStatementCollection statements) : this(CodeDeclarationType.StatementCollection)
        {
            this.Statements.AddRange(statements);
        }

        public CodeDeclaration(CodeCommentStatementCollection comments) : this(CodeDeclarationType.CommentCollection)
        {
            this.Comments.AddRange(comments);
        }

        #endregion

        #region Properties

        public CodeDeclarationType Type
        {
            get { return _type; }
            private set { _type = value; }
        }

        public CodeTypeMemberCollection Members 
        {
            get { return _members; }
            private set { _members = value; }
        }

        public CodeStatementCollection Statements
        {
            get { return _statements; }
            private set { _statements = value; }
        }

        public CodeCommentStatementCollection Comments
        {
            get { return _comments; }
            private set { _comments = value; }
        }

        #endregion
    }

    internal enum CodeDeclarationType
    {
        MemberCollection,
        StatementCollection,
        CommentCollection
    }
}

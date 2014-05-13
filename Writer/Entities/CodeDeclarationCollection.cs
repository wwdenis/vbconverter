using System;
using System.Collections.Generic;
using System.Text;
using System.CodeDom;

namespace VBConverter.CodeWriter.Entities
{
    internal class CodeDeclarationCollection : List<CodeDeclaration>
    {
        #region Properties

        public CodeTypeMemberCollection Members
        {
            get 
            {
                CodeTypeMemberCollection members = new CodeTypeMemberCollection();

                foreach (CodeDeclaration item in this)
                    members.AddRange(item.Members);

                return members;
            }
        }

        public CodeStatementCollection Statements
        {
            get 
            {
                CodeStatementCollection statements = new CodeStatementCollection();

                foreach (CodeDeclaration item in this)
                    statements.AddRange(item.Statements);

                return statements;
            }
        }

        public CodeCommentStatementCollection Comments
        {
            get
            {
                CodeCommentStatementCollection comments = new CodeCommentStatementCollection();

                foreach (CodeDeclaration item in this)
                    comments.AddRange(item.Comments);

                return comments;
            }
        }

        #endregion
    }
}

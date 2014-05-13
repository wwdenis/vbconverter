using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.CodeDom;

namespace VBConverter.CodeWriter.Entities
{
    internal class CodeVariableDeclaratorCollection
    {
        private CodeCollectionType _collectionType;
        private CodeTypeMemberCollection _members;
        private CodeStatementCollection _statements;
        
        public CodeCollectionType CollectionType
        {
            get { return _collectionType; }
            private set { _collectionType = value; }
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

        public CodeVariableDeclaratorCollection(CodeCollectionType collectionType)
        {
            this.CollectionType = collectionType;

            this.Members = new CodeTypeMemberCollection();
            this.Statements = new CodeStatementCollection();
        }

        public CodeVariableDeclaratorCollection(CodeTypeMemberCollection members) : this(CodeCollectionType.TypeMember)
        {
            this.Members.AddRange(members);
        }

        public CodeVariableDeclaratorCollection(CodeStatementCollection statements) : this(CodeCollectionType.Statement)
        {
            this.Statements.AddRange(statements);
        }
    }

    internal enum CodeCollectionType
    {
        TypeMember,
        Statement
    }
}

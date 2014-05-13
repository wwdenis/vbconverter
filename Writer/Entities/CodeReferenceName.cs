using System;
using System.Collections.Generic;
using System.Text;

using System.CodeDom;

namespace VBConverter.CodeWriter.Entities
{
    internal class CodeReferenceName
    {
        #region Fields

        private string _name;
        private bool _isVariable;
        private int _rank;
        private CodeExpressionCollection _arrayLengths;

        #endregion

        #region Constructors

        public CodeReferenceName()
        {
        }

        public CodeReferenceName(string name)
        {
            this.Name = name;
        }

        public CodeReferenceName(string name, int rank, CodeExpressionCollection arrayType) : this(name)
        {
            this.IsVariable = true;
            this.Rank = rank;
            this.ArrayLengths = arrayType;
        }

        #endregion

        #region Properties

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        public bool IsVariable
        {
            get { return _isVariable; }
            set { _isVariable = value; }
        }

        public int Rank
        {
            get { return _rank; }
            set { _rank = value; }
        }

        public CodeExpressionCollection ArrayLengths
        {
            get { return _arrayLengths; }
            set { _arrayLengths = value; }
        }

        #endregion
    }
}

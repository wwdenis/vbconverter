using System;
using System.Collections.Generic;
using System.Text;

using System.CodeDom;

namespace VBConverter.CodeWriter.Entities
{
    internal class CodeParameter
    {
        private bool _isOptional;
        private CodeParameterDeclarationExpression _parameter;
        private CodeExpression _initValue;

        public bool IsOptional
        {
            get { return _isOptional; }
            set { _isOptional = value; }
        }

        public CodeParameterDeclarationExpression Parameter
        {
            get { return _parameter; }
            set { _parameter = value; }
        }


        public CodeExpression InitValue
        {
            get { return _initValue; }
            set { _initValue = value; }
        }
    }
}

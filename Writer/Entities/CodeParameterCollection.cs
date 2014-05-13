using System;
using System.Collections.Generic;
using System.Text;

using System.CodeDom;

namespace VBConverter.CodeWriter.Entities
{
    internal class CodeParameterCollection : List<CodeParameter>
    {
        public CodeParameterCollection() : base()
        {
        }

        public CodeParameterDeclarationExpressionCollection Parameters
        {
            get
            {
                CodeParameterDeclarationExpressionCollection list = new CodeParameterDeclarationExpressionCollection();
                foreach (CodeParameter item in this)
                    list.Add(item.Parameter);
                
                return list;
            }
        }

        public CodeExpressionCollection Initializers
        {
            get
            {
                CodeExpressionCollection list = new CodeExpressionCollection();
                foreach (CodeParameter item in this)
                    list.Add(item.InitValue);

                return list;
            }
        }
    }
}

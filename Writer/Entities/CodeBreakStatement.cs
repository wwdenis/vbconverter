using System;
using System.Collections.Generic;
using System.Text;

using System.CodeDom;

namespace VBConverter.CodeWriter.Entities
{
    internal class CodeBreakStatement : CodeSnippetStatement
    {
        public CodeBreakStatement() : base("break;")
        {
        }
    }
}

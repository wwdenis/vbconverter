using System;
using System.Collections.Generic;
using System.Text;

namespace VBConverter.CodeParser.Trees
{
    /// <summary>
    /// The type of a file tree (VB6)
    /// </summary>
    public enum FileType
    {
        None,
        Class,
        Module,
        Form
    }
}

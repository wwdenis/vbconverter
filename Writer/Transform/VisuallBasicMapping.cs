using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Xml;
using System.IO;

using System.CodeDom;
using Microsoft.VisualBasic;

using VBConverter.CodeWriter.Resources;

namespace VBConverter.CodeWriter.Transform
{
    internal struct KeywordMapping
    {
        public readonly string Library;
        public readonly string Class;
        public readonly string VisualBasicKeyword;
        public readonly string DotNetKeyword;

        public KeywordMapping(string library, string @class, string visualBasicKeyword, string dotNetKeyword)
        {
            this.Library = library;
            this.Class = @class;
            this.VisualBasicKeyword = visualBasicKeyword;
            this.DotNetKeyword = dotNetKeyword;
        }
    }

    internal abstract class VisuallBasicUtil
    {
        public static CodeMethodReferenceExpression MapIntrinsicName(CodeMethodReferenceExpression reference)
        {
            if (reference == null || string.IsNullOrEmpty(reference.MethodName))
                return null;

            string library = string.Empty;
            string @class = string.Empty;
            string name = reference.MethodName;

            if (reference.TargetObject is CodeVariableReferenceExpression)
            {
                CodeVariableReferenceExpression targetVariable = (CodeVariableReferenceExpression)reference.TargetObject;

                @class = targetVariable.VariableName;
                library = string.Empty;
            }
            else if (reference.TargetObject is CodeMethodReferenceExpression)
            {
                CodeMethodReferenceExpression targetName = (CodeMethodReferenceExpression)reference.TargetObject;
                CodeVariableReferenceExpression targetLibrary = targetName.TargetObject as CodeVariableReferenceExpression;

                @class = targetName.MethodName;
                if (targetLibrary != null)
                    library = targetLibrary.VariableName;
            }

            return MapIntrinsicName(library, @class, name);
        }

        public static CodeMethodReferenceExpression MapIntrinsicName(CodeVariableReferenceExpression reference)
        {
            if (reference == null || string.IsNullOrEmpty(reference.VariableName))
                return null;

            string library = string.Empty;
            string @class = string.Empty;
            string name = reference.VariableName;

            return MapIntrinsicName(library, @class, name);
        }

        private static CodeMethodReferenceExpression MapIntrinsicName(string library, string @class, string name)
        {
            TextReader reader = new StringReader(Main.VisualBasicMappings);
            DataSet ds = new DataSet();
            ds.ReadXml(reader);

            string newClass = string.Empty;
            string newName = string.Empty;
            
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                if (string.Compare((string)row["supported"], "yes", true) == 0 && 
                    string.Compare((string)row["name"], name, true) == 0 && 
                    (string.Compare((string)row["class"], @class, true) == 0 || string.IsNullOrEmpty(@class)) && 
                    (string.Compare((string)row["library"], library, true) == 0 || string.IsNullOrEmpty(library))
                    )
                {
                    newClass = (string)row["new_class"];
                    newName = (string)row["new_name"];

                    if (string.IsNullOrEmpty(newClass))
                        newClass = (string)row["class"];

                    if (string.IsNullOrEmpty(newName))
                        newName = (string)row["name"];

                    break;
                }
            }
            
            CodeMethodReferenceExpression result = null;
            if (!string.IsNullOrEmpty(newName))
            {
                CodeVariableReferenceExpression variable = new CodeVariableReferenceExpression(newClass);
                result = new CodeMethodReferenceExpression(variable, newName);
            }
            
            return result;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

using System.CodeDom;
using System.CodeDom.Compiler;
using Microsoft.CSharp;

using VBConverter.CodeParser;
using VBConverter.CodeParser.Base;
using VBConverter.CodeParser.Error;
using VBConverter.CodeParser.Serializers;
using VBConverter.CodeParser.Tokens;
using VBConverter.CodeParser.Trees;
using VBConverter.CodeParser.Trees.Arguments;
using VBConverter.CodeParser.Trees.Attributes;
using VBConverter.CodeParser.Trees.CaseClause;
using VBConverter.CodeParser.Trees.Collections;
using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Declarations;
using VBConverter.CodeParser.Trees.Declarations.Members;
using VBConverter.CodeParser.Trees.Expressions;
using VBConverter.CodeParser.Trees.Imports;
using VBConverter.CodeParser.Trees.Initializers;
using VBConverter.CodeParser.Trees.Modifiers;
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.Parameters;
using VBConverter.CodeParser.Trees.Statements;
using VBConverter.CodeParser.Trees.TypeNames;
using VBConverter.CodeParser.Trees.TypeParameters;
using VBConverter.CodeParser.Trees.Types;
using VBConverter.CodeParser.Trees.VariableDeclarators;
using Microsoft.VisualBasic;

namespace VBConverter.CodeWriter
{
    public class ConverterEngine
    {

        #region Fields

        private LanguageVersion _version;
        private string _source;
        private string _result;
        private List<SyntaxError> _errors;
        private FileType _fileType;
        private string _fileName;
        private DestinationLanguage _resultType;

        #endregion

        #region Properties

        public LanguageVersion Version
        {
            get
            {
                return _version;
            }
            private set
            {
                if (!Enum.IsDefined(typeof(LanguageVersion), value))
                    throw new ArgumentNullException("Version");

                _version = value;
            }
        }

        public string Source
        {
            get 
            { 
                return _source; 
            }
            private set 
            {
                //if (string.IsNullOrEmpty(value))
                //    throw new ArgumentNullException("Source");

                if (string.Compare(_source, value, true) != 0)
                {
                    Result = string.Empty;
                    Errors.Clear();
                    _source = FilterSource(value);
                }
            }
        }

        public string Result
        {
            get { return _result; }
            private set { _result = value; }
        }

        public List<SyntaxError> Errors
        {
            get 
            {
                if (_errors == null)
                    _errors = new List<SyntaxError>();

                return _errors; 
            }
            private set 
            { 
                _errors = value; 
            }
        }

        public FileType FileType
        {
            get { return _fileType; }
            private set { _fileType = value; }
        }

        public string FileName
        {
            get { return _fileName; }
            set { _fileName = value; }
        }

        public DestinationLanguage ResultType
        {
            get { return _resultType; }
            set { _resultType = value; }
        }

        public string ResultFileExtension
        {
            get { return GetCodeProvider().FileExtension; }
        }

        #endregion

        #region Constructor

        public ConverterEngine(LanguageVersion version, string source)
        {
            this.Version = version;
            this.Source = source;
            this.ResultType = DestinationLanguage.CSharp;
        }

        #endregion

        #region Static Methods

        static public string FilterSource(string text)
        {
            string[] lines = text.Split('\n');
            string previous = string.Empty;
            
            int first = 0;
            int current = 0;

            foreach (string line in lines)
            {
                current++;
                bool isPreviousAttribute = previous.Trim().StartsWith("Attribute ", StringComparison.InvariantCultureIgnoreCase);
                bool isLineAttribute = line.Trim().StartsWith("Attribute ", StringComparison.InvariantCultureIgnoreCase);

                if (isPreviousAttribute && !isLineAttribute)
                {
                    first = current;
                    break;
                }

                previous = line;
            }

            StringBuilder builder = new StringBuilder();
            string result = string.Empty;

            if (first > 0)
            {
                current = 0;
                foreach (string line in lines)
                {
                    current++;
                    if (current >= first)
                    {
                        if (!line.Trim().StartsWith("Attribute ", StringComparison.InvariantCultureIgnoreCase))
                            builder.Append(line + '\n');
                    }
                }
                result = builder.ToString();
            }
            else
            {
                result = text;
            }

            return result;
        }

        #endregion

        #region Private Methods

        private string OutputCode(Tree tree)
        {
            CodeGenerator generator = new CodeGenerator(Version);

            CodeTypeDeclaration fileTree = generator.Convert(tree);
            CodeNamespace global = generator.CreateNamespace("Global", fileTree);
            CodeCompileUnit unit = generator.CreateCompileUnit(global);

            // Create a StringWriter (derived from TextWriter) to output the code
            StringBuilder builder = new StringBuilder();
            StringWriter writer = new StringWriter(builder);

            // Create an instance of the CodeDomProvider and pass the compile unit
            CodeDomProvider provider = GetCodeProvider();
            
            // Options for code generation
            CodeGeneratorOptions options = new CodeGeneratorOptions();

            // Generate the code using the code provider
            provider.GenerateCodeFromCompileUnit(unit, writer, options);
            writer.Close();

            string output = builder.ToString();

            return output;
        }

        private CodeDomProvider GetCodeProvider()
        {
            // Create an instance of the CodeDomProvider and pass the compile unit
            CodeDomProvider provider = null;
            switch (ResultType)
            {
                case DestinationLanguage.CSharp:
                    provider = new CSharpCodeProvider();
                    break;
                case DestinationLanguage.VisualBasic:
                    provider = new VBCodeProvider();
                    break;
                default:
                    break;
            }
            return provider;
        }

        #endregion

        #region Public Methods

        public bool Convert()
        {
            List<SyntaxError> errors = null;
            Tree tree = null;
            
            // Parsing order: File, Declaration, Statement, Expression 
            for (int step = 0; step < 4; step++)
            {
                Scanner scanner = new Scanner(Source, Version);
                Parser parser = new Parser();
                List<SyntaxError> parsingErrors = new List<SyntaxError>();

                switch (step)
                {
                    case 0:
                        tree = parser.ParseFile(scanner, parsingErrors);
                        break;
                    case 1:
                        tree = parser.ParseDeclaration(scanner, parsingErrors);
                        break;
                    case 2:
                        tree = parser.ParseStatement(scanner, parsingErrors);
                        break;
                    case 3:
                        tree = parser.ParseExpression(scanner, parsingErrors);
                        break;
                }

                if (scanner.IsEndOfFile)
                {
                    errors = parsingErrors;
                    if (errors.Count == 0)
                        break;
                }
            }

            if (tree is FileTree)
                tree = new FileTree(FileType.Module, FileName,((FileTree)tree).Declarations, tree.Span);

            string result = string.Empty;
            bool success = (errors.Count == 0);
            
            if (success)
                result = OutputCode(tree);
            
            this.Result = result;
            this.Errors = errors;

            return success;
        }

        public string GetErrors()
        {
            string text = string.Empty;
            StringBuilder builder = new StringBuilder();
            foreach (SyntaxError item in this.Errors)
                builder.Append("\n" + item);

            if (builder.Length > 0)
                text = "### ERRORS ###" + builder.ToString();
            
            return text;
        }

        #endregion

    }

    public enum DestinationLanguage
    {
        CSharp,
        VisualBasic
    }
}


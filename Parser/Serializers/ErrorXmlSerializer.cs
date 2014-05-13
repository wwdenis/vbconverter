using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
//
// Visual Basic .NET Parser
//
// Copyright (C) 2005, Microsoft Corporation. All rights reserved.
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
// EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
// MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
//
using VBConverter.CodeParser.Error;

namespace VBConverter.CodeParser.Serializers
{
	public class ErrorXmlSerializer : XmlSerializer<SyntaxError>
    {
        #region Constructor

        public ErrorXmlSerializer(XmlWriter Writer) : base(Writer)
		{
		}

		public ErrorXmlSerializer(XmlWriter Writer, bool showPosition) : base(Writer, showPosition)
		{
        }

        #endregion

        #region Public Methods

        public override void Serialize(SyntaxError SyntaxError)
		{
            Output.WriteStartElement(SyntaxError.Type.ToString());
			SerializePosition(SyntaxError.Span);
            Output.WriteString(SyntaxError.ToString());
			Output.WriteEndElement();
        }

        #endregion
    }
}

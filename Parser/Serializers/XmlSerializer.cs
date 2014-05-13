using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace VBConverter.CodeParser.Serializers
{
	/// <summary>
	/// Provides a base Serializer implementation
	/// </summary>
	public abstract class XmlSerializer<CustomType>
    {
        #region Fields

        private XmlWriter _output;
		private bool _outputPosition = true;

        #endregion

        #region Constructor

        public XmlSerializer(XmlWriter Writer)
		{
			this.Output = Writer;
		}

		public XmlSerializer(XmlWriter Writer, bool showPosition) : this(Writer)
		{
			this.OutputPosition = showPosition;
        }

        #endregion

        #region Properties

        public bool OutputPosition 
        {
			get { return _outputPosition; }
			set { _outputPosition = value; }
        }

        public XmlWriter Output
        {
            get { return _output; }
            set { _output = value; }
        }

        #endregion

        #region Implemented Methods

        public void SerializePosition(Span span)
        {
            if (!OutputPosition)
                return;

            Location start = span.Start;
            Location finish = span.Finish;
            string[] prefixes = null;
            long[] values = null;

            if (start.IsValid && finish.IsValid)
            {
                prefixes = new string[] { "startLine", "startCol", "endLine", "endCol" };
                values = new long[] { start.Line, start.Column, finish.Line, finish.Column };
            }
            else if (start.IsValid)
            {
                prefixes = new string[] { "line", "col" };
                values = new long[] { start.Line, start.Column };
            }
            else if (finish.IsValid)
            {
                prefixes = new string[] { "line", "col" };
                values = new long[] { finish.Line, finish.Column };
            }

            for (int i = 0; i < prefixes.Length; i++)
                Output.WriteAttributeString(prefixes[i], values[i].ToString());
        }

        public void SerializePosition(Location location)
        {
            SerializePosition(new Span(location, Location.Empty));
        }

		public void Serialize(List<CustomType> items)
		{
			foreach (CustomType item in items)
				Serialize(item);
        }

        #endregion

        #region Abstract Methods

        public abstract void Serialize(CustomType Item);

        #endregion
    }
}

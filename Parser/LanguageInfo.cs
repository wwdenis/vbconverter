using System;
using System.Collections.Generic;
using System.Text;
using System.Globalization;

namespace VBConverter.CodeParser
{
    /// <summary>
    /// Provides detailed language information.
    /// </summary>
    public class LanguageInfo
    {
        private LanguageVersion _Value;
        private string _Name;

        public string Name
        {
            get { return this._Name; }
        }

        public LanguageVersion Value
        {
            get { return this._Value; }
            set
            {
                this._Value = value;
                this._Name = Translate(value);
            }
        }

        public LanguageInfo(LanguageVersion _Value)
        {
            this.Value = _Value;
        }

        public override string ToString()
        {
            return this.Name;
        }

        public static LanguageInfo[] GetList()
        {
            Array EnumList = Enum.GetValues(typeof(LanguageVersion));
            List<LanguageInfo> InfoList = new List<LanguageInfo>();

            foreach (int Value in EnumList)
            {
                LanguageInfo Item = new LanguageInfo((LanguageVersion)Value);
                InfoList.Add(Item);
            }

            return InfoList.ToArray();
        }

        public static string Translate(LanguageVersion Value)
        {
            string Name = "Undefined";
            switch (Value.ToString())
            {
                case "VB6":
                    Name = "Visual Basic 6.0";
                    break;
                case "VB7":
                    Name = "Visual Basic 2003";
                    break;
                case "VB8":
                    Name = "Visual Basic 2005";
                    break;
            }
            return Name;
        }

        public static LanguageVersion ParseVersion(string Value)
        {
            double Parsed = 0;
            LanguageVersion Version = LanguageVersion.Non;
            if (double.TryParse(Value, NumberStyles.AllowDecimalPoint, new CultureInfo("en-US"), out Parsed))
            {
                switch (Parsed.ToString("0.0"))
                {
                    case "6.0":
                        Version = LanguageVersion.VB6;
                        break;
                    case "7.1":
                        Version = LanguageVersion.VB7;
                        break;
                    case "8.0":
                        Version = LanguageVersion.VB8;
                        break;
                }
            }

            return Version;
        }

        public static bool ImplementsVB60(LanguageVersion Versions)
        {
            return Implements(Versions, LanguageVersion.VB6);
        }

        public static bool ImplementsVB71(LanguageVersion Versions)
        {
            return Implements(Versions, LanguageVersion.VB7);
        }

        public static bool ImplementsVB80(LanguageVersion Versions)
        {
            return Implements(Versions, LanguageVersion.VB8);
        }

        public static bool Implements(LanguageVersion versions, LanguageVersion value)
        {
            return ((versions & value) == value);
        }
    }
}
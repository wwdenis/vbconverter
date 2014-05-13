using System;
using System.Collections.Generic;
using System.Text;

namespace VBConverter.UI.Comparers
{
    internal class CaseInsensitiveStringComparer : EqualityComparer<string>
    {
        public override bool Equals(string first, string second)
        {
            return string.Compare(first, second, true) == 0;
        }

        public override int GetHashCode(string obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return obj.ToLower().GetHashCode();
        }
    }
}

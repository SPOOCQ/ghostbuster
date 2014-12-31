#region License

/*
Copyright (c) 2009, G.W. van der Vegt
All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided
that the following conditions are met:

* Redistributions of source code must retain the above copyright notice, this list of conditions and the
  following disclaimer.

* Redistributions in binary form must reproduce the above copyright notice, this list of conditions and
  the following disclaimer in the documentation and/or other materials provided with the distribution.

* Neither the name of G.W. van der Vegt nor the names of its contributors may be
  used to endorse or promote products derived from this software without specific prior written
  permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL
THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF
THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/
#endregion License

namespace Ghostbuster
{
    //See: http://www.codeproject.com/KB/recipes/wildcardtoregex.aspx
    //usage: Wildcard wildcard = new Wildcard("*.txt", RegexOptions.IgnoreCase); 
    //       if(wildcard.IsMatch(file)) 
    //       
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Represents a wildcard running on the
    /// <see cref="System.Text.RegularExpressions"/> engine.
    /// </summary>
    public class Wildcard : Regex
    {
        /// <summary>
        /// Initializes a wildcard with the given search pattern.
        /// </summary>
        ///
        /// <param name="pattern"> The wildcard pattern to match. </param>
        public Wildcard(string pattern)
            : base(WildcardToRegex(pattern))
        {
            Pattern = pattern;
        }

        /// <summary>
        /// Initializes a wildcard with the given search pattern and options.
        /// </summary>
        ///
        /// <param name="pattern"> The wildcard pattern to match. </param>
        /// <param name="options">  A combination of one or more
        ///                         <see cref="System.Text.RegexOptions"/>. </param>
        public Wildcard(string pattern, RegexOptions options) :
            base(WildcardToRegex(pattern), options)
        {
            Pattern = pattern;
        }

        /// <summary>
        /// Gets or sets the pattern.
        /// </summary>
        ///
        /// <value>
        /// The pattern.
        /// </value>
        public string Pattern
        {
            get;
            set;
        }

        /// <summary>
        /// Converts a wildcard to a regex.
        /// </summary>
        ///
        /// <param name="pattern"> The wildcard pattern to convert. </param>
        ///
        /// <returns>
        /// A regex equivalent of the given wildcard.
        /// </returns>
        private static string WildcardToRegex(string pattern)
        {
            return "^" + Regex.Escape(pattern).Replace("\\*", ".*").Replace("\\?", ".") + "$";
        }
    }
}

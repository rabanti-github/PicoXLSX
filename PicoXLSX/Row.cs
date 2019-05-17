using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PicoXLSX
{
    /// <summary>
    /// Struct representing a row of a worksheet
    /// </summary>
    public struct Row
    {
        #region constants
        /// <summary>
        /// Default row height as constant
        /// </summary>
        public const float DEFAULT_ROW_HEIGHT = 15f;
        /// <summary>
        /// Minimum row height as constant
        /// </summary>
        public const float MIN_ROW_HEIGHT = 0f;
        /// <summary>
        /// Maximum row height as constant
        /// </summary>
        public const float MAX_ROW_HEIGHT = 409.5f;
        /// <summary>
        /// Minimum row number (zero-based) as constant
        /// </summary>
        public const int MIN_ROW_NUMBER = 0;
        /// <summary>
        /// Maximum row number (zero-based) as constant
        /// </summary>
        public const int MAX_ROW_NUMBER = 1048575;
        #endregion
    }
}

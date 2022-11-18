using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Collections;

namespace AdressesUtility
{
    /// <summary>Tooltip utility functions</summary>
    public static class ToolTipUtil
    {
        /// <summary>Set the time the tool tip shall be shown, how quick, etc.</summary>
        public static void SetDelays(ref ToolTip io_tool_tip)
        {
            // Set up the delays for the ToolTip.
            io_tool_tip.AutoPopDelay = 50000; // Default is 5000
            io_tool_tip.InitialDelay = 500;  // Default is 1000
            io_tool_tip.ReshowDelay = 100; // Default is 500
            // Force the ToolTip text to be displayed whether or not the form is active.
            io_tool_tip.ShowAlways = true;

        }
    }

}

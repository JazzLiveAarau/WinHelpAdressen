using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AdressesUtility
{
    /// <summary>Application version utility functions</summary>
    public static class VersionUtil
    {
        /// <summary>Return the version part of an installer filename</summary>
        public static string GetVersionStringFromInstallerFileName(string i_file_name)
        {
            string ret_string = "";

            bool b_setup_file = i_file_name.Contains("-version-");
            if (!b_setup_file)
                return ret_string;

            bool b_setup_file_beta = false;
            if (b_setup_file && i_file_name.Contains("-Beta"))
            {
                b_setup_file_beta = true;
            }

            if (b_setup_file && b_setup_file_beta)
                return ret_string;

            int index_version_start = i_file_name.IndexOf("-version-") + 9;
            int index_version_end = i_file_name.IndexOf(".exe");
            int length_version = index_version_end - index_version_start;

            ret_string = i_file_name.Substring(index_version_start, length_version);


            return ret_string;
        }


        // major, minor, build, and revision numbers 
        /// <summary>Return versions as numbers</summary>
        /// <param name="i_version_str">String defining the version</param>
        /// <param name="i_separation">Separation character between the numbers</param>
        /// <param name="o_major">Major number</param>
        /// <param name="o_minor">Minor number</param>
        /// <param name="o_build">Build number</param>
        public static bool VersionNumbers(string i_version_str, string i_separation, ref int o_major, ref int o_minor, ref int o_build)
        {
            o_major = -12345;
            o_minor = -12345;
            o_build = -12345;

            if (i_separation.Length != 1)
                return false;

            string number_string = "";
            for (int string_index = 0; string_index < i_version_str.Length; string_index++)
            {
                string current_char = i_version_str.Substring(string_index, 1);

                if (current_char != i_separation)
                {
                    number_string = number_string + current_char;
                }
                else if (current_char == i_separation && o_major < 0)
                {
                    if (!Int32.TryParse(number_string, out o_major))
                        return false;
                    number_string = "";
                }
                else if (current_char == i_separation && o_minor < 0)
                {
                    if (!Int32.TryParse(number_string, out o_minor))
                        return false;
                    number_string = "";
                }


                if (string_index == i_version_str.Length - 1 && o_build < 0)
                {
                    if (!Int32.TryParse(number_string, out o_build))
                        return false;
                    number_string = "";
                }
            }

            return true;
        } // VersionNumbers

    }  // VersionUtil
}

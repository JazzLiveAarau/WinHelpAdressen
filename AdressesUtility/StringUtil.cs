using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using System.Text.RegularExpressions;


namespace AdressesUtility
{
    /// <summary>String utility functions</summary>
    public static class StringUtil
    {
        /// <summary>Remove invalid characters for CSV files</summary>
        public static string RemoveInvalidCharsForCsv(string i_input_string, out bool o_changed)
        {
		   o_changed = false;
		   
		   string output_string = "";

           string[] not_allowed_chars = { ",", "\n", "^", "%", "#", "\"", "\t" };

           for (int input_index = 0; input_index < i_input_string.Length; input_index++)
           {
               string current_char = i_input_string.Substring(input_index, 1);

               string output_char = current_char;

               for (int unvalid_index = 0; unvalid_index < not_allowed_chars.Length; unvalid_index++)
               {
                  
                   if (current_char.CompareTo(not_allowed_chars[unvalid_index]) == 0)
                   {
                       o_changed = true;
                       output_char = "";
                   }
               }

               output_string = output_string + output_char;

           }
		   
		   return output_string;
        } // RemoveInvalidCharsForCsv

        /// <summary>Remove all characters except numbers</summary>
        public static string RemoveAllCharsButNumbers(string i_input_string, out bool o_changed)
        {
            o_changed = false;

            string output_string = "";

            string[] allowed_numbers = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };

            for (int input_index = 0; input_index < i_input_string.Length; input_index++)
            {
                string current_char = i_input_string.Substring(input_index, 1);

                bool is_a_number = false;

                for (int valid_index = 0; valid_index < allowed_numbers.Length; valid_index++)
                {

                    if (current_char.CompareTo(allowed_numbers[valid_index]) == 0)
                    {
                        is_a_number = true;
                    }
                }

                if (is_a_number)
                {
                    output_string = output_string + current_char;
                }
                else
                {
                    o_changed = true;
                }

            }

            return output_string;
        } // RemoveAllCharsButNumbers


    } // StringUtil
} //  AdressesUtility

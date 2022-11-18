using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;

namespace AdressesUtility
{
    static public class FileUtil
    {
        /// <summary>Get full name to the config (XML) file for this application.</summary>
        public static string ConfigFileName(string i_name, string i_exe_directory)
        {
            string config_file_name = i_name.Trim() + ".config";

            string full_name = Path.Combine(i_exe_directory, config_file_name);

            return full_name;
        }

        /// <summary>Get full name to the directory with the list of adresses. Create the directory if not existing.</summary>
        public static string EmailAdressesDirectory(string i_addresses_directory, string i_exe_directory)
        {
            string email_addresses_directory = Path.Combine(i_exe_directory, i_addresses_directory);

            if (!Directory.Exists(email_addresses_directory))
            {
                Directory.CreateDirectory(email_addresses_directory);
            }

            return email_addresses_directory;
        }

        /// <summary>Get full name to the file with addresses.</summary>
        public static string AddressesFileName(string i_addresses_file_name, string i_addresses_directory, string i_exe_directory)
        {
            string path_file_name_addresses = Path.Combine(EmailAdressesDirectory(i_addresses_directory, i_exe_directory), i_addresses_file_name);

            return path_file_name_addresses;
        }

        /// <summary>Get local name for the backup file with addresses (year, month, day, hour, second and machine are added to the file name).</summary>
        public static string BackupAddressesFileName(string i_addresses_file_name, string i_addresses_directory, string i_exe_directory)
        {
            string time_addresses_file_name = Path.GetFileNameWithoutExtension(i_addresses_file_name);

            time_addresses_file_name = time_addresses_file_name + "_" + TimeUtil.YearMonthDayHourMinSec()
                + "_" + System.Environment.MachineName + Path.GetExtension(i_addresses_file_name);

            string path_time_file_name_addresses = Path.Combine(EmailAdressesDirectory(i_addresses_directory, i_exe_directory), time_addresses_file_name);

            return path_time_file_name_addresses;
        }

        /// <summary>Get server file name that shall be the same as local file name.</summary>
        public static string GetServerFileName(string i_full_file_name)
        {
            return Path.GetFileName(i_full_file_name);
        }

        /// <summary>Get backup server file name that shall be the same as local file name but in a subdirectory.</summary>
        public static string GetServerBackupFileName(string i_full_file_name, string i_sub_dir)
        {
            string file_without_path = Path.GetFileName(i_full_file_name);
            string file_with_backup_dir = i_sub_dir + @"/" + file_without_path;

            return file_with_backup_dir;
        }

        /// <summary>Get the path to a subdirectory. Create the directory if not existing.</summary>
        public static string SubDirectory(string i_subdir_name, string i_exe_directory)
        {
            string sub_directory = Path.Combine(i_exe_directory, i_subdir_name);

            if (!Directory.Exists(sub_directory))
            {
                Directory.CreateDirectory(sub_directory);
            }

            return sub_directory;
        }

        /// <summary>Get full name to a file on a subdirectory.</summary>
        public static string SubDirectoryFileName(string i_file_name, string i_subdir_name, string i_exe_directory)
        {
            string path_file_name = Path.Combine(SubDirectory(i_subdir_name, i_exe_directory), i_file_name);

            return path_file_name;
        }

        /// <summary>Create a file if missing, i.e. create a copy of the file from the resources (string).</summary>
        public static void CreateFileFromResourcesStringIfMissing(string i_full_file_name, string i_file_resources)
        {
            if (File.Exists(i_full_file_name))
            {
                return; // File exists already. Do nothing.
            }

            string o_error = "";

            try
            {
                using (FileStream fileStream = new FileStream(i_full_file_name, FileMode.Create))
                // Without System.Text.Encoding.Default there are problems with ä ö ü
                using (StreamWriter stream_writer = new StreamWriter(fileStream, System.Text.Encoding.Default))
                {
                    stream_writer.Write(i_file_resources);

                    stream_writer.Close();
                }
            }

            catch (FileNotFoundException) { o_error = "File not found"; return; }
            catch (DirectoryNotFoundException) { o_error = "Directory not found"; return; }
            catch (InvalidOperationException) { o_error = "Invalid operation"; return; }
            catch (InvalidCastException) { o_error = "invalid cast"; return; }
            catch (Exception e)
            {
                o_error = " Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!";
                return;
            }
        }

        /// <summary>
        /// Get files with given extensions
        /// </summary>
        /// <param name="i_extension">Array of extensions (with point)</param>
        /// <param name="i_directory">Search directory</param>
        ///  <param name="i_reverse">Reverse order of output array</param>
        /// <param name="o_file_names">Array of found files with paths</param>
        /// <returns>false if directory not exists or the input array of extensions is empty</returns>
        static public bool GetFilesDirectory(string[] i_extensions, string i_directory, bool i_reverse, out string[] o_file_names)
        {
            ArrayList files_string_array = new ArrayList();
            o_file_names = (string[])files_string_array.ToArray(typeof(string));

            if (!Directory.Exists(i_directory))
            {
                return false;
            }

            for (int i_ext = 0; i_ext < i_extensions.Length; i_ext++)
            {
                string current_ext = i_extensions[i_ext];

                string[] files_ext = Directory.GetFiles(i_directory, "*" + current_ext);

                foreach (string file_ext in files_ext)
                {
                    files_string_array.Add(file_ext);
                }
            }

            if (i_reverse)
            {
                files_string_array.Reverse();
            }

            o_file_names = (string[])files_string_array.ToArray(typeof(string));

            return true;
        } // GetFilesDirectory


    } // FileUtil
}

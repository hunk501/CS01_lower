using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace CS01
{
    class Files
    {

        public string APP_NAME = "CS02";
        public string REGEDIT_PATH = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";

        public string getCurrentUsername()
        {
            string user = Environment.UserName;
            return user;
        }

        public bool addRegistry(string executable_path)
        {            
            try
            {                
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(this.REGEDIT_PATH, true))
                {
                    key.SetValue(this.APP_NAME, "\"" + executable_path + "\"");
                }

                //MessageBox.Show("Success", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return true;
            }
            catch (Exception e)
            {
                //MessageBox.Show("Error:" + e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.WriteLine("ERROR: "+ e.Message);
                return false;
            }
        }

    }
}

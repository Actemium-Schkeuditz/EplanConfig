using System;
using System.Security.Principal;
using System.Windows;

namespace EplanCMHL
{
    class EplanNutzer
    {
        private static void Main()

        {
            string domainName = @"VED\ORG-AC-DESK001-BU00103-Users";// CMSK001SVIS03 -SK001-10-10-RW";

            if (IsUserInGroup(domainName))
                MessageBox.Show("Treffer");
            else
                MessageBox.Show("kein Treffer")
                ;

        }

        public static bool IsUserInGroup(string groupName)
        {
            //string testGruppName = " ";

            foreach (IdentityReference group in WindowsIdentity.GetCurrent().Groups)
            {
                //MessageBox.Show((group.Translate(typeof(NTAccount)).Value));
                // testGruppName = group.Translate(typeof(NTAccount)).Value;
                // MessageBox.Show(testGruppName);
                if (groupName == group.Translate(typeof(NTAccount)).Value)
                {

                    return true;
                }
            }
            return false;
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Windows.Forms;
using System.IO;
using System.Security.Principal;
using System.Security.AccessControl;

namespace FocasSmartDataCollection
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
        }

        public override void Install(System.Collections.IDictionary stateSaver)
        {
            base.Install(stateSaver);
            Form frm = new AppConfigModifier(this.Context);
            frm.BringToFront();
            frm.ShowDialog();           

            try
            {
                var filepath = this.Context.Parameters["assemblypath"];
                filepath = Path.GetDirectoryName(filepath);
                //filepath = Path.Combine(filepath, "Logs");

                // This gets the "Authenticated Users" group, no matter what it's called
                SecurityIdentifier sid = new SecurityIdentifier(WellKnownSidType.AuthenticatedUserSid, null);

                // Create the rules
                FileSystemAccessRule writerule = new FileSystemAccessRule(sid, FileSystemRights.FullControl, AccessControlType.Allow);

                if (!string.IsNullOrEmpty(filepath) && Directory.Exists(filepath))
                {
                    // Get your file's ACL
                    DirectorySecurity fsecurity = Directory.GetAccessControl(filepath);

                    // Add the new rule to the ACL
                    fsecurity.AddAccessRule(writerule);

                    // Set the ACL back to the file
                    Directory.SetAccessControl(filepath, fsecurity);


                    // *** Always allow objects to inherit on a directory                
                    InheritanceFlags iFlags = InheritanceFlags.ObjectInherit;
                    iFlags = InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit;

                    // *** Add Access rule for the inheritance
                    var AccessRule = new FileSystemAccessRule(sid, FileSystemRights.FullControl, iFlags, PropagationFlags.InheritOnly,AccessControlType.Allow);
                    bool Result = false;
                    fsecurity.ModifyAccessRule(AccessControlModification.Add, AccessRule, out Result);                    
                    Directory.SetAccessControl(filepath, fsecurity);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        static bool SetAcl(string destinationDirectory)
        {
            FileSystemRights Rights = (FileSystemRights)0;
            Rights = FileSystemRights.FullControl;

            // *** Add Access Rule to the actual directory itself
            FileSystemAccessRule AccessRule = new FileSystemAccessRule("Users", Rights,
                                        InheritanceFlags.None,
                                        PropagationFlags.NoPropagateInherit,
                                        AccessControlType.Allow);

            DirectoryInfo Info = new DirectoryInfo(destinationDirectory);
            DirectorySecurity Security = Info.GetAccessControl(AccessControlSections.Access);

            bool Result = false;
            Security.ModifyAccessRule(AccessControlModification.Set, AccessRule, out Result);

            if (!Result)
                return false;

            // *** Always allow objects to inherit on a directory
            InheritanceFlags iFlags = InheritanceFlags.ObjectInherit;
            iFlags = InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit;

            // *** Add Access rule for the inheritance
            AccessRule = new FileSystemAccessRule("Users", Rights,
                                        iFlags,
                                        PropagationFlags.InheritOnly,
                                        AccessControlType.Allow);
            Result = false;
            Security.ModifyAccessRule(AccessControlModification.Add, AccessRule, out Result);

            if (!Result)
                return false;

            Info.SetAccessControl(Security);

            return true;
        }
    }
}
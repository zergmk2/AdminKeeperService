using NLog;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Management;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AdminKeeperService
{
    public partial class AdminKeeper : ServiceBase
    {
        private const int THREADSLEEPINTERVAL = 30 * 1000 * 60; //30 minutes
        private const string ADMINGROUPNAME = "Administrators";
        private Thread AdminKeepingThread;
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public AdminKeeper()
        {
            //InitializeComponent();
            AdminKeepingThread = new Thread(new ThreadStart(AdminKeepingThreadFunc));
        }

        protected override void OnStart(string[] args)
        {
            logger.Info("OnStart invoked");
            AdminKeepingThread.Start();
        }

        protected override void OnStop()
        {
            logger.Info("OnStop invoked");
        }

        public void AdminKeepingThreadFunc()
        {
            while (true)
            {
                try
                {
                    logger.Info("Admin keeping thread awaked.");
                    string userName = GetCurrentUser();
                    if (userName.Contains(@"\"))
                    {
                       userName = userName.Substring(userName.IndexOf(@"\") + 1);
                    }
                    logger.Info("Get current user's name : " + userName);
                    List<string> AdminUsers = GetUsersInAdministrators();
                    if (AdminUsers != null && AdminUsers.Count > 0)
                    {
                        if (AdminUsers.Exists(x => x.Equals(userName, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            logger.Info($"Found {userName} in admin group");
                        }
                        else
                        {
                            logger.Info($"Can't find {userName} in admin group");
                            AddToGroup(userName);
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Error("Exception occurred in AdminKeepingThreadFunc : " + ex);
                }
                Thread.Sleep(THREADSLEEPINTERVAL);
            }
        }

        private string GetCurrentUser()
        {
            string userName = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT UserName FROM Win32_ComputerSystem");
                ManagementObjectCollection collection = searcher.Get();
                userName = (string)collection.Cast<ManagementBaseObject>().First()["UserName"];
            }
            catch (Exception ex)
            {
                logger.Error("Exception occurred when get current user's name : " + ex);
            }
            return userName;
        }

        private List<string> GetUsersInAdministrators()
        {
            List<string> users = new List<string>();
            try
            {
                DirectoryEntry localMachine = new DirectoryEntry("WinNT://" + Environment.MachineName);
                DirectoryEntry admGroup = localMachine.Children.Find(ADMINGROUPNAME, "group");
                object members = admGroup.Invoke("members", null);
                foreach (object groupMember in (IEnumerable)members)
                {
                    DirectoryEntry member = new DirectoryEntry(groupMember);
                    users.Add(member.Name);
                }
#if false
                using (PrincipalContext context = new PrincipalContext(ContextType.Machine))
                {
                    using (var group = GroupPrincipal.FindByIdentity(context, ADMINGROUPNAME))
                    {
                        if (group == null)
                        {
                            logger.Error(string.Format("{0} group doesn't exist", ADMINGROUPNAME));
                        }
                        else
                        {
                            var adminusers = group.GetMembers();
                            foreach (var user in adminusers)
                            {
                                UserPrincipal theUser = user as UserPrincipal;
                                if (theUser != null)
                                {
                                    logger.Info("++++++++++++++++++++++++++++++");
                                    logger.Info((user as UserPrincipal).DisplayName);
                                    logger.Info((user as UserPrincipal).SamAccountName);
                                    logger.Info((user as UserPrincipal).Name);
                                    logger.Info((user as UserPrincipal).LastLogon);
                                }
                            }
                        }

                    }
                }

#endif
            }
            catch (Exception ex)
            {
                logger.Error("Exception occurred when get all users in admin group : " + ex);
            }
            return users;
        }

        private void AddToGroup(string userName)
        {
            try
            {
                PrincipalContext systemContext = new PrincipalContext(ContextType.Machine, null);
                PrincipalContext domainContext = new PrincipalContext(ContextType.Domain, null);
                GroupPrincipal groupPrincipal = null;
                //Check if user object already exists
                logger.Info("Checking if User Exists.");
                //UserPrincipal usr = UserPrincipal.FindByIdentity(systemContext, ADMINUSERNAME);
                UserPrincipal domainusr = UserPrincipal.FindByIdentity(domainContext, userName);
                if (domainusr != null)
                {
                    logger.Info("Found domain user : " + userName);
                    groupPrincipal = GroupPrincipal.FindByIdentity(systemContext, ADMINGROUPNAME);

                    if (groupPrincipal != null)
                    {
                        //check if user is a member
                        logger.Info("Checking if itadmin is part of Administrators Group");
                        if (groupPrincipal.Members.Contains(domainusr))
                        {
                            logger.Info("Administrators already contains " + userName);
                            return;
                        }
                        //Adding the user to the group
                        logger.Info("Adding itadmin to Administrators Group");
                        groupPrincipal.Members.Add(domainusr);
                        groupPrincipal.Save();
                    }
                    else
                    {
                        logger.Info("Could not find the group Administrators");
                    }
                }
#if false


                //Create the new UserPrincipal object
                Console.WriteLine("Building User Information");
                UserPrincipal userPrincipal = new UserPrincipal(systemContext);
                userPrincipal.Name = ADMINUSERNAME;
                userPrincipal.DisplayName = "IT Administrative User";
                userPrincipal.PasswordNeverExpires = true;
                userPrincipal.SetPassword(ADMINUSERNAME);
                userPrincipal.Enabled = true;

                try
                {
                    Console.WriteLine("Creating New User");
                    userPrincipal.Save();
                }
                catch (Exception E)
                {
                    Console.WriteLine("Failed to create user.");
                    Console.WriteLine("Exception: " + E);

                    Console.WriteLine();
                    Console.WriteLine("Press Any Key to Continue");
                    Console.ReadLine();
                    return;
                }


#endif

                logger.Info("Cleaning Up");
                groupPrincipal.Dispose();
                //userPrincipal.Dispose();
                systemContext.Dispose();
                domainContext.Dispose();
            }
            catch (Exception ex)
            {
                logger.Error("Exception occurred when get all users in admin group : " + ex);

            }
        }

        private void AddLocalAdminGroup()
        {
            // Start the child process.
            Process p = new Process();
            // Redirect the output stream of the child process.
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.FileName = "net";
            p.StartInfo.Arguments = "localgroup administrators lu_cao /add";
            p.Start();
            // Do not wait for the child process to exit before
            // reading to the end of its redirected stream.
            // p.WaitForExit();
            // Read the output stream first and then wait.
            string output = p.StandardOutput.ReadToEnd();
            p.WaitForExit();

            logger.Info(output);
        }
    }
}

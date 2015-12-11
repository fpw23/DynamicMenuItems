//------------------------------------------------------------------------------
// <copyright file="DynamicMenu.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------
using EnvDTE;
using EnvDTE80;
using System.ComponentModel.Design;
using System;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using DynamicMenuItems.Classes;
using System.Collections.Generic;
using System.Linq;

namespace DynamicMenuItems
{

    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class DynamicMenu
    {
        private List<Tuple<int, string, string>> _menuItems = new List<Tuple<int, string, string>>();
        private DynamicItemMenuCommand _rootMenuItem;

        private int idCount = 0;
        private DTE2 dte2;
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("acac0ca9-d496-4208-9d28-07e6c887f79b");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="DynamicMenu"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private DynamicMenu(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;


            this._menuItems.Add(Tuple.Create(this.idCount++, "Classes", "Do Class 1"));
            this._menuItems.Add(Tuple.Create(this.idCount++, "Classes", "Do Class 2"));
            this._menuItems.Add(Tuple.Create(this.idCount++, "Classes", "Do Class 3"));
            this._menuItems.Add(Tuple.Create(this.idCount++, "Resources", "Do Resource 1"));
            this._menuItems.Add(Tuple.Create(this.idCount++, "Resources", "Do Resource 2"));
            this._menuItems.Add(Tuple.Create(this.idCount++, "Resources", "Do Resource 3"));

            dte2 = (DTE2)this.ServiceProvider.GetService(typeof(DTE));

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                // Add the DynamicItemMenuCommand for the expansion of the root item into N items at run time. 
                CommandID dynamicItemRootId = new CommandID(new Guid(DynamicMenuPackageGuids.guidDynamicMenuPackageCmdSet), (int)DynamicMenuPackageGuids.cmdidMyCommand);
                this._rootMenuItem = new DynamicItemMenuCommand(dynamicItemRootId,
                    IsValidDynamicItem,
                    OnInvokedDynamicItem,
                    OnBeforeQueryStatusDynamicItem);
                commandService.AddCommand(this._rootMenuItem);

                var menuContainer = new DynamicItemMenuContainer(commandService, this.dte2, this._menuItems);

            }




        }


        private void OnInvokedDynamicItem(object sender, EventArgs args)
        {
            DynamicItemMenuCommand invokedCommand = (DynamicItemMenuCommand)sender;

            UIHierarchy uih = dte2.ToolWindows.SolutionExplorer;
            Array selectedItems = (Array)uih.SelectedItems;
            UIHierarchyItem selectedItem = selectedItems.GetValue(0) as UIHierarchyItem;
            var testName = "";
            ProjectItem prjItem = selectedItem.Object as ProjectItem;
            if (prjItem == null)
            {
                Project prj = selectedItem.Object as Project;

                if (prj == null)
                {
                    return;
                }
                else
                {
                    testName = prj.Name;
                }
            }
            else
            {
                testName = prjItem.Name;
            }

            var matches = this._menuItems.Where(c => c.Item2 == testName).OrderBy(c => c.Item1).ToList();

            if (matches.Count == 0)
            {
                return;
            }

            bool isRootItem = (invokedCommand.MatchedCommandId == 0);

            // The index is set to 1 rather than 0 because the Solution.Projects collection is 1-based.
            int indexForDisplay = (isRootItem ? 0 : (invokedCommand.MatchedCommandId - (int)DynamicMenuPackageGuids.cmdidMyCommand));

            var match = matches[indexForDisplay];

            if (match == null)
            {
                return;
            }

            var newId = this.idCount++;

            this._menuItems.Add(Tuple.Create(newId, "ConsoleApplication4", string.Concat("Text From Code: ", newId)));

            string message = string.Format(CultureInfo.CurrentCulture, "Woo Hoo, you clicked item {0} - {1}", match.Item1, match.Item3);
            string title = "DynamicMenu";

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.ServiceProvider,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private bool IsValidDynamicItem(int commandId)
        {
            UIHierarchy uih = dte2.ToolWindows.SolutionExplorer;
            Array selectedItems = (Array)uih.SelectedItems;
            UIHierarchyItem selectedItem = selectedItems.GetValue(0) as UIHierarchyItem;
            var testName = "";
            ProjectItem prjItem = selectedItem.Object as ProjectItem;
            if (prjItem == null)
            {
                Project prj = selectedItem.Object as Project;

                if (prj == null)
                {

                    return false;
                }
                else
                {
                    testName = prj.Name;
                }
            }
            else
            {
                testName = prjItem.Name;
            }

            var matchCount = this._menuItems.Where(c => c.Item2 == testName).OrderBy(c => c.Item1).Count();

            if (matchCount == 0)
            {
                return false;
            }

            //System.Diagnostics.Debug.WriteLine("Is Valid Dynamic Item -- Command Id: {0}, {1}, {2}", commandId, (int)DynamicMenuPackageGuids.cmdidMyCommand, commandId - (int)DynamicMenuPackageGuids.cmdidMyCommand);

            // The match is valid if the command ID is >= the id of our root dynamic start item 
            // and the command ID minus the ID of our root dynamic start item
            // is less than or equal to the number of projects in the solution.
            return (commandId >= (int)DynamicMenuPackageGuids.cmdidMyCommand) && ((commandId - (int)DynamicMenuPackageGuids.cmdidMyCommand) < matchCount);
        }



        private void OnBeforeQueryStatusDynamicItem(object sender, EventArgs args)
        {
            DynamicItemMenuCommand matchedCommand = (DynamicItemMenuCommand)sender;

            UIHierarchy uih = dte2.ToolWindows.SolutionExplorer;
            Array selectedItems = (Array)uih.SelectedItems;
            UIHierarchyItem selectedItem = selectedItems.GetValue(0) as UIHierarchyItem;
            var testName = "";
            ProjectItem prjItem = selectedItem.Object as ProjectItem;
            if (prjItem == null)
            {
                Project prj = selectedItem.Object as Project;

                if (prj == null)
                {
                    matchedCommand.Enabled = false;
                    matchedCommand.Visible = false;
                    matchedCommand.MatchedCommandId = 0;
                    return;
                }
                else
                {
                    testName = prj.Name;
                }
            }
            else
            {
                testName = prjItem.Name;
            }

            var matches = this._menuItems.Where(c => c.Item2 == testName).OrderBy(c => c.Item1).ToList();

            if (matches.Count == 0)
            {
                matchedCommand.Enabled = false;
                matchedCommand.Visible = false;
                matchedCommand.MatchedCommandId = 0;
                return;
            }

            matchedCommand.Enabled = true;
            matchedCommand.Visible = true;

            // Find out whether the command ID is 0, which is the ID of the root item.
            // If it is the root item, it matches the constructed DynamicItemMenuCommand,
            // and IsValidDynamicItem won't be called.
            bool isRootItem = (matchedCommand.MatchedCommandId == 0);

            // The index is set to 1 rather than 0 because the Solution.Projects collection is 1-based.
            int indexForDisplay = (isRootItem ? 0 : (matchedCommand.MatchedCommandId - (int)DynamicMenuPackageGuids.cmdidMyCommand));

            matchedCommand.Text = matches[indexForDisplay].Item3;
            matchedCommand.MatchedCommandId = 0;

            System.Diagnostics.Debug.WriteLine("On Before Query Status Dynamic Item -- {0}, {1}, {2}", indexForDisplay, matchedCommand.MatchedCommandId, matchedCommand.Text);

        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static DynamicMenu Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new DynamicMenu(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            string title = "DynamicMenu";

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.ServiceProvider,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}

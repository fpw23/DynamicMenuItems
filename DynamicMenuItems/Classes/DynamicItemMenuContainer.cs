using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DynamicMenuItems.Classes
{
    public class DynamicItemMenuContainer : OleMenuCommand
    {
        private EnvDTE80.DTE2 _dte;
        private List<Tuple<int, string, string>> _menuItems;

        public DynamicItemMenuContainer(OleMenuCommandService mcs, EnvDTE80.DTE2 dte, List<Tuple<int, string, string>> menuItems)
            : base(Callback, new CommandID(Guid.Parse(DynamicMenuPackageGuids.guidDynamicMenuPackageCmdSet), (int)DynamicMenuPackageGuids.MyMenuController))
        {
            this._dte = dte;
            this._menuItems = menuItems;
            mcs.AddCommand(this);
            this.BeforeQueryStatus += DynamicItemMenuContainer_BeforeQueryStatus;

        }

        private void DynamicItemMenuContainer_BeforeQueryStatus(object sender, EventArgs e)
        {
            UIHierarchy uih = this._dte.ToolWindows.SolutionExplorer;
            Array selectedItems = (Array)uih.SelectedItems;
            UIHierarchyItem selectedItem = selectedItems.GetValue(0) as UIHierarchyItem;
            var testName = "";
            ProjectItem prjItem = selectedItem.Object as ProjectItem;
            if (prjItem == null)
            {
                Project prj = selectedItem.Object as Project;

                if (prj == null)
                {
                    this.Visible = false;
                    return;
                } else
                {
                    testName = prj.Name;
                }
            } else
            {
                testName = prjItem.Name;
            }

            var matchCount = this._menuItems.Where(c => c.Item2 == testName).OrderBy(c => c.Item1).Count();

            if (matchCount == 0)
            {
                this.Visible = false;
                return;
            }

            this.Visible = true;
        }

        private static void Callback(object sender, EventArgs e)
        {
            


        }

    }
}

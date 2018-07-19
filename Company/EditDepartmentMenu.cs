using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AVEVA.CDMS.WebApi;
using AVEVA.CDMS.Server;
using System.Collections;

namespace AVEVA.CDMS.HXEPC_Plugins
{
    internal class EditDepartmentMenu : ExWebMenu
    {

        /// <summary>
        /// 决定菜单的状态
        /// </summary>
        /// <returns></returns>
        public override enWebMenuState MeasureMenuState()
        {
            try
            {
                Project project = base.SelProjectList[0];
                if (project != null && project.TempDefn.KeyWord == "PRODOCUMENTADMIN")
                {
                    return enWebMenuState.Enabled;
                }
            }
            catch { }
            return enWebMenuState.Hide;
        }

    }
}
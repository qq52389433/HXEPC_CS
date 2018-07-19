using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AVEVA.CDMS.WebApi;
using AVEVA.CDMS.Server;
using System.Collections;

namespace AVEVA.CDMS.HXEPC_Plugins
{
    internal class AddCompanyMenu : ExWebMenu
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
                if (project != null && project.ParentProject != null 
                    && (project.ParentProject.Code == "收文" || project.ParentProject.Code == "发文"
                        || project.ParentProject.Description == "收文" || project.ParentProject.Description == "发文"
                        )
                    && project.ParentProject.TempDefn.Code == "COM_SUBDOCUMENT")
                {
                    return enWebMenuState.Enabled;
                }
            }
            catch { }
            return enWebMenuState.Hide;
        }

    }
}
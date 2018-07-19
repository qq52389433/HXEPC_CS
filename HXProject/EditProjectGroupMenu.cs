using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AVEVA.CDMS.WebApi;
using AVEVA.CDMS.Server;

namespace AVEVA.CDMS.HXEPC_Plugins
{
    internal class EditProjectGroupMenu : ExWebMenu
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
                if (project != null && project.TempDefn.KeyWord == "HXNY_DOCUMENTSYSTEM")
                {
                    return enWebMenuState.Enabled;
                }
            }
            catch { }
            return enWebMenuState.Hide;
        }
    }

}

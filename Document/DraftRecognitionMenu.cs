using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AVEVA.CDMS.WebApi;
using AVEVA.CDMS.Server;
using System.Collections;

namespace AVEVA.CDMS.HXEPC_Plugins
{
    /// <summary>
    /// 起草认质认价单
    /// </summary>
    internal class DraftRecognitionMenu : ExWebMenu
    {
        private Project SelectedProject;

        /// <summary>
        /// 决定菜单的状态
        /// </summary>
        /// <returns></returns>
        public override enWebMenuState MeasureMenuState()
        {
            try
            {
                if (base.SelProjectList.Count <= 0)
                {
                    return enWebMenuState.Hide;
                }

                Project project = base.SelProjectList[0];   //选择目录
                Project ProfessionProject = null;           //专业

                if (project != null)
                {
                    Project parentProject = project;

                    bool flag = false;
                    //while (parentProject != null)
                    //{
                        if (project.Code == "认质认价" || project.Description == "认质认价")
                        {
                            flag = true;
                        }
                    if (flag)
                    {
                        return enWebMenuState.Enabled;
                    }


                }
            }
            catch { }
            return enWebMenuState.Hide;
        }

    }
}
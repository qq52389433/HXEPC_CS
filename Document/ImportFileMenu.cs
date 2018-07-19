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
    /// 起草红头文
    /// </summary>
    internal class ImportFileInfoMenu : ExWebMenu
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

                    //if (
                    //   //项目管理类
                    //   ((parentProject != null && parentProject.ParentProject != null &&
                    //   (parentProject.Code == "信函" || parentProject.Description == "信函") && 
                    //   (parentProject.ParentProject.Code == "发文" || parentProject.ParentProject.Description == "发文")) ||
                    //   (parentProject.ParentProject != null && parentProject.ParentProject.ParentProject != null &&
                    //   (parentProject.ParentProject.Code == "信函" || parentProject.ParentProject.Description== "信函") &&
                    //   (parentProject.ParentProject.ParentProject.Code == "发文" || parentProject.ParentProject.ParentProject.Description== "发文"))
                    //  ) ||
                    //   //运营管理类
                    //   ((parentProject != null && parentProject.ParentProject != null &&
                    //    (parentProject.Code == "发文" || parentProject.Description == "发文") && 
                    //      (parentProject.ParentProject.Code == "信函" || parentProject.ParentProject.Description == "信函")) ||
                    //    (parentProject.ParentProject != null && parentProject.ParentProject.ParentProject != null &&
                    //   ( parentProject.ParentProject.Code == "发文" || parentProject.ParentProject.Description == "发文" )&&
                    //   (parentProject.ParentProject.ParentProject.Code == "信函" || 
                    //     parentProject.ParentProject.ParentProject.Description == "信函"))
                    //   ))
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
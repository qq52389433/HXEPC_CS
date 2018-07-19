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
    internal class DraftTransmittalCNMenu : ExWebMenu
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
                    //在通信管理文件夹下，才可以起草
                    Project commProj = CommonFunction.getParentProjectByTempDefn(project, "PRO_COMMUNICATION");
                    if (commProj == null)
                    {
                        return enWebMenuState.Hide;
                    }

                    Project parentProject = project;

                    bool flag = false;
                    //while (parentProject != null)
                    //{
                    // if (parentProject.Code == "文件传递单" && parentProject.ParentProject.Code == "发文")
                    //if ((parentProject != null && parentProject.ParentProject != null &&
                    // parentProject.Code == "文件传递单" && parentProject.ParentProject.Code == "发文") ||
                    // (parentProject.ParentProject != null && parentProject.ParentProject.ParentProject != null &&
                    // parentProject.ParentProject.Code == "文件传递单" &&
                    // parentProject.ParentProject.ParentProject.Code == "发文")
                    //)
                    // if (
                    // //项目管理类
                    // ((parentProject != null && parentProject.ParentProject != null &&
                    // parentProject.Code == "文件传递单" && parentProject.ParentProject.Code == "发文") ||
                    // (parentProject.ParentProject != null && parentProject.ParentProject.ParentProject != null &&
                    // parentProject.ParentProject.Code == "文件传递单" &&
                    // parentProject.ParentProject.ParentProject.Code == "发文")
                    //) ||
                    // //运营管理类
                    // ((parentProject != null && parentProject.ParentProject != null &&
                    //  parentProject.Code == "发文" && parentProject.ParentProject.Code == "文件传递单") ||
                    //  (parentProject.ParentProject != null && parentProject.ParentProject.ParentProject != null &&
                    //  parentProject.ParentProject.Code == "发文" &&
                    //  parentProject.ParentProject.ParentProject.Code == "文件传递单")
                    // ))
                    if (
                       //项目管理类
                       ((parentProject != null && parentProject.ParentProject != null &&
                       (parentProject.Code == "文件传递单" || parentProject.Description == "文件传递单") &&
                       (parentProject.ParentProject.Code == "发文" || parentProject.ParentProject.Description == "发文")) ||
                       (parentProject.ParentProject != null && parentProject.ParentProject.ParentProject != null &&
                       (parentProject.ParentProject.Code == "文件传递单" || parentProject.ParentProject.Description == "文件传递单") &&
                       (parentProject.ParentProject.ParentProject.Code == "发文" || parentProject.ParentProject.ParentProject.Description == "发文"))
                      ) ||
                       //运营管理类
                       ((parentProject != null && parentProject.ParentProject != null &&
                        (parentProject.Code == "发文" || parentProject.Description == "发文") &&
                          (parentProject.ParentProject.Code == "文件传递单" || parentProject.ParentProject.Description == "文件传递单")) ||
                        (parentProject.ParentProject != null && parentProject.ParentProject.ParentProject != null &&
                       (parentProject.ParentProject.Code == "发文" || parentProject.ParentProject.Description == "发文") &&
                       (parentProject.ParentProject.ParentProject.Code == "文件传递单" ||
                         parentProject.ParentProject.ParentProject.Description == "文件传递单"))
                       ))
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
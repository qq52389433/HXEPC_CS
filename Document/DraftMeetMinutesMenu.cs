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
    /// 起草会议纪要
    /// </summary>
    internal class DraftMeetMinutesMenu : ExWebMenu
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
                    // if (parentProject.Code == "会议纪要" && parentProject.ParentProject.Code == "发文")
                    //if ((parentProject != null && parentProject.ParentProject != null &&
                    //      parentProject.Code == "会议纪要" && parentProject.ParentProject.Code == "发文") ||
                    //     (parentProject.ParentProject != null && parentProject.ParentProject.ParentProject != null &&
                    //      parentProject.ParentProject.Code == "会议纪要" &&
                    //      parentProject.ParentProject.ParentProject.Code == "发文")
                    //)
                    // if (
                    // //项目管理类
                    // ((parentProject != null && parentProject.ParentProject != null &&
                    // parentProject.Code == "会议纪要" && parentProject.ParentProject.Code == "发文") ||
                    // (parentProject.ParentProject != null && parentProject.ParentProject.ParentProject != null &&
                    // parentProject.ParentProject.Code == "会议纪要" &&
                    // parentProject.ParentProject.ParentProject.Code == "发文")
                    //) ||
                    // //运营管理类
                    // ((parentProject != null && parentProject.ParentProject != null &&
                    //  parentProject.Code == "发文" && parentProject.ParentProject.Code == "会议纪要") ||
                    //  (parentProject.ParentProject != null && parentProject.ParentProject.ParentProject != null &&
                    //  parentProject.ParentProject.Code == "发文" &&
                    //  parentProject.ParentProject.ParentProject.Code == "会议纪要")
                    // ))
                    if (
                       //项目管理类
                       ((parentProject != null && parentProject.ParentProject != null &&
                       (parentProject.Code == "会议纪要" || parentProject.Description == "会议纪要") &&
                       (parentProject.ParentProject.Code == "发文" || parentProject.ParentProject.Description == "发文")) ||
                       (parentProject.ParentProject != null && parentProject.ParentProject.ParentProject != null &&
                       (parentProject.ParentProject.Code == "会议纪要" || parentProject.ParentProject.Description == "会议纪要") &&
                       (parentProject.ParentProject.ParentProject.Code == "发文" || parentProject.ParentProject.ParentProject.Description == "发文"))
                      ) ||
                       //运营管理类
                       ((parentProject != null && parentProject.ParentProject != null &&
                        (parentProject.Code == "发文" || parentProject.Description == "发文") &&
                          (parentProject.ParentProject.Code == "会议纪要" || parentProject.ParentProject.Description == "会议纪要")) ||
                        (parentProject.ParentProject != null && parentProject.ParentProject.ParentProject != null &&
                       (parentProject.ParentProject.Code == "发文" || parentProject.ParentProject.Description == "发文") &&
                       (parentProject.ParentProject.ParentProject.Code == "会议纪要" ||
                         parentProject.ParentProject.ParentProject.Description == "会议纪要"))
                       ))
                    {
                        flag = true;
                    }

                    ////查找阶段和专业
                    //bool flag = false;
                    //while (parentProject != null)
                    //{
                    //    if (parentProject.Code == "收发文")
                    //    {
                    //        flag = true;
                    //    }
                    //    if ((parentProject.TempDefn != null) && (parentProject.TempDefn.Code == "PROFESSION"))
                    //    {
                    //        ProfessionProject = parentProject;
                    //        break;
                    //    }
                    //    parentProject = parentProject.ParentProject;
                    //}
                    //if (ProfessionProject == null || flag == false)
                    //{
                    //    return enWebMenuState.Hide;
                    //}


                    ////判断是否为主设或者专业设计人
                    //flag = false;
                    //string valueByKeyWord = ProfessionProject.GetValueByKeyWord("PROFESSIONOWNER");  //主设
                    //if (!string.IsNullOrEmpty(valueByKeyWord) && (valueByKeyWord.ToLower() == project.dBSource.LoginUser.Code.ToLower()))
                    //{
                    //    flag = true;
                    //}
                    //else
                    //{
                    //    valueByKeyWord = ProfessionProject.GetValueByKeyWord("PROFESSIONDESIGN");      //专业设计人
                    //    if (!string.IsNullOrEmpty(valueByKeyWord) && (valueByKeyWord.ToLower() == project.dBSource.LoginUser.Code.ToLower()))
                    //    {
                    //        flag = true;
                    //    }
                    //}
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
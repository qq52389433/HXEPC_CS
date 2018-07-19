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
    internal class DraftReportMenu : ExWebMenu
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
                    //if (parentProject.Code == "红头文" && parentProject.ParentProject.Code == "发文")
                    //if ((parentProject != null && parentProject.ParentProject != null &&
                    //      parentProject.Code == "红头文" && parentProject.ParentProject.Code == "发文") ||
                    //     (parentProject.ParentProject != null && parentProject.ParentProject.ParentProject != null &&
                    //      parentProject.ParentProject.Code == "红头文" &&
                    //      parentProject.ParentProject.ParentProject.Code == "发文")
                    //)
                    //if (
                    //   //项目管理类
                    //   ((parentProject != null && parentProject.ParentProject != null &&
                    //   (parentProject.Code == "红头文" || parentProject.Description == "红头文") && 
                    //   (parentProject.ParentProject.Code == "发文" || parentProject.ParentProject.Description == "发文")) ||
                    //   (parentProject.ParentProject != null && parentProject.ParentProject.ParentProject != null &&
                    //   (parentProject.ParentProject.Code == "红头文" || parentProject.ParentProject.Description== "红头文") &&
                    //   (parentProject.ParentProject.ParentProject.Code == "发文" || parentProject.ParentProject.ParentProject.Description== "发文"))
                    //  ) ||
                    //   //运营管理类
                    //   ((parentProject != null && parentProject.ParentProject != null &&
                    //    (parentProject.Code == "发文" || parentProject.Description == "发文") && 
                    //      (parentProject.ParentProject.Code == "红头文"|| parentProject.ParentProject.Description == "红头文")) ||
                    //    (parentProject.ParentProject != null && parentProject.ParentProject.ParentProject != null &&
                    //   ( parentProject.ParentProject.Code == "发文" || parentProject.ParentProject.Description == "发文" )&&
                    //   (parentProject.ParentProject.ParentProject.Code == "红头文" || 
                    //     parentProject.ParentProject.ParentProject.Description == "红头文"))
                     //  ))
                     if (parentProject.Description.IndexOf("通知")>=0  &&
                        parentProject.Description.IndexOf("请示") >= 0 && 
                        parentProject.Description.IndexOf("报告")>= 0 )
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
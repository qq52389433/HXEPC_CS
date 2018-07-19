using System;
using System.Collections;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Web;
using System.Web.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using AVEVA.CDMS.Server;
using AVEVA.CDMS.Common;
using AVEVA.CDMS.WebApi;
using System.Runtime.Serialization;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
//using System.Data.SQLite;
using LinqToDB;


//using System.Web.Script.Serialization;

namespace AVEVA.CDMS.HXEPC_Plugins
{
    public class EnterPoint
    {
        public static string PluginName = "HXEPC";

        public static void Init()
        {
            //WebApi.WebExploreEvent


            //添加流程按钮事件处理
            //记录本插件的唯一标记

            //记录是否已加载
            bool isLoad = false;

            foreach (WebWorkFlowEvent.Before_WorkFlow_SelectUsers_Event_Class EventClass in WebWorkFlowEvent.ListBeforeWFSelectUsers)
            {
                if (EventClass.PluginName == PluginName)
                {
                    isLoad = true;
                    break;
                }
            }

            if (isLoad==false)
            {
                //拖拽文件后处理
                //WebExploreEvent.OnAfterCreateNewObject += new WebExploreEvent.Explorer_AfterCreateNewObject(Document.OnAfterCreateNewObject);

                //拖拽文件后处理
                WebExploreEvent.Explorer_AfterCreateNewObject_Event AfterCreateNewObject = new WebExploreEvent.Explorer_AfterCreateNewObject_Event(Document.OnAfterCreateNewObject);
                WebExploreEvent.Explorer_AfterCreateNewObject_Event_Class Explorer_AfterCreateNewObject_Event_Class = new WebExploreEvent.Explorer_AfterCreateNewObject_Event_Class();
                Explorer_AfterCreateNewObject_Event_Class.Event = AfterCreateNewObject;
                Explorer_AfterCreateNewObject_Event_Class.PluginName = PluginName;
                WebExploreEvent.ListAfterCreateNewObject.Add(Explorer_AfterCreateNewObject_Event_Class);

                //添加流程按钮事件处理
                WebWorkFlowEvent.Before_WorkFlow_SelectUsers_Event BeforeWFSelectUsers =new WebWorkFlowEvent.Before_WorkFlow_SelectUsers_Event(AVEVA.CDMS.HXEPC_Plugins.EnterPoint.BeforeWF);
                WebWorkFlowEvent.Before_WorkFlow_SelectUsers_Event_Class Before_WorkFlow_SelectUsers_Event_Class  = new WebWorkFlowEvent.Before_WorkFlow_SelectUsers_Event_Class();
                Before_WorkFlow_SelectUsers_Event_Class.Event = BeforeWFSelectUsers;
                Before_WorkFlow_SelectUsers_Event_Class.PluginName = PluginName;
                WebWorkFlowEvent.ListBeforeWFSelectUsers.Add(Before_WorkFlow_SelectUsers_Event_Class);

                //添加流程文档创建者撤回流程事件处理
                WebWorkFlowEvent.Before_Revoke_WorkFlow_Event BeforeRevokeWorkFlow = new WebWorkFlowEvent.Before_Revoke_WorkFlow_Event(AVEVA.CDMS.HXEPC_Plugins.EnterPoint.RevokeWorkFlow);
                WebWorkFlowEvent.Before_Revoke_WorkFlow_Event_Class Before_Revoke_WorkFlow_Event_Class = new WebWorkFlowEvent.Before_Revoke_WorkFlow_Event_Class();
                Before_Revoke_WorkFlow_Event_Class.Event = BeforeRevokeWorkFlow;
                Before_Revoke_WorkFlow_Event_Class.PluginName = PluginName;
                WebWorkFlowEvent.ListBeforeRevokeWorkFlow.Add(Before_Revoke_WorkFlow_Event_Class);

                ////添加文档显示事件处理
                WebDocEvent.Before_Get_Doc_List_Event BeforeGetDocs = new WebDocEvent.Before_Get_Doc_List_Event(BeforeGetDocList);
                WebDocEvent.Before_Get_Doc_List_Event_Class Before_GetDocs_Event_Class = new WebDocEvent.Before_Get_Doc_List_Event_Class();
                Before_GetDocs_Event_Class.Event = BeforeGetDocs;
                Before_GetDocs_Event_Class.PluginName = PluginName;
                WebDocEvent.ListBeforeGetDocs.Add(Before_GetDocs_Event_Class);

                ////添加文件夹显示事件处理
                WebProjectEvent.Before_Get_Project_List_Event BeforeGetProjects = new WebProjectEvent.Before_Get_Project_List_Event(BeforeGetProjectList);
                WebProjectEvent.Before_Get_Project_List_Event_Class Before_GetProjects_Event_Class = new WebProjectEvent.Before_Get_Project_List_Event_Class();
                Before_GetProjects_Event_Class.Event = BeforeGetProjects;
                Before_GetProjects_Event_Class.PluginName = PluginName;
                WebProjectEvent.ListBeforeGetProjects.Add(Before_GetProjects_Event_Class);

                //添加获取文件夹图标事件处理
                WebProjectEvent.After_Get_Project_Icon_Event AfterGetProjectIcon = new WebProjectEvent.After_Get_Project_Icon_Event(AfterGetProjectIconFun);
                WebProjectEvent.After_Get_Project_Icon_Event_Class AfterGetProjectIcon_Event_Class = new WebProjectEvent.After_Get_Project_Icon_Event_Class();
                AfterGetProjectIcon_Event_Class.Event = AfterGetProjectIconFun;
                AfterGetProjectIcon_Event_Class.PluginName = PluginName;
                WebProjectEvent.ListAfterGetProjectIconEvent.Add(AfterGetProjectIcon_Event_Class);

                ////添加文档预览事件处理
                WebDocEvent.Before_Preview_Doc_Event BeforePreviewDocEvent = new WebDocEvent.Before_Preview_Doc_Event(BeforePreviewDoc);
                WebDocEvent.Before_Preview_Doc_Event_Class Before_Preview_Doc_Event_Class = new WebDocEvent.Before_Preview_Doc_Event_Class();
                Before_Preview_Doc_Event_Class.Event = BeforePreviewDocEvent;
                Before_Preview_Doc_Event_Class.PluginName = PluginName;
                WebDocEvent.ListBeforePreviewDoc.Add(Before_Preview_Doc_Event_Class);

                ////添加文档下载事件处理
                WebDocEvent.Before_Download_File_Event BeforeDownloadFileEvent = new WebDocEvent.Before_Download_File_Event(BeforeDownloadFile);
                WebDocEvent.Before_Download_File_Event_Class Before_Download_File_Event_Class = new WebDocEvent.Before_Download_File_Event_Class();
                Before_Download_File_Event_Class.Event = BeforeDownloadFileEvent;
                Before_Download_File_Event_Class.PluginName = PluginName;
                WebDocEvent.ListBeforeDownloadFile.Add(Before_Download_File_Event_Class);

                //添加选择用户事件处理
                WebWorkFlowEvent.Before_Select_User_Event BeforeSelectUserEvent = new WebWorkFlowEvent.Before_Select_User_Event(BeforeSelectUser);
                WebWorkFlowEvent.Before_Select_User_Event_Class Before_Select_User_Event_Class = new WebWorkFlowEvent.Before_Select_User_Event_Class();
                Before_Select_User_Event_Class.Event = BeforeSelectUserEvent;
                Before_Select_User_Event_Class.PluginName = PluginName;
                WebWorkFlowEvent.ListBeforeSelectUser.Add(Before_Select_User_Event_Class);
            }
        }

        public static List<ExWebMenu> CreateNewExMenu()
        {
            try
            {
                List<ExWebMenu> menuList = new List<ExWebMenu>();

                CreatePrjDocumentMenu createPrjDocumentMenu = new CreatePrjDocumentMenu
                {
                    MenuName = "生成立项单...",
                    MenuType = enWebMenuType.Single,
                    MenuPosition = enWebMenuPosition.TVProject

                };
                menuList.Add(createPrjDocumentMenu);

                DraftDocumentMenu sendDocumentMenu = new DraftDocumentMenu
                {
                    MenuName = "起草红头文...",
                    MenuType = enWebMenuType.Single,
                    MenuPosition = enWebMenuPosition.TVProject

                };
                menuList.Add(sendDocumentMenu);

                DraftMeetMinutesMenu draftMeetMinutesMenu = new DraftMeetMinutesMenu
                {
                    MenuName = "起草会议纪要...",
                    MenuType = enWebMenuType.Single,
                    MenuPosition = enWebMenuPosition.TVProject

                };
                menuList.Add(draftMeetMinutesMenu);


                DraftTransmittalCNMenu draftTransmittalCNMenu = new DraftTransmittalCNMenu
                {
                    MenuName = "起草文件传递单(中文)...",
                    MenuType = enWebMenuType.Single,
                    MenuPosition = enWebMenuPosition.TVProject

                };
                menuList.Add(draftTransmittalCNMenu);

                DraftLetterCNMenu draftLetterCNMenu = new DraftLetterCNMenu
                {
                    MenuName = "起草信函(中文)...",
                    MenuType = enWebMenuType.Single,
                    MenuPosition = enWebMenuPosition.TVProject

                };
                menuList.Add(draftLetterCNMenu);

                DraftRecognitionMenu draftRecognition = new DraftRecognitionMenu()
                {
                    MenuName = "起草认质认价单...",
                    MenuType = enWebMenuType.Single,
                    MenuPosition = enWebMenuPosition.TVProject

                };

                menuList.Add(draftRecognition);

                DraftReportMenu draftReport = new DraftReportMenu()
                {
                    MenuName = "起草报告、请示、通知...",
                    MenuType = enWebMenuType.Single,
                    MenuPosition = enWebMenuPosition.TVProject

                };

                menuList.Add(draftReport);

                AddCompanyMenu companyMenu = new AddCompanyMenu
                {
                    MenuName = "添加参建单位...",
                    MenuType = enWebMenuType.Single,
                    MenuPosition = enWebMenuPosition.TVProject

                };
                menuList.Add(companyMenu);
                
                EditCompanyMenu editCompanyMenu = new EditCompanyMenu
                {
                    MenuName = "编辑参建单位...",
                    MenuType = enWebMenuType.Single,
                    MenuPosition = enWebMenuPosition.TVProject

                };
                menuList.Add(editCompanyMenu);

                EditDepartmentMenu editDepartmentMenu = new EditDepartmentMenu
                {
                    MenuName = "编辑项目部门...",
                    MenuType = enWebMenuType.Single,
                    MenuPosition = enWebMenuPosition.TVProject

                };
                menuList.Add(editDepartmentMenu);

                EditProjectGroupMenu editProjectGroupMenu = new EditProjectGroupMenu
                {
                    MenuName = "编辑项目组...",
                    MenuType = enWebMenuType.Single,
                    MenuPosition = enWebMenuPosition.TVProject

                };
                menuList.Add(editProjectGroupMenu);

                EditProjectInfoMenu editProjectInfoMenu = new EditProjectInfoMenu
                {
                    MenuName = "编辑项目资料...",
                    MenuType = enWebMenuType.Single,
                    MenuPosition = enWebMenuPosition.TVProject

                };
                menuList.Add(editProjectInfoMenu);


                ImportFileInfoMenu importFileInfoMenu = new ImportFileInfoMenu
               {
                   MenuName = "导入文件...",
                   MenuType = enWebMenuType.Single,
                   MenuPosition = enWebMenuPosition.TVProject

               };
                menuList.Add(importFileInfoMenu);

                return menuList;
            }
            catch { }
            return null;
        }
        //显示文档列表前的事件，用于按条件筛选用户能看到的文档列表
        public static bool BeforeGetProjectList(ref List<Project> projectList) {
            if (projectList.Count <= 0)
            {
                return false;
            }
            try
            {

                #region 函数逻辑
                //  1.获取当前登录用户

                // 2.获取当前项目的单位用户组和项目部用户组

                //  3.判断当前用户是否在单位用户组里面

                //   4.判读当前用户是否在项目部用户组里面

                //  5.1.如果当前用户在单位用户组里面,判断父目录是否是项目根目录

                // 5.2.如果父目录是项目根目录，去掉除通信类文件夹和存档管理文件夹目录外的所有文件夹 
                #endregion



                Project project = projectList[0];

                if (project == null) return false;

 

                Project rootProj = CommonFunction.getParentProjectByTempDefn(project, "HXNY_DOCUMENTSYSTEM");

                if (rootProj != null)
                {
                    DBSource dbsource = project.dBSource;

                    //1.获取当前登录用户
                    User curUser = project.dBSource.LoginUser;

                    // 2.获取当前项目的单位用户组和项目部用户组

                    Server.Group unitGroup = dbsource.GetGroupByName(rootProj.Code + "_ALLUnit");
                    Server.Group projGroup = dbsource.GetGroupByName(rootProj.Code + "_ProGroup");

                    //  3.判断当前用户是否在单位用户组里面
                    bool isUnitSec = false;
                    bool isProjUser = false;
                    if (unitGroup.UserList.Contains(curUser)) {
                        isUnitSec = true;
                    }
                    if (projGroup.UserList.Contains(curUser))
                    {
                        isProjUser = true;
                    }

                    //  5.1.如果当前用户在单位用户组里面,判断父目录是否是项目根目录
                    if (isUnitSec)
                    {
                        if (project.ParentProject.TempDefn.KeyWord == "HXNY_DOCUMENTSYSTEM")
                        {
                            // 5.2.如果父目录是项目根目录，去掉除通信类文件夹和存档管理文件夹目录外的所有文件夹 
                            List<Project> resultProjs = new List<Project>();
                            foreach (Project proj in projectList) {
                                if ( proj.Description == "通信文件"|| proj.Description == "流程管理") {
                                    resultProjs.Add(proj);
                                }
                            }
                            projectList = resultProjs;
                        }
                    }
                }


            }
            catch { }
            return false;
        }

        //显示文档列表前的事件，用于按条件筛选用户能看到的文档列表
        public static bool BeforeGetDocList(ref List<Doc> docList,string filter) {
            //搜索文档
            //if (!string.IsNullOrEmpty(filter))
            //return SearchDocList(ref docList,   filter);

            if (docList.Count <= 0) {
                return false;
            }

            try
            {
                #region 判断是否是收发文目录，是否需要隐藏文件
                //判断是否是收发文目录，是否需要隐藏文件
                bool isRsProject = false;
                foreach (Doc doc in docList)
                {
                    try
                    {
                        string docTemp = "";
                        TempDefn doctempdefn = null;
                        if (doc.ShortCutDoc != null)
                        {
                            if (doc.ShortCutDoc.TempDefn == null)
                            {
                                continue;
                            }

                            doctempdefn = doc.ShortCutDoc.TempDefn;

                        }
                        else
                        {
                            doctempdefn = doc.TempDefn;
                        }
                        if (doctempdefn == null)
                        {
                            continue;
                        }
                        docTemp = doctempdefn.KeyWord;


                        if (docTemp == "CATALOGUING" || docTemp == "FILETRANSMIT" ||
                            docTemp == "MEETINGSUMMARY" || docTemp == "DOCUMENTFILE")
                        {
                            AttrData data;
                            //文文档模板
                            if ((data = doc.GetAttrDataByKeyWord("CA_ATTRTEMP")) != null)
                            {
                                string strData = data.ToString;
                                if (strData == "LETTERFILE" || strData == "FILETRANSMIT" ||
                                   strData == "MEETINGSUMMARY" || strData == "DOCUMENTFILE")
                                {
                                    isRsProject = true;
                                    break;
                                }
                            }
                        }
                    }
                    catch { }
                }

                Doc ddoc =docList.Find(d => d.Project != null);
                Project proj = ddoc.Project;

                if (proj == null) {
                    //防止在某些情况下文档找不到目录，把所有文档都显示出来
                    docList = new List<Doc>();
                    return false;
                }

                if (proj == null || proj.TempDefn == null) return false;

                Project commProj = CommonFunction.getParentProjectByTempDefn(proj, "PRO_COMMUNICATION");
                if (commProj == null) {
                    return false;
                }

                //if (ddoc.Project != null) {
                //    commProj = CommonFunction.getParentProjectByTempDefn(proj, "PRO_COMMUNICATION");
                //    if (commProj == null)
                //    {
                //        return false;
                //    }
                //}

                Project pproj = proj.ParentProject;
                DBSource dbsource = proj.dBSource;
                if (pproj == null) return false;

                if (!isRsProject || !(proj.TempDefn.KeyWord == "COM_COMTYPE" ||
                    pproj.Code == "收文" || pproj.Code == "发文" ||
                    pproj.Description == "收文" || pproj.Description == "发文"
                    ))
                {
                    return false;
                }
                #endregion

                User curUser = proj.dBSource.LoginUser;


                #region 筛选文档列表的步骤说明
                //1. 获取所有有流程（流程未完成），或者创建者是登录人的文档，缩小查询的范围

                //2 获取当前目录是收文目录还是发文目录

                //////////////////发文目录////////////////
                //2.1 如果是发文目录，就遍历文档

                //2.1.0 此步骤暂时忽略///////////////////////////////////////////////////////////////////////////////////////////////////////////////如果文档是快捷方式，就跳转到下一个文档（发文目录是实体文件 ，收文目录只是发文目录里的文档的快捷方式）

                //2.1.1获取文档的发文单位，并获取该发文单位的所有成员

                //2.2.2. 判断当前用户是否是发文单位的成员，不是发文单位的成员就跳到下一个文档

                //2.2.3 如果当前用户是当前文档的发文单位的成员，就把文档添加到文档列表

                //////////////////收文目录////////////////
                //2.2  如果是收文目录，就遍历文档

                //2.2.0 此步骤暂时忽略////////////////////////////////////////////////////////////////////////////////////////////////////////////// 如果文档不是快捷方式，就跳转到下一个文档

                //2.2.1 获取文档的收文单位和抄送单位

                //2.2.2 判断当前用户是否是收文单位或者抄送单位的成员，不是收文单位或者抄送单位的成员就跳到下一个文档

                //2.2.3 如果当前用户是当前文档的收文单位或者抄送单位的成员，就把文档添加到文档列表

                #endregion


                #region 1. 获取所有有流程（流程未完成），或者创建者是登录人的文档，缩小查询的范围
                //1. 获取所有有流程（流程未完成），或者创建者是登录人的文档，缩小查询的范围
                List<Doc> resultDocList = new List<Doc>();
                if ( string.IsNullOrEmpty(filter) && 
                    pproj.ParentProject != null && pproj.ParentProject.Description == "通信管理")
                {
                    //如果是通信管理目录下的收发文文档列表，就要隐藏掉流程已经完成的文档
                    foreach (Doc d in docList)
                    {
                        try
                        {

                            //if (((d.WorkFlow == null && d.Creater == curUser) ||
                            //        (d.WorkFlow != null && d.WorkFlow.O_WorkFlowStatus != enWorkFlowStatus.Finish
                            //          && d.WorkFlow.CuWorkState.CuWorkUser.User == curUser))
                            // || (d.ShortCutDoc != null &&
                            //       (d.ShortCutDoc.WorkFlow != null && d.ShortCutDoc.WorkFlow.O_WorkFlowStatus != enWorkFlowStatus.Finish
                            //           && d.ShortCutDoc.WorkFlow.CuWorkState.CuWorkUser.User == curUser))
                            //    )
                            if ((d.ShortCutDoc == null &&
                                        ((d.WorkFlow == null && d.Creater == curUser) ||
                                        (d.WorkFlow != null && d.WorkFlow.O_WorkFlowStatus != enWorkFlowStatus.Finish
                                          )))
                                 || (d.ShortCutDoc != null &&
                                       (d.ShortCutDoc.WorkFlow != null && d.ShortCutDoc.WorkFlow.O_WorkFlowStatus != enWorkFlowStatus.Finish
                                           ))
                                    )
                            {
                                List<WorkUser> wus = d.ShortCutDoc == null ?
                                    d.WorkFlow.CuWorkState.WorkUserList :
                                    d.ShortCutDoc.WorkFlow.CuWorkState.WorkUserList;

                                foreach (WorkUser wu in wus)
                                {
                                    if (wu.User != null && wu.User == curUser)
                                    {
                                        resultDocList.Add(d);
                                    }
                                }
                            }

                        }
                        catch { }
                    }
                }
                else
                {
                    resultDocList = docList;
                }
                #endregion

                #region 2. 获取当前目录是收文目录还是发文目录
                // 2. 获取当前目录是收文目录还是发文目录
                Project rsProject = CommonFunction.getParentProjectByTempDefn(proj, "COM_SUBDOCUMENT");

                if (rsProject == null)
                {
                    WebApi.CommonController.WebWriteLog(DateTime.Now.ToString() + ":" + "没有查找到" + proj.KeyWord + "," + proj.Code + "的收发文目录。");
                    docList = new List<Doc>();
                    return false;
                }
                string projectType = "";
                if (rsProject.Code == "收文" || rsProject.Description == "收文")
                {
                    projectType = "收文";
                }
                else if (rsProject.Code == "发文" || rsProject.Description == "发文")
                {
                    projectType = "发文";
                }
                else
                {
                    WebApi.CommonController.WebWriteLog(DateTime.Now.ToString() + ":" + "没有查找到" + proj.KeyWord + "," + proj.Code + "的收发文目录。");
                    docList = new List<Doc>();
                    return false;
                }
                #endregion

                List<Doc> reDocList = new List<Doc>();

                #region 2.1 如果是发文目录，就遍历文档
                if (projectType == "发文")
                {
                    AttrData data;
                    foreach (Doc docItem in resultDocList)
                    {
                        try
                        {
                            Doc doc = docItem;
                            bool isShort = false;
                            if (doc.ShortCutDoc != null)
                            {
                                doc = doc.ShortCutDoc;
                                isShort = true;
                            }

                            if (doc.Creater == curUser && isShort == false)
                            {
                                reDocList.Add(docItem);
                                continue;
                            }

                            #region 判断流程用户里面有没有当前用户
                            //if (doc.WorkFlow != null && doc.WorkFlow.CuWorkState != null) {
                            //    bool isAdd = false;
                            //    foreach (WorkUser wu in doc.WorkFlow.CuWorkState.WorkUserList)
                            //    {
                            //        if (wu.User == curUser)
                            //        {
                            //            reDocList.Add(docItem);
                            //            isAdd = true;
                            //            break;
                            //        }
                            //    }
                            //    if (isAdd == true)
                            //    {
                            //        continue;
                            //    }
                            //}
                            //判断流程用户里面有没有当前用户
                            //if (doc.WorkFlow != null && doc.WorkFlowList != null)
                            //{
                            //    bool isAdd = false;
                            //    foreach (WorkState ws in doc.WorkFlow.WorkStateList)
                            //    {
                            //        foreach (WorkUser wu in ws.WorkUserList)
                            //        {
                            //            if (wu.User == curUser)
                            //            {
                            //                reDocList.Add(docItem);
                            //                isAdd = true;
                            //                break;
                            //            }
                            //        }
                            //        if (isAdd == true)
                            //        {
                            //            break;
                            //        }
                            //    }
                            //        if (isAdd == true)
                            //        {
                            //           continue;
                            //        }
                            //}

                            #endregion

                            string strAttrKeyword = "";

                            #region 如果文档没有模板，就找流程上其他有模板的文档
                            //如果文档没有模板，就找流程上其他有模板的文档
                            if (doc.TempDefn == null && doc.WorkFlow != null)
                            {
                                Doc twd = null;
                                foreach (Doc wd in doc.WorkFlow.DocList)
                                {
                                    if (wd.TempDefn != null)
                                    {
                                        twd = wd;
                                        break;
                                    }
                                }
                                if (twd != null)
                                {
                                    doc = twd;
                                }
                            }

                            #endregion

                            if (doc.TempDefn != null && doc.TempDefn.KeyWord == "CATALOGUING")
                            {
                                strAttrKeyword = "CA_SENDERCODE";
                            }

                            //2.1.1获取文档的发文单位，并获取该发文单位的所有成员
                            if ((data = doc.GetAttrDataByKeyWord(strAttrKeyword)) != null)
                            {
                                string senderCode = data.ToString;
                                AVEVA.CDMS.Server.Group group = dbsource.GetGroupByName(senderCode);
                                if (group == null)
                                {
                                    continue;
                                }

                                //2.2.2. 判断当前用户是否是发文单位的成员，不是发文单位的成员就跳到下一个文档
                                if (group.AllUserList.Contains(curUser))
                                    //2.2.3 如果当前用户是当前文档的发文单位的成员，就把文档添加到文档列表
                                    reDocList.Add(docItem);
                            }

                        }
                        catch { }
                    }
                    docList = reDocList;
                    return true;
                }
                #endregion

                //2.2  如果是收文目录，就遍历文档
                else if (projectType == "收文")
                {
                    AttrData data;
                    foreach (Doc docItem in resultDocList)
                    {
                        try
                        {
                            Doc doc = docItem;
                            bool isShort = false;
                            if (doc.ShortCutDoc != null)
                            {
                                doc = doc.ShortCutDoc;
                                isShort = true;
                            }

                            if (doc.Creater == curUser && isShort == false)
                            {
                                reDocList.Add(docItem);
                                continue;
                            }

                            #region 如果文档没有模板，就找流程上其他有模板的文档
                            //如果文档没有模板，就找流程上其他有模板的文档
                            if (doc.TempDefn == null && doc.WorkFlow != null)
                            {
                                Doc twd = null;
                                foreach (Doc wd in doc.WorkFlow.DocList)
                                {
                                    if (wd.TempDefn != null)
                                    {
                                        twd = wd;
                                        break;
                                    }
                                }
                                if (twd != null)
                                {
                                    doc = twd;
                                }
                            }

                            #endregion

                            string strRecAttrKeyword = "";
                            string strCopyAttrKeyword = "";
                            if (doc.TempDefn.KeyWord == "CATALOGUING")
                            {
                                strRecAttrKeyword = "CA_MAINFEEDERCODE";
                                strCopyAttrKeyword = "CA_COPYCODE";
                            }

                            //2.2.1 获取文档的收文单位和抄送单位
                            List<Server.Group> copyGroupList = new List<Server.Group>();
                            if ((data = doc.GetAttrDataByKeyWord(strRecAttrKeyword)) != null)
                            {
                                string senderCode = data.ToString;
                                AVEVA.CDMS.Server.Group group = dbsource.GetGroupByName(senderCode);
                                if (group == null)
                                {
                                    continue;
                                }

                                // 2.2.2 判断当前用户是否是收文单位或者抄送单位的成员，不是收文单位或者抄送单位的成员就跳到下一个文档
                                if (group.AllUserList.Contains(curUser))
                                    //2.2.3 如果当前用户是当前文档的收文单位的成员，就把文档添加到文档列表
                                    reDocList.Add(docItem);
                                continue;
                            }



                            if ((data = doc.GetAttrDataByKeyWord(strCopyAttrKeyword)) != null)
                            {
                                string senderCode = data.ToString;
                                string[] strArry = senderCode.Split(new char[] { ',' });
                                bool isAdd = false;
                                foreach (string strCopy in strArry)
                                {

                                    AVEVA.CDMS.Server.Group group = dbsource.GetGroupByName(strCopy);
                                    if (group == null)
                                    {
                                        continue;
                                    }

                                    // 2.2.2 判断当前用户是否是抄送单位的成员，不是收文单位或者抄送单位的成员就跳到下一个文档
                                    if (group.AllUserList.Contains(curUser))
                                    {
                                        //2.2.3 如果当前用户是当前文档的抄送单位的成员，就把文档添加到文档列表
                                        reDocList.Add(docItem);
                                        isAdd = true;
                                        break;
                                    }

                                }
                            }
                        }
                        catch { }

                    }
                    docList = reDocList;
                    return true;
                }

                docList = resultDocList;
            }
            catch (Exception ex){
                WebApi.CommonController.WebWriteLog(DateTime.Now.ToString() + ":" + "获取文件列表错误," + ex.Message);

            }
            //docList = docList.Where(d=>(( d.WorkFlow==null && d.Creater==curUser) || d.WorkFlow != null)
            // || (d.ShortCutDoc != null && ((d.ShortCutDoc.WorkFlow == null && d.ShortCutDoc.Creater == curUser) || d.ShortCutDoc.WorkFlow != null))
            // //|| d.WorkFlow.O_WorkFlowStatus != enWorkFlowStatus.Finish  //流出已经走完的文档不显示
            // ).ToList<Doc>();

            return false;
        }

        //预览文件前处理的事件
        public static bool BeforePreviewDoc(string PlugName, Doc doc, ref bool PVRight) {
            if (PlugName != PluginName)
            {
                return false;
            }

            DBSource dbsource = doc.dBSource;
            User curUser = dbsource.LoginUser;
            if (curUser.IsAdmin) {
                return true;
            }
            #region 函数逻辑
            //1.获取文档密级属性

            //2.如果密级是公开的，就返回可以显示

            //3.如果密级是受限的（发文部门和收文部门可以查看），获取收文部门和发文部门的所有用户成员

            //3.1 如果当前用户在收文部门和发文部门的所有用户成员里面，就可以显示，否则不可以显示

            //4.如果密级是商业秘密的（流程里面的人和部门领导可以查看），获取流程里面的所有人员和部门领导

            //4.1 如果当前用户在流程里面的所有人员和部门领导里面的，就可以显示，否则不可以显示 
            #endregion

            #region 1.获取文档密级属性


            string secretgrade = "";
            AttrData data;
            //获取发送方代码
            if ((data = doc.GetAttrDataByKeyWord("CA_SECRETGRADE")) != null)
            {
                secretgrade = data.ToString;
            }
            if (string.IsNullOrEmpty(secretgrade)) {
                return false;
            }
            #endregion

            //2.如果密级是公开的，就返回可以显示(不修改PVRight)
            if (secretgrade == "公开") return true;

            #region 3.如果密级是受限的（发文部门和收文部门可以查看），获取收文部门和发文部门的所有用户成员

            if (secretgrade == "受限") {
                //先把文档预览权限设置为否，再判断是否有文档预览权限
                PVRight = false;

                #region 判断是否是发文部门的成员，如果是，就返回可以预览文件
                //获取发文部门
                string senderCode = "";
                if ((data = doc.GetAttrDataByKeyWord("CA_SENDERCODE")) != null)
                {
                    senderCode = data.ToString;
                }

                if (!string.IsNullOrEmpty(senderCode))
                {
                    //获取发文部门的所有人员
                    AVEVA.CDMS.Server.Group group = dbsource.GetGroupByName(senderCode);
                    if (group != null)
                    {

                        //3.1.1. 判断当前用户是否是发文单位的成员
                        if (group.AllUserList.Contains(curUser))
                        //3.1.2 如果当前用户是当前文档的发文单位的成员，就返回可以预览文档
                        {
                            PVRight = true;
                            return true;
                        }
                    }
                }
                #endregion

                //获取收文部门
                //"CA_MAINFEEDERCODE"
                #region 判断是否是收文部门的成员，如果是，就返回可以预览文件
                //获取发文部门
                string recverCode = "";
                if ((data = doc.GetAttrDataByKeyWord("CA_MAINFEEDERCODE")) != null)
                {
                    recverCode = data.ToString;
                }

                if (!string.IsNullOrEmpty(recverCode))
                {
                    //获取发文部门的所有人员
                    AVEVA.CDMS.Server.Group recGroup = dbsource.GetGroupByName(recverCode);
                    if (recGroup != null)
                    {

                        //3.1.1. 判断当前用户是否是发文单位的成员
                        if (recGroup.AllUserList.Contains(curUser))
                        //3.1.2 如果当前用户是当前文档的发文单位的成员，就返回可以预览文档
                        {
                            PVRight = true;
                            return true;
                        }
                    }
                }
                #endregion

                //如果不在收文和发文部门里面，就返回不可以预览
                PVRight = false;
                return true;
            }

            #endregion

            //4.如果密级是商业秘密的（流程里面的人和部门领导可以查看），获取流程里面的所有人员和部门领导
            if (secretgrade == "商业秘密")
            {
                //先把文档预览权限设置为否，再判断是否有文档预览权限
                PVRight = false;

                #region 判断流程用户里面有没有当前用户
              
                
                if (doc.WorkFlow != null && doc.WorkFlowList != null)
                {
                    bool isAdd = false;
                    foreach (WorkState ws in doc.WorkFlow.WorkStateList)
                    {
                        if (ws == null) continue;

                        foreach (WorkUser wu in ws.WorkUserList)
                        {
                            if (wu.User == curUser)
                            {
                                //如果在流程里面，就返回可以预览
                                PVRight = true;
                                return true;
                            }
                        }
                    }
                }

                #endregion

                #region 再查找文档模板里面定义里面的流程用户，是否有当前用户
                //再查找文档模板里面定义里面的流程用户，是否有当前用户
                //string senderCode = "";
                //编制人 
                if ((data = doc.GetAttrDataByKeyWord("LE_DESIGN")) != null)
                {
                    if (curUser.ToString == data.ToString) PVRight = true;
                }

                //校核人 
                if ((data = doc.GetAttrDataByKeyWord("LE_CHECK")) != null)
                {
                    if (curUser.ToString == data.ToString) PVRight = true;
                }

                //审核人
                if ((data = doc.GetAttrDataByKeyWord("LE_AUDIT")) != null)
                {
                    if (curUser.ToString == data.ToString) PVRight = true;
                }

                //审定人
                if ((data = doc.GetAttrDataByKeyWord("LE_AUDIT2")) != null)
                {
                    if (curUser.ToString == data.ToString) PVRight = true;
                }

                //批准人
                if ((data = doc.GetAttrDataByKeyWord("LE_APPROV")) != null)
                {
                    if (curUser.ToString == data.ToString) PVRight = true;
                }
                if (PVRight == true)
                {
                    return true;
                }
                #endregion

                #region 判断是否发文部门的部门领导，是就返回可以预览文档
                //判断是否发文部门的部门领导，是就返回可以预览文档
                string senderCode = "";
                if ((data = doc.GetAttrDataByKeyWord("CA_SENDERCODE")) != null)
                {
                    senderCode = data.ToString;
                }

                if (!string.IsNullOrEmpty(senderCode))
                {
                    //获取发文部门的部门领导
                    AVEVA.CDMS.Server.Group group = dbsource.GetGroupByName(senderCode);
                    if (group != null)
                    {
                        AVEVA.CDMS.Server.Group ldGroup = group.AllGroupList.Find(g => g.Description == "部长");
                        if (ldGroup != null)
                        {
                            if (ldGroup.AllUserList.Contains(curUser))
                            //3.1.2 如果当前用户是当前文档的发文单位的成员，就返回可以预览文档
                            {
                                PVRight = true;
                                return true;
                            }
                        }
                    }
                }
                #endregion

                #region 判断是否收文部门的部门领导，是就返回可以预览文档
                //判断是否收文部门的部门领导，是就返回可以预览文档
                string recverCode = "";
                if ((data = doc.GetAttrDataByKeyWord("CA_MAINFEEDERCODE")) != null)
                {
                    recverCode = data.ToString;
                }

                if (!string.IsNullOrEmpty(recverCode))
                {
                    //获取发文部门的部门领导
                    AVEVA.CDMS.Server.Group group = dbsource.GetGroupByName(recverCode);
                    if (group != null)
                    {
                        AVEVA.CDMS.Server.Group ldGroup = group.AllGroupList.Find(g => g.Description == "部长");
                        if (ldGroup != null)
                        {
                            if (ldGroup.AllUserList.Contains(curUser))
                            //3.1.2 如果当前用户是当前文档的发文单位的成员，就返回可以预览文档
                            {
                                PVRight = true;
                                return true;
                            }
                        }
                    }
                }
                #endregion
                return true;
            }
              
            return false;
        }

        //下载文件前处理的事件
        public static bool BeforeDownloadFile(string PlugName, Doc doc, ref bool DLRight)
        {
            if (PlugName != PluginName)
            {
                return false;
            }
            DLRight = false;
            User curUser = doc.dBSource.LoginUser;
            if (curUser.IsAdmin) {
                DLRight = true;
                return true;
            }

            return false;
        }

        public static bool SearchDocList(ref List<Doc> docList, string filter) {


            return false;
        }

        //显示文件夹图标的事件，，用于按条件获取文件夹图标
        public static bool AfterGetProjectIconFun(Project project,ref string iconClass)//ref List<Doc> docList)
        {
            //return true;
            User curUser = project.dBSource.LoginUser;

            List<Project> projectList = project.AllProjectList; 
            projectList.Add(project);

            foreach (Project proj in projectList)
            {
                try { 
                if (proj.TempDefn!=null && proj.TempDefn.KeyWord == "COM_COMTYPE")
                {
     

                    foreach (Doc d in proj.DocList)
                    {
                        try
                        {

                                //排除文档实体发文目录，且流程状态是收文状态的文档
                                if ((d.ShortCutDoc == null && 
                                         (d.WorkFlow != null && d.WorkFlow.CuWorkState.Code!= "RECUNIT" &&
                                              d.WorkFlow.O_WorkFlowStatus != enWorkFlowStatus.Finish
                                              ))
                                     || (d.ShortCutDoc != null &&
                                       (d.ShortCutDoc.WorkFlow != null && d.ShortCutDoc.WorkFlow.O_WorkFlowStatus != enWorkFlowStatus.Finish
                                           ))
                                    )
                            {
                                List<WorkUser> wus = d.ShortCutDoc == null ?
                                    d.WorkFlow.CuWorkState.WorkUserList :
                                    d.ShortCutDoc.WorkFlow.CuWorkState.WorkUserList;

                                foreach (WorkUser wu in wus)
                                {
                                    if (wu.User != null && wu.User == curUser)
                                    {
                                        //resultDocList.Add(d);
                                        iconClass = "final2";
                                        return true;
                                    }
                                }
                            }

                        }
                        catch { }
                     
                    }
                    }
                }
                catch { }
            }
            return true;
        }

        public static ExReJObject BeforeWF(WorkFlow wf, WorkStateBranch wsb)
        {
            ExReJObject reJo = new ExReJObject();
            try
            {
                #region 项目立项流程
                if (wf.DefWorkFlow.O_Code == "CREATEPROJECT")
                {
                    if (wsb.defStateBrach.O_Description == "同意" || wsb.defStateBrach.O_Description == "结束")
                    {
                        //如果是填写好表单后再来提交下一流程，就流转到下一流程状态
                        if (!string.IsNullOrEmpty(wf.O_suser3) && wf.O_suser3 == "pass")
                        {
                            //当reJo的成功状态返回为真时，继续流转到下一流程状态的操作
                            reJo.success = true;
                            return reJo;
                        }

                        DBSource dbsource = wf.dBSource;

                        Doc m_Doc = wf.doc;
                        if (m_Doc == null)
                        {
                            reJo.success = true;
                            return reJo;
                        }

                        string projectCode = m_Doc.GetAttrDataByKeyWord("PROCODE").ToString;

                        string projectDescCN = m_Doc.GetAttrDataByKeyWord("PRONAME").ToString;

                        //string sourceUnite= m_Doc.GetAttrDataByKeyWord("PRO_COMPANY").ToString;
                        string sourceUnite = "";
                        if (string.IsNullOrEmpty(sourceUnite)) sourceUnite = "";

                        JArray jaAttr = new JArray(
                            new JObject(
                                new JProperty("name", "projectCode"), new JProperty("value", projectCode)
                                ),
                            new JObject(
                                new JProperty("name", "projectDescCN"), new JProperty("value", projectDescCN)
                                ),
                            new JObject(
                                new JProperty("name", "sourceUnite"), new JProperty("value", sourceUnite)
                                )
                                );

                        ExReJObject reJo2 = CreateProject.CreateRootProjectX(dbsource, jaAttr);
                    }

                   

                }
                #endregion

                #region 发文流程
                //发文流程
                if (wf.DefWorkFlow.O_Code == "COMMUNICATIONWORKFLOW")
                {

                    #region 批准人盖章
                    if (wf.CuWorkState.Code == "APPROV" && wsb.defStateBrach.O_Code == "TOSECRE" && wf.dBSource.LoginUser.O_userno == wf.CuWorkState.CuWorkUser.O_userno)
                    {
                        //如果是填写好表单后再来提交下一流程，就流转到下一流程状态
                        if (!string.IsNullOrEmpty(wf.O_suser3) && wf.O_suser3 == "approvpass")
                        {
                            //当reJo的成功状态返回为真时，继续流转到下一流程状态的操作
                            reJo.success = true;
                            return reJo;
                        }



                        //弹出重新设置发文编号窗口
                        reJo.data = new JArray(new JObject(
                            new JProperty("state", "RunFunc"),
                                new JProperty("plugins", "HXEPC_Plugins"),
                                new JProperty("DefWorkFlow", wf.DefWorkFlow.O_Code),
                                new JProperty("CuWorkState", wf.CuWorkState.Code),
                                 new JProperty("FuncName", "documenteSeal"),
                                new JProperty("DocKeyword", wf.doc.KeyWord)
                                 ));
                        //当reJo的成功状态返回为假时，中断流转到下一流程状态的操作
                        reJo.success = false;
                        return reJo;
                    }
                    #endregion

                    #region 文控取号发出
                    //发文流程，文秘发出后，将发出的公文放置在共有发文目录下
                    if (wf.CuWorkState.Code == "SECRETARILMAN" && wsb.defStateBrach.O_Code == "TORECUNIT" && wf.dBSource.LoginUser.O_userno == wf.CuWorkState.CuWorkUser.O_userno)
                    {
                        //wf.O_suser3 = "pass";
                        //wf.Modify();

                        //如果是填写好表单后再来提交下一流程，就流转到下一流程状态
                        if (!string.IsNullOrEmpty(wf.O_suser3) && wf.O_suser3 == "pass")
                        {

                            //复制快捷方式到收文目录
                            Project proj = wf.DocList[0].Project;

                            Project rsProject = CommonFunction.getParentProjectByTempDefn(proj, "PRO_COMMUNICATION");

                            Project recProj = CommonFunction.GetProjectByDesc(rsProject, "收文");

                            Project recCommProj = recProj.GetProjectByName(proj.Code);

                            foreach (Doc doc in wf.DocList)
                            {
                                recCommProj.NewDoc(doc);
                            }

                            //当reJo的成功状态返回为真时，继续流转到下一流程状态的操作
                            reJo.success = true;
                            return reJo;
                        }

                        string docType = "";
                        string FileCode = "";
                        Doc wfDoc = null;// new Doc();
                        foreach (Doc doc in wf.DocList)
                        {
                            if (doc.TempDefn.KeyWord == "CATALOGUING")
                            {
                                docType = "信函";
                                string strDocType = "";
                                AttrData data;
                                //文文档模板
                                if ((data = doc.GetAttrDataByKeyWord("CA_ATTRTEMP")) != null)
                                {
                                    if (data.ToString == "LETTERFILE")
                                    {
                                        strDocType = "LET";
                                    }
                                    else if (data.ToString == "FILETRANSMIT")
                                    {
                                        strDocType = "TRA";
                                    }
                                    else if (data.ToString == "MEETINGSUMMARY")
                                    {
                                        strDocType = "MOM";
                                    }
                                }
                                wfDoc = doc;
                                string senderCode = "", recerCode = "";

                                if ((data = doc.GetAttrDataByKeyWord("CA_SENDERCODE")) != null)
                                {
                                    senderCode = data.ToString;
                                    FileCode = FileCode + data.ToString + "-";

                                }
                                if ((data = doc.GetAttrDataByKeyWord("CA_MAINFEEDERCODE")) != null)
                                {
                                    recerCode = data.ToString;
                                    FileCode = FileCode + data.ToString + "-";
                                }
                                string docnum = Document.getDocNumber(wf.dBSource, "", strDocType, senderCode, recerCode);
                                FileCode = FileCode + strDocType + "-" + docnum;
                                break;
                            }
                        }

                        //if (wfDoc == null) {
                        //    wfDoc = wf.doc;
                        //}

                        //wf.O_suser3 = "pass";
                        //wf.Modify();

                        //Project proj = wf.doc.Project;

                        //if (proj.Description == "信函") {

                        //    FileCode=""
                        //}
                        //弹出重新设置发文编号窗口
                        reJo.data = new JArray(new JObject(
                            new JProperty("state", "RunFunc"),
                                new JProperty("plugins", "HXEPC_Plugins"),
                                new JProperty("DefWorkFlow", wf.DefWorkFlow.O_Code),
                                new JProperty("CuWorkState", wf.CuWorkState.Code),
                                // new JProperty("FuncName", "letterCNFillInfo"),
                                new JProperty("FuncName", "resetFileCode"),
                                new JProperty("DocKeyword", wf.doc.KeyWord),
                                new JProperty("FileCode", FileCode)
                                ));
                        //当reJo的成功状态返回为假时，中断流转到下一流程状态的操作
                        reJo.success = false;
                        return reJo;
                    }
                    #endregion

                    #region 收文部门归档
                    if (wf.CuWorkState.Code == "RECUNIT" && wsb.defStateBrach.O_Code == "TOEND" && wf.dBSource.LoginUser.O_userno == wf.CuWorkState.CuWorkUser.O_userno)
                    {
                        try
                        {
                            Project proj = wf.doc.Project;
                            Doc doc = null;
                            string docType = "", senderKeyword = "", recerKeyword = "";
                            foreach (Doc docItem in wf.DocList)
                            {
                                if (docItem.TempDefn.KeyWord == "CATALOGUING")
                                {
                                    doc = docItem;
                                    AttrData data;
                                    //文文档模板
                                    if ((data = doc.GetAttrDataByKeyWord("CA_ATTRTEMP")) != null)
                                    {
                                        if (data.ToString == "LETTERFILE")
                                        {
                                            docType = "信函";
                                        }
                                        else if (data.ToString == "FILETRANSMIT")
                                        {
                                            docType = "文件传递单";
                                        }
                                        else if (data.ToString == "MEETINGSUMMARY")
                                        {
                                            docType = "会议纪要";
                                        }
                                    }
                                    senderKeyword = "CA_SENDERCODE";
                                    recerKeyword = "CA_MAINFEEDERCODE";
                                    break;
                                }
                            }

                            if (doc == null)
                            {
                                reJo.success = true;
                                return reJo;
                            }

                            //发文单位代码
                            string senderCode = doc.GetAttrDataByKeyWord(senderKeyword).ToString;
                            //收文单位代码
                            string recerCode = doc.GetAttrDataByKeyWord(recerKeyword).ToString;

                            //运营管理类文件目录
                            Project rootProj = CommonFunction.getParentProjectByTempDefn(proj, "OPERATEADMIN");
                            //存档管理目录
                            Project storaProj = CommonFunction.GetProjectByDesc(rootProj, "存档管理");
                            //通信类目录
                            Project commProj = CommonFunction.GetProjectByDesc(storaProj, "通信类");

                            //发文单位目录
                            Project senderProj = commProj.GetProjectByName(senderCode);
                            Project sendProj = null;
                            //发文单位目录下的发文下的函件类型目录
                            if (senderProj == null)
                            {
                                reJo.msg = "归档失败，发文目录不存在！";
                                reJo.success = false;
                                return reJo;

                            }
                            try
                            {
                                sendProj = senderProj.GetProjectByName("发文").GetProjectByName(docType);
                                if (sendProj == null)
                                {
                                    reJo.msg = "归档失败，发文目录不存在！";
                                    reJo.success = false;
                                    return reJo;
                                }
                            }
                            catch
                            {
                                reJo.msg = "归档失败，发文目录不存在！";
                                reJo.success = false;
                                return reJo;
                            }


                            //收文单位目录
                            Project recerProj = commProj.GetProjectByName(recerCode);
                            //收文单位目录下的发文下的函件类型目录
                            if (recerProj == null)
                            {
                                reJo.msg = "归档失败，收文目录不存在！";
                                reJo.success = false;
                                return reJo;
                            }

                            try
                            {
                                Project recProj = recerProj.GetProjectByName("收文").GetProjectByName(docType);
                                if (recProj == null)
                                {
                                    reJo.msg = "归档失败，收文目录不存在！！";
                                    reJo.success = false;
                                    return reJo;
                                }
                                foreach (Doc cpdoc in wf.DocList)
                                {

                                    //sendProj.NewDoc(cpdoc);
                                    //recProj.NewDoc(cpdoc);
                                    #region 判断是否是信函模板
                                    bool bConventToPdf = true;
                                    if (cpdoc.TempDefn.KeyWord != "CATALOGUING")
                                    {
                                        bConventToPdf = false;
                                    }
                                    AttrData data;

                                    if ((data = doc.GetAttrDataByKeyWord("CA_ATTRTEMP")) == null)
                                    {
                                        bConventToPdf = false;
                                    }


                                    string strData = data.ToString;
                                    if (!(strData == "LETTERFILE" || strData == "FILETRANSMIT" ||
                                       strData == "MEETINGSUMMARY" || strData == "DOCUMENTFILE"))
                                    {
                                        bConventToPdf = false;
                                    }


                                    #endregion
                                    #region 转换PDF到文件所在目录

                                    if (bConventToPdf)
                                    {
                                        try
                                        {
                                            //    string filecode = cpdoc.Code;
                                            //string filedesc = cpdoc.Description;

                                            string sourceFileName = cpdoc.FullPathFile;
                                            string targetFileName = sourceFileName.Substring(0, sourceFileName.LastIndexOf(".") + 1) + "pdf";

                                            CDMSPdf.ConvertToPdf(cpdoc.FullPathFile, targetFileName);
                                            cpdoc.O_filename = cpdoc.O_filename.Substring(0, cpdoc.O_filename.LastIndexOf(".") + 1) + "pdf";
                                            //cpdoc.SetFileName(cpdoc.O_filename.Substring(0, cpdoc.O_filename.LastIndexOf(".") + 1) + "pdf");
                                            cpdoc.Modify();
                                        }
                                        catch { }
                                    }
                                    #endregion

                                    sendProj.NewDoc(cpdoc);
                                    recProj.NewDoc(cpdoc);


                                    ////存档管理下的发文目录创建文档
                                    //Doc sendDocItem = sendProj.NewDoc(filecode + ".pdf", filecode, filedesc, cpdoc.TempDefn);

                                    ////存档管理下的发文目录创建文档
                                    ////Doc recDocItem = recProj.NewDoc(filecode + ".pdf", filecode, filedesc, cpdoc.TempDefn);
                                    //recProj.NewDoc(sendDocItem);

                                    //sendDocItem.WorkFlow = cpdoc.WorkFlow;
                                    //sendDocItem.Modify();
                                    //try
                                    //{
                                    //    CDMSPdf.ConvertToPdf(cpdoc.FullPathFile, sendDocItem.FullPathFile);

                                    //}
                                    //catch { }
                                    //List<AttrData> dataList = cpdoc.AttrDataList;//.GetAttrDataList();

                                    //foreach (AttrData dataItem in dataList)
                                    //{
                                    //    AttrData data;

                                    //    if ((data = sendDocItem.AttrDataList.Find(ad => ad.DefnID == dataItem.DefnID)) != null)
                                    //    {
                                    //        data.SetCodeDesc(dataItem.ToString);
                                    //    }
                                    //}
                                    //sendDocItem.AttrDataList.SaveData();
                                    //再修改源文件时间
                                    //System.IO.File.SetLastWriteTime(sourceFileName, srcUpdateTime);


                                }

                            }
                            catch
                            {
                                reJo.msg = "归档失败，收文目录不存在！！！";
                                reJo.success = false;
                                return reJo;
                            }


                            reJo.success = true;
                            return reJo;

                        }
                        catch
                        {
                            reJo.success = true;
                            return reJo;
                        }
                    }
                    #endregion

                    //收文部门分发办理
                    if (wf.CuWorkState.Code == "RECUNIT" && wsb.defStateBrach.O_Code == "TOCONTROL1" && wf.dBSource.LoginUser.O_userno == wf.CuWorkState.CuWorkUser.O_userno)
                    {
                        try
                        {
                            string gpKeyword = "";
                            Server.Group gp = CommonFunction.GetUserRootOrgGroup(wf.dBSource.LoginUser);
                            if (gp != null) {
                                gpKeyword = gp.KeyWord;
                            }

                            //弹出重新设置发文编号窗口DistriProcess
                            reJo.data = new JArray(new JObject(
                                new JProperty("state", "RunFunc"),
                                    new JProperty("plugins", "HXEPC_Plugins"),
                                    new JProperty("DefWorkFlow", wf.DefWorkFlow.O_Code),
                                    new JProperty("CuWorkState", wf.CuWorkState.Code),
                                    new JProperty("FuncName", "distriProcess"),
                                    new JProperty("DocKeyword", wf.doc.KeyWord),
                                    new JProperty("ProjectKeyword", wf.doc.Project.KeyWord),
                                    new JProperty("GroupType","org"),
                                    new JProperty("GroupKeyword", gpKeyword)
                                    ));
                            //当reJo的成功状态返回为假时，中断流转到下一流程状态的操作
                            reJo.success = false;
                            return reJo;
                        }
                        catch {
                            reJo.msg = "分发办理失败！";
                            reJo.success = false;
                            return reJo;
                        }
                    }
                }
                #endregion

                #region 收文流程
                if (wf.DefWorkFlow.O_Code == "RECEIVED")
                {
                    AttrData attrDataByKeyWord;
                    //if ((wf.CuWorkState.Code == "DESIGNMAN") && ((wf.DocList != null) && (wf.DocList.Count > 0)))
                    //{
                    //    attrDataByKeyWord = wf.DocList[0].GetAttrDataByKeyWord("DRAFTMAN");
                    //    if (attrDataByKeyWord != null)
                    //    {
                    //        attrDataByKeyWord.SetCodeDesc(wf.CuWorkState.CuWorkUser.User.ToString);
                    //        wf.DocList[0].AttrDataList.SaveData();
                    //    }
                    //}
                    if (wsb.defStateBrach.O_Description == "回复")
                    {
                        //如果是填写好表单后再来提交下一流程，就流转到下一流程状态
                        if (!string.IsNullOrEmpty(wf.O_suser3) && wf.O_suser3 == "pass")
                        {
                            //当reJo的成功状态返回为真时，继续流转到下一流程状态的操作
                            reJo.success = true;
                            return reJo;
                        }

                        //弹出填写收发文单位，发文编号窗口
                        reJo.data = new JArray(new JObject(
                            new JProperty("state", "RunFunc"),
                                new JProperty("plugins", "HXEPC_Plugins"),
                                new JProperty("DefWorkFlow", wf.DefWorkFlow.O_Code),
                                new JProperty("CuWorkState", wf.CuWorkState.Code),
                                new JProperty("FuncName", "replyLetterCN"),
                                new JProperty("DocKeyword", wf.doc.KeyWord)
                                ));
                        //当reJo的成功状态返回为假时，中断流转到下一流程状态的操作
                        reJo.success = false;
                        return reJo;
                        
                    }
                } 
                #endregion
            }
            catch
            { //throw; 
            }
            //当reJo的成功状态返回为真时，继续流转到下一流程状态的操作
            reJo.success = true;
            return reJo;
        }

        public static bool RevokeWorkFlow(string PlugName, WorkFlow wf) {
            if (PlugName != PluginName)
            {
                return false;
            }

            return true;
        }

        public static bool BeforeSelectUser(string PlugName, WorkStateBranch wsBranch, ref string tabType, ref string tabPara)
        {
            if (PlugName != PluginName)
            {
                return false;
            }
            DBSource dbsource = wsBranch.dBSource;
            User curUser = dbsource.LoginUser;

            string groupCode = "";
            string groupKeyword = "";
            string groupType = "";

            //获取组织机构用户组
            foreach (AVEVA.CDMS.Server.Group groupOrg in dbsource.AllGroupList)
            {
                if ((groupOrg.ParentGroup == null) && (groupOrg.O_grouptype == enGroupType.Organization))
                {
                    if (groupOrg.AllUserList.Contains(curUser))
                    {

                        groupCode = groupOrg.Code;
                        groupKeyword = groupOrg.KeyWord;
                        groupType = "org";
                           break;
                    }
                }
            }

            tabType = groupType;
            tabPara = groupKeyword;

            return true;
        }

        //public static JObject GetEditCompanyDefault(string sid, string ProjectKeyword)
        //{
        //    return Company.GetEditCompanyDefault(sid, ProjectKeyword);

        //}

    }
}

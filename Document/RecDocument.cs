using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using AVEVA.CDMS.Server;
using AVEVA.CDMS.Common;
using AVEVA.CDMS.WebApi;

namespace AVEVA.CDMS.HXEPC_Plugins
{
    public class RecDocument
    {
        //线程锁 
        internal static Mutex muxConsole = new Mutex();

        /// <summary>
        /// 处理收文表单
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="DocKeyword"></param>
        /// <returns></returns>
        public static JObject ReceiveDocument(string sid, string DocKeyword, string docAttrJson)
        {
            ExReJObject reJo = new ExReJObject();
            try
            {
                User curUser = DBSourceController.GetCurrentUser(sid);
                if (curUser == null)
                {
                    reJo.msg = "登录验证失败！请尝试重新登录！";
                    return reJo.Value;
                }

                DBSource dbsource = curUser.dBSource;
                if (dbsource == null)
                {
                    reJo.msg = "登录验证失败！请尝试重新登录！";
                    return reJo.Value;
                }

                Doc m_Doc = dbsource.GetDocByKeyWord(DocKeyword);

                if (m_Doc == null)
                {
                    reJo.msg = "参数错误！文档不存在！";
                    return reJo.Value;
                }

                #region 获取收文处理表单的属性

                string strProjectCode = "", strProjectDesc="", strCommUnit = "", strRecDate="",
                    strRecCode ="", strRecNumber="", strPages="",
                    strSendCode="", strNeedReply="", strReplyDate="",
                    strUrgency="", strRemark="",
                   strFileCode="", strTitle = "";

                JArray jaAttr = (JArray)JsonConvert.DeserializeObject(docAttrJson);

                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    //获取项目号
                    if (strName == "ProjectCode") strProjectCode = strValue;

                    //获取项目名称
                    if (strName == "ProjectDesc") strProjectDesc = strValue;

                    //获取来文单位
                    else if (strName == "CommUnit") strCommUnit = strValue;

                    //获取来文单位
                    else if (strName == "RecDate") strRecDate = strValue;

                    //获取收文编码
                    else if (strName == "RecCode") strRecCode = strValue;

                    //获取收文编码
                    else if (strName == "RecNumber") strRecNumber = strValue;

                    //获取文件编码
                    else if (strName == "FileCode") strFileCode = strValue;

                  

                    //获取文件题名
                    else if (strName == "Title") strTitle = strValue;

                    //获取页数
                    else if (strName == "Pages") strPages = strValue;

                    //获取我方发文编码
                    else if (strName == "SendCode") strSendCode = strValue;

                    //获取是否要求回文
                    else if (strName == "NeedReply") strNeedReply = strValue;

                    //获取回文期限
                    else if (strName == "ReplyDate") strReplyDate = strValue;

                    //获取紧急程度
                    else if (strName == "Urgency") strUrgency = strValue;

                    //获取备注
                    else if (strName == "Remark") strRemark = strValue;
                }
                #endregion

                #region 根据收文单模板，生成收文单文档

                //获取立项单文档所在的目录
                //Project m_Project = m_NewProject;

                List<TempDefn> docTempDefnByCode = dbsource.GetTempDefnByCode("RECEIPT");
                TempDefn docTempDefn = (docTempDefnByCode != null && docTempDefnByCode.Count > 0 )? docTempDefnByCode[0] : null;
                if (docTempDefn == null)
                {
                    reJo.msg = "没有与其相关的模板管理，创建无法正常完成";
                    return reJo.Value;
                }

                IEnumerable<string> source = from docx in m_Doc.Project.DocList select docx.Code;
                string filename = strRecCode + " 收文单";
                if (source.Contains<string>(filename))
                {

                    for (int i = 1; i < 0x3e8; i++)
                    {
                        filename = strRecCode + " 收文单"+i.ToString();
                        if (!source.Contains<string>(filename))
                        {
                            break;
                        }
                    }
                }

                //文档名称
                //Doc docItem = m_Doc.Project.NewDoc(filename + ".docx", filename, "", docTempDefn);
                Doc docItem = m_Doc.Project.NewDoc(filename + ".docx", filename, "");
                if (docItem == null)
                {
                    reJo.msg = "新建信函出错！";
                    return reJo.Value;
                }

                #endregion

                #region 设置收文文档附加属性
                m_Doc.TempDefn = docTempDefn;

                AttrData data;
                //项目名称
                if ((data = m_Doc.GetAttrDataByKeyWord("RE_PRONAME")) != null)
                {
                    data.SetCodeDesc(strProjectDesc);
                }
                //项目代码
                if ((data = m_Doc.GetAttrDataByKeyWord("RE_PROCODE")) != null)
                {
                    data.SetCodeDesc(strProjectCode);
                }
                //来文单位
                if ((data = m_Doc.GetAttrDataByKeyWord("RE_SENDUNIT")) != null)
                {
                    data.SetCodeDesc(strCommUnit);
                }
                //收文日期
                if ((data = m_Doc.GetAttrDataByKeyWord("RE_RECDATE")) != null)
                {
                    data.SetCodeDesc(strRecDate);
                }
                //收文编码
                if ((data = m_Doc.GetAttrDataByKeyWord("RE_RECCODE")) != null)
                {
                    data.SetCodeDesc(strRecCode);
                }
                //收文编号
                if ((data = m_Doc.GetAttrDataByKeyWord("RE_RECNUMBER")) != null)
                {
                    data.SetCodeDesc(strRecNumber);
                }
                //文件编码
                if ((data = m_Doc.GetAttrDataByKeyWord("RE_FILECODE")) != null)
                {
                    data.SetCodeDesc(strFileCode);
                }
                //文件题名
                if ((data = m_Doc.GetAttrDataByKeyWord("RE_FILETITLE")) != null)
                {
                    data.SetCodeDesc(strTitle);
                }
                //页数
                if ((data = m_Doc.GetAttrDataByKeyWord("RE_PAGE")) != null)
                {
                    data.SetCodeDesc(strPages);
                }
                //发文编码
                if ((data = m_Doc.GetAttrDataByKeyWord("RE_SENDCODE")) != null)
                {
                    data.SetCodeDesc(strSendCode);
                }

                //是否回文
                if ((data = m_Doc.GetAttrDataByKeyWord("RE_IFREPLY")) != null)
                {
                    data.SetCodeDesc(strNeedReply);
                }
                //回文日期
                if ((data = m_Doc.GetAttrDataByKeyWord("RE_REPLYDATE")) != null)
                {
                    data.SetCodeDesc(strReplyDate);
                }
                //紧急程度
                if ((data = m_Doc.GetAttrDataByKeyWord("RE_URGENTDEGREE")) != null)
                {
                    data.SetCodeDesc(strUrgency);
                }

                //备注
                if ((data = m_Doc.GetAttrDataByKeyWord("RE_NOTE")) != null)
                {
                    data.SetCodeDesc(strRemark);
                }
                //著录人
                if ((data = m_Doc.GetAttrDataByKeyWord("RE_DESIGN")) != null)
                {
                    data.SetCodeDesc(curUser.ToString);
                }


                ////保存项目属性，存进数据库
                m_Doc.AttrDataList.SaveData();
                //m_Doc.Modify();
                #endregion

                //修改文档编码
                m_Doc.O_itemname = strRecCode; //strRecNumber;
                //m_Doc.O_itemdesc = strTitle;
                m_Doc.Modify();

                #region 录入数据进入word表单
                Hashtable htUserKeyWord = new Hashtable();

                //格式化日期
                DateTime RecDate = Convert.ToDateTime(strRecDate);
                DateTime ReplyDate = Convert.ToDateTime(strReplyDate);
                string recDate = RecDate.ToShortDateString().ToString().Replace("-", ".").Replace("/", ".");
                string replyDate = ReplyDate.ToShortDateString().ToString().Replace("-", ".").Replace("/", ".");

                htUserKeyWord.Add("PRONAME", strProjectDesc);//项目名称
                htUserKeyWord.Add("PROCODE", strProjectCode);//项目代码
                htUserKeyWord.Add("SENDUNIT", strCommUnit);//来文单位
                htUserKeyWord.Add("RECDATE", recDate);//收文日期
                htUserKeyWord.Add("RECCODE", strRecCode);//收文编码
                htUserKeyWord.Add("RECNUMBER", strRecNumber);//收文编号
                htUserKeyWord.Add("FILECODE", strFileCode);//文件编码
                htUserKeyWord.Add("FILETITLE", strTitle);//文件题名
                htUserKeyWord.Add("PAGE", strPages);//页数
                htUserKeyWord.Add("SENDCODE", strSendCode);//发文编码
                htUserKeyWord.Add("IFREPLY", strNeedReply);//是否回文
                htUserKeyWord.Add("REPLYDATE", replyDate);//回文日期
                htUserKeyWord.Add("URGENTDEGREE", strUrgency);//紧急程度
                htUserKeyWord.Add("NOTE", strRemark);//备注
                htUserKeyWord.Add("DESIGN", curUser.Description);//著录人


                string workingPath = dbsource.LoginUser.WorkingPath;


                try
                {
                    //上传下载文档
                    string exchangfilename = "收文单模板";

                    //获取网站路径
                    string sPath = System.Web.HttpContext.Current.Server.MapPath("/ISO/HXEPC/");

                    //获取模板文件路径
                    string modelFileName = sPath + exchangfilename + ".docx";

                    //获取即将生成的联系单文件路径
                    string locFileName = docItem.FullPathFile;

                    FileInfo info = new FileInfo(locFileName);

                    if (System.IO.File.Exists(modelFileName))
                    {
                        //如果存储子目录不存在，就创建目录
                        if (!Directory.Exists(info.Directory.FullName))
                        {
                            Directory.CreateDirectory(info.Directory.FullName);
                        }

                        //复制模板文件到存储目录，并覆盖同名文件
                        System.IO.File.Copy(modelFileName, locFileName, true);


                        //线程锁 
                        muxConsole.WaitOne();
                        try
                        {
                            //把参数直接写进office
                            CDMSWebOffice office = new CDMSWebOffice
                            {
                                CloseApp = true,
                                VisibleApp = false
                            };
                            office.Release(true);
                            office.WriteDataToDocument(docItem, locFileName, htUserKeyWord, htUserKeyWord);
                        }
                        catch { }
                        finally
                        {

                            //解锁
                            muxConsole.ReleaseMutex();
                        }
                    }


                    int length = (int)info.Length;
                    docItem.O_size = new int?(length);
                    docItem.Modify();

                }
                catch { }
                #endregion

                ////启动流程
                WorkFlow flow = dbsource.NewWorkFlow(new List<Doc> { docItem, m_Doc }, "RECEIVED");
                if (flow == null)
                {
                    //AssistFun.PopUpPrompt("自动启动流程失败!请手动启动");
                    reJo.msg = "自动启动流程失败!请手动启动";
                    return reJo.Value;
                }
                else
                {
                    if ((flow != null) && (flow.CuWorkState != null))
                    {
                        //if (((flow.CuWorkState == null) || (flow.CuWorkState.workStateBranchList == null)) || (flow.CuWorkState.workStateBranchList.Count <= 0))
                        //{
                        //    //MessageBox.Show("新建流程不存在下一状态,提交失败!");
                        //    //doc.dBSource.ProgramRun = false;
                        //    flow.Delete();
                        //    reJo.msg = "新建流程不存在下一状态,提交失败!";
                        //    return reJo.Value;
                        //    //return;
                        //}
                        //WorkStateBranch branch = flow.CuWorkState.workStateBranchList[0];
                        //if (branch == null)
                        //{
                        //    reJo.msg = "获取流程分支失败!";
                        //    return reJo.Value;
                        //}

                        Project rootProj = CommonFunction.getParentProjectByTempDefn(m_Doc.Project, "HXNY_DOCUMENTSYSTEM");
                        if (rootProj == null)
                        {
                            reJo.msg = "获取项目根目录失败!";
                            return reJo.Value;
                        }

                        string UserList = "";
                        AttrData secData;
                        //rootProj.GetAttrDataByKeyWord("SECRETARILMAN");
                        if ((secData = rootProj.GetAttrDataByKeyWord("SECRETARILMAN")) != null)
                        {
                            UserList = secData.ToString;
                        }

                        if (string.IsNullOrEmpty(UserList))
                        {
                            reJo.msg = "项目文控未设置！";
                            return reJo.Value;
                        }

                        //ExReJObject wfReJo = WebWorkFlowEvent.GotoNextStateAndSelectUser(flow.CuWorkState.workStateBranchList[0],userlist);
                        ////if (!WorkFlowEvent.GotoNextStateAndSelectUser(flow.CuWorkState.workStateBranchList[0]))
                        //if (!wfReJo.success)
                        //{
                        //    //doc.dBSource.ProgramRun = false;
                        //    flow.Delete();
                        //    flow.Delete();
                        //    reJo.msg = "自动启动收文流程失败，请手动启动流程!";
                        //    return reJo.Value;
                        //}

                        ////刷新数据源
                        //DBSourceController.RefreshDBSource(sid);

                        #region 获取下一状态用户
                        string[] userArray = (string.IsNullOrEmpty(UserList) ? "" : UserList).Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                        Server.Group group = new Server.Group();
                        //List<User> m_UserList = new List<User>();
                        //启动工作流程
                        //反转列表
                        //m_UserList.Reverse();
                        foreach (string strObj in userArray)
                        {
                            string strUser = strObj.IndexOf("__") >= 0 ? strObj.Substring(0, strObj.IndexOf("__")) : strObj;
                            object obj = dbsource.GetUserByName(strUser);
                            
                            if (obj is User)
                            {
                                //m_UserList.Add((User)obj);
                                group.AddUser((User)obj);
                            }
                        }
                        if (group.UserList.Count <= 0)
                        {
                            reJo.msg = "获取下一流程状态用户错误，自动启动流程失败！请手动启动流程";
                            return reJo.Value;
                        }
                        #endregion

                        //获取下一状态

                        WorkState ws = new WorkState();

                            DefWorkState dws = flow.DefWorkFlow.DefWorkStateList.Find(s => s.KeyWord == "SECRETARILMAN");// CHECK");
                            ws.DefWorkState = dws;


                        ////启动流程
                        WorkStateBranch branch = flow.CuWorkState.workStateBranchList[0];
                        branch.NextStateAddGroup(group);

                        ExReJObject GotoNextReJo = WebWorkFlowEvent.GotoNextStateAndSelectUser(flow.CuWorkState.workStateBranchList[0]);

                        if (!GotoNextReJo.success)
                        {
                            //  doc.dBSource.ProgramRun = false;
                            flow.Delete();
                            flow.Delete();

                            reJo.msg = "自动启动流程失败！请手动启动流程";
                            return reJo.Value;
                        }

                        DBSourceController.RefreshDBSource(sid);

                        //return GotoNextReJo.Value;

                        reJo.data = new JArray(new JObject(new JProperty("ProjectKeyword", m_Doc.Project.KeyWord),
                            new JProperty("DocKeyword", m_Doc.KeyWord)));
                        reJo.success = true;
                        return reJo.Value;

                    }


                }

                reJo.success = true;
                return reJo.Value;
            }
            catch (Exception e)
            {
                reJo.msg = e.Message;
                CommonController.WebWriteLog(reJo.msg);
            }
            return reJo.Value;
        }


        /// <summary>
        /// 获取创建收文单表单的默认配置
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="DocKeyword"></param>
        /// <returns></returns>
        public static JObject GetReceiveDocumentDefault(string sid, string DocKeyword)
        {
            ExReJObject reJo = new ExReJObject();
            try
            {
                User curUser = DBSourceController.GetCurrentUser(sid);
                if (curUser == null)
                {
                    reJo.msg = "登录验证失败！请尝试重新登录！";
                    return reJo.Value;
                }

                DBSource dbsource = curUser.dBSource;
                if (dbsource == null)
                {
                    reJo.msg = "登录验证失败！请尝试重新登录！";
                    return reJo.Value;
                }

                Doc m_Doc = dbsource.GetDocByKeyWord(DocKeyword);

                if (m_Doc == null)
                {
                    reJo.msg = "参数错误！文档不存在！";
                    return reJo.Value;
                }

                //项目代码
                string RootProjectCode = m_Doc.GetValueByKeyWord("HXNY_DOCUMENTSYSTEM_CODE");

                //项目名称
                string RootProjectDesc = m_Doc.GetValueByKeyWord("HXNY_DOCUMENTSYSTEM_DESC");

                //来文单位
                string CommUnit = "";
                string CommUnitCode = "";
                if (m_Doc.Project.TempDefn.KeyWord == "COM_UNIT")
                {
                    CommUnit = m_Doc.Project.Description;
                    CommUnitCode = m_Doc.Project.Code;
                }

                
                Project recTypeProject = CommonFunction.getParentProjectByTempDefn(m_Doc.Project, "COM_COMTYPE");

                #region 获取收文编号
                string recType = "";
                if (recTypeProject.Code == "信函" || recTypeProject.Description == "信函")
                {
                    recType = "LET";
                }
                else if (recTypeProject.Code == "文件传递单" || recTypeProject.Description == "文件传递单")
                {
                    recType = "TRA";
                }

                Project recUnitProject = CommonFunction.getParentProjectByTempDefn(m_Doc.Project, "HXNY_DOCUMENTSYSTEM");

                string recUnitCode = "";
                if (recUnitProject != null)
                {
                    recUnitCode = recUnitProject.GetValueByKeyWord("PRO_COMPANY");
                }

                string ONShoreCommCode = "";
                if (recUnitProject != null)
                {
                    ONShoreCommCode = recUnitProject.GetValueByKeyWord("RPO_ONSHORE");
                }

                string OFFShoreCommCode = "";
                if (recUnitProject != null)
                {
                    OFFShoreCommCode = recUnitProject.GetValueByKeyWord("RPO_OFFSHORE");
                }

                //string runNum = getDocNumber(dbsource, RootProjectCode, recType, CommUnitCode, recUnitCode);
                //if (string.IsNullOrEmpty(runNum)) runNum = "001";

                string recNumber = Document.getDocTempNumber(dbsource, RootProjectCode, "R", recType, CommUnitCode, recUnitCode);

               // string recNumber = RootProjectCode + "-" + CommUnitCode + "-" + recUnitCode + "-" + recType + "-" + runNum;
                #endregion

                string recCode = m_Doc.Code;

                //string DocNumber = getDocNumber(m_Doc.Project, RootProjectCode, strCompany);
                string strDesc = m_Doc.O_itemname;

                JObject joRecCompany = new JObject();
                JObject joSendCompany = new JObject();

                List<DictData> dictDataList = dbsource.GetDictDataList("Unit");
                foreach (DictData data6 in dictDataList)
                {
                    if (!string.IsNullOrEmpty(data6.O_sValue3)&& data6.O_sValue3==curUser.ToString)
                    {

                        joSendCompany.Add(new JProperty(data6.O_Code, data6.O_Desc));
                    }
                }

                if (!string.IsNullOrEmpty(ONShoreCommCode))
                {
                    joRecCompany.Add(new JProperty(ONShoreCommCode, ONShoreCommCode));
                }

                if (!string.IsNullOrEmpty(OFFShoreCommCode)) {
                    joRecCompany.Add(new JProperty(OFFShoreCommCode, OFFShoreCommCode));
                }

                JObject joData = new JObject(
                    new JProperty("RootProjectCode", RootProjectCode),
                    new JProperty("RootProjectDesc", RootProjectDesc),
                    //new JProperty("DocNumber", DocNumber),
                    new JProperty("DraftmanCode", curUser.Code),
                    new JProperty("DraftmanDesc", curUser.Description),
                    new JProperty("CommUnit", CommUnit),
                    new JProperty("RecCode", recCode),
                    new JProperty("RecNumber", recNumber),
                     new JProperty("RecCompanyList", joRecCompany),
                    new JProperty("SendCompanyList", joSendCompany),
                    new JProperty("DocType",recType)


                    );

                reJo.data = new JArray(joData);
                reJo.success = true;
                return reJo.Value;
            }
            catch (Exception e)
            {
                reJo.msg = e.Message;
                CommonController.WebWriteLog(reJo.msg);
            }
            return reJo.Value;
        }

        public static JObject GetRecDocumentNumber(string sid,string ProjectCode, string DocType, string SendCompany, string RecCompany)
        {
            ExReJObject reJo = new ExReJObject();
            try
            {
                User curUser = DBSourceController.GetCurrentUser(sid);
                if (curUser == null)
                {
                    reJo.msg = "登录验证失败！请尝试重新登录！";
                    return reJo.Value;
                }

                DBSource dbsource = curUser.dBSource;
                if (dbsource == null)
                {
                    reJo.msg = "登录验证失败！请尝试重新登录！";
                    return reJo.Value;
                }

                ///Document.getDocNumber(dbsource, ProjectCode,"LET", SendCompany, RecCompany);
                string runNum = Document.getDocNumber(dbsource, ProjectCode, DocType, SendCompany, RecCompany);
                if (string.IsNullOrEmpty(runNum)) runNum = "001";
                reJo.success = true;
                reJo.data = new JArray(new JObject(new JProperty("RunNum", runNum)));



            }
            catch (Exception exception)
            {
                WebApi.CommonController.WebWriteLog(exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace);
                //AssistFun.PopUpPrompt(exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace);
                reJo.msg = "获取信函编号失败！" + exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace;
            }
            return reJo.Value;
        }


        /// <summary>
        /// 收文流程设置通过回复并提交到下一流程
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="DocKeyword"></param>
        /// <returns></returns>
        public static JObject RecWorflowPassReplyState(string sid, string DocKeyword)
        {
            ExReJObject reJo = new ExReJObject();
            try
            {
                User curUser = DBSourceController.GetCurrentUser(sid);
                if (curUser == null)
                {
                    reJo.msg = "登录验证失败！请尝试重新登录！";
                    return reJo.Value;
                }

                DBSource dbsource = curUser.dBSource;
                if (dbsource == null)
                {
                    reJo.msg = "登录验证失败！请尝试重新登录！";
                    return reJo.Value;
                }

                Doc m_Doc = dbsource.GetDocByKeyWord(DocKeyword);

                if (m_Doc == null)
                {
                    reJo.msg = "参数错误！文档不存在！";
                    return reJo.Value;
                }

                WorkFlow flow = m_Doc.WorkFlow;
                flow.O_suser3 = "pass";
                flow.Modify();

                WorkStateBranch wsb = null;
                // m_Doc.WorkFlow.
                wsb = flow.CuWorkState.workStateBranchList.Find(w=>w.defStateBrach.O_Description == "回复");
                if (wsb == null) {
                    reJo.msg = "流程分支不存在！";
                    return reJo.Value;
                }

                ExReJObject GotoNextReJo = WebWorkFlowEvent.GotoNextStateAndSelectUser(wsb);// flow.CuWorkState.workStateBranchList[0]);

                if (!GotoNextReJo.success)
                {
                }
                    reJo.success = true;
                return reJo.Value;
            }
            catch (Exception e)
            {
                reJo.msg = e.Message;
                CommonController.WebWriteLog(reJo.msg);
            }
            return reJo.Value;
        }

        //获取文档最大编号
        private static string getDocNumber(DBSource dbsource, string RootProjectCode,string RecType, string strSendCompany, string strRecCompany)
        {
            try
            {
                //string RootProjectCode = proj.GetValueByKeyWord("DESIGNPROJECT_CODE");

                //获取文档前缀
                string sendCompanyCode = strSendCompany.IndexOf("__") >= 0 ? strSendCompany.Substring(0, strSendCompany.IndexOf("__")) : strSendCompany;

                string recCompanyCode = strRecCompany.IndexOf("__") >= 0 ? strRecCompany.Substring(0, strRecCompany.IndexOf("__")) : strRecCompany;


                //编码前缀
                string strPrefix = RootProjectCode + "-" + sendCompanyCode + "-" + recCompanyCode + "-"+ RecType + "-";

                List<Doc> docList = dbsource.SelectDoc(string.Format("select * from CDMS_Doc where o_itemname like '%{0}%' and o_dmsstatus !=10 order by o_itemname", strPrefix));
                if (docList == null || docList.Count == 0)
                {
                    return "001";
                }
                else
                {
                    Doc doc = docList[docList.Count - 1];
                    //int docCodeLength = 0;
                    //docCodeLength = strPrefix.IndexOf("-LET-");
                    //docCodeLength = docCodeLength > 0 ? 0 : docCodeLength + 3 + 1;
                    //int tempNum = Convert.ToInt32(doc.O_itemname.Substring(docCodeLength, 3));

                    int tempNum = Convert.ToInt32(doc.O_itemname.Substring(strPrefix.Length, 3));
                    //3位数，不够位数补零
                    return (tempNum + 1).ToString("d3");
                }
            }
            catch
            {
                return "001";
            }

        }

    }
}

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
using LinqToDB;

namespace AVEVA.CDMS.HXEPC_Plugins
{
    public class Document
    {

        //线程锁 
        internal static Mutex muxConsole = new Mutex();

        /// <summary>
        /// 获取创建信函表单的默认配置
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="ProjectKeyword"></param>
        /// <returns></returns>
        public static JObject GetDraftLetterCNDefault(string sid, string ProjectKeyword,string DraftOnProject)
        {
            return LetterCN.GetDraftLetterCNDefault(sid, ProjectKeyword, DraftOnProject);
        }

        /// <summary>
        /// 获取回复信函表单的默认配置
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="DocKeyword"></param>
        /// <returns></returns>
        public static JObject GetReplyLetterCNDefault(string sid, string DocKeyword)
        {
            return LetterCN.GetReplyLetterCNDefault(sid, DocKeyword);
        }

        /// <summary>
        /// 收文流程设置通过回复并提交到下一流程
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="DocKeyword"></param>
        /// <returns></returns>
        public static JObject RecWorflowPassReplyState(string sid, string DocKeyword)
        {
            return RecDocument.RecWorflowPassReplyState(sid, DocKeyword);
        }

        public static JObject GetLetterCNNumber(string sid, string ProjectCode, string SendCompany, string RecCompany)
        {
            return LetterCN.GetLetterCNNumber(sid, ProjectCode, SendCompany, RecCompany);
        }

        public static JObject GetDepartmentSecUser(string sid, string DepartmentList)
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

                string[] strArry = DepartmentList.Split(new char[] { ',' });
                //bool isAdd = false;
                string resultUserList = "";
                foreach (string strDepartment in strArry)
                {
                    resultUserList = resultUserList + CommonFunction.GetSecUserByDepartmentCode(dbsource, strDepartment) + ";";
                }

                if (!string.IsNullOrEmpty(resultUserList))
                {
                    resultUserList = resultUserList.Substring(0, resultUserList.Length - 1);
                }

                reJo.data = new JArray(new JObject(new JProperty("userList", resultUserList)));

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

        public static JObject DraftDocument(string sid, string ProjectKeyword, string DocAttrJson)
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
                JArray jaAttr = (JArray)JsonConvert.DeserializeObject(DocAttrJson);

                Project m_Project = dbsource.GetProjectByKeyWord(ProjectKeyword);

                //定位到发文目录
                //m_Project = LocalProject(m_Project);

                if (m_Project == null)
                {
                    reJo.msg = "参数错误！文件夹不存在！";
                    return reJo.Value;
                }

                #region 获取项目参数项目

                //获取项目参数项目
                string documentCode = "", title = "", deliveryUnit = "",
                    content = "", contact = "", tel = "";



                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    //获取函件编号
                    if (strName == "documentCode") documentCode = strValue.Trim();

                    //获取标题
                    else if (strName == "title") title = strValue;

                    //获取主送单位
                    else if (strName == "deliveryUnit") deliveryUnit = strValue;

                    //获取正文内容
                    else if (strName == "content") content = strValue;

                    //获取联系人
                    else if (strName == "contact") contact = strValue;

                    //获取联系电话
                    else if (strName == "tel") tel = strValue;

                }


                if (string.IsNullOrEmpty(documentCode))
                {
                    reJo.msg = "请填写函件编号！";
                    return reJo.Value;
                }
                else if (string.IsNullOrEmpty(title))
                {
                    reJo.msg = "请填写函件标题！";
                    return reJo.Value;
                }
                else if (string.IsNullOrEmpty(deliveryUnit))
                {
                    reJo.msg = "请填写主送单位！";
                    return reJo.Value;
                }

                #endregion

                reJo.data = new JArray(new JObject(new JProperty("projectKeyword", m_Project.KeyWord)));
                reJo.success = true;
                return reJo.Value;

                //AVEVA.CDMS.WebApi.DBSourceController.RefreshDBSource(sid);


            }
            catch (Exception e)
            {
                reJo.msg = e.Message;
                CommonController.WebWriteLog(reJo.msg);
            }

            return reJo.Value;
        }

        /// <summary>
        /// 获取创建信函表单的默认配置
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="ProjectKeyword"></param>
        /// <returns></returns>
        public static JObject GetDraftMeetMinutesCNDefault(string sid, string ProjectKeyword)
        {
            return MeetMinutesCN.GetDraftMeetMinutesCNDefault(sid, ProjectKeyword);
        }

        public static JObject GetMeetMinutesCNNumber(string sid, string ProjectCode, string SendCompany, string RecCompany)
        {
            return MeetMinutesCN.GetMeetMinutesCNNumber(sid, ProjectCode, SendCompany, RecCompany);
        }

        public static JObject DraftMeetMinutesCN(string sid, string ProjectKeyword, string DocAttrJson)
        {
            return MeetMinutesCN.DraftMeetMinutesCN(sid, ProjectKeyword, DocAttrJson);
        }


        public static JObject DraftLetterCN(string sid, string ProjectKeyword, string DocAttrJson, string FileListJson)
        {
            return LetterCN.DraftLetterCN(sid, ProjectKeyword, DocAttrJson, FileListJson);
        }

        public static JObject GetDraftTransmittalCNDefault(string sid, string ProjectKeyword)
        {
            return TransmittalCN.GetDraftTransmittalCNDefault(sid, ProjectKeyword);
        }

        public static JObject GetTransmittalCNNumber(string sid, string ProjectCode, string SendCompany, string RecCompany)
        {
            return TransmittalCN.GetTransmittalCNNumber(sid, ProjectCode, SendCompany, RecCompany);
        }
        public static JObject DraftTransmittalCN(string sid, string ProjectKeyword, string DocAttrJson, string FileListJson)
        {
            return TransmittalCN.DraftTransmittalCN(sid, ProjectKeyword, DocAttrJson, FileListJson);
        }

        /// <summary>
        /// 启动信函表单发文流程
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="docKeyword"></param>
        /// <param name="DocList"></param>
        /// <param name="ApprovPath"></param>
        /// <param name="UserList"></param>
        /// <returns></returns>
        public static JObject LetterStartWorkFlow(string sid, string docKeyword,
            string DocList, string ApprovPath, string UserList,string SendUnitCode) {
            return LetterCN.LetterStartWorkFlow(sid, docKeyword, DocList, ApprovPath, UserList, SendUnitCode);
        }

        /// <summary>
        /// 流程流转到文控时，文控填写收发文单位和发文编码
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="DocKeyword"></param>
        /// <returns></returns>
        public static JObject LetterCNFillInfo(string sid, string DocKeyword, string DocAttrJson)
        {
            return LetterCN.LetterCNFillInfo(sid, DocKeyword, DocAttrJson);
        }

        //拖拽文件后，发起收文流程
        public static JObject ReceiveDocument(string sid, string DocKeyword, string docAttrJson)
        {
            return RecDocument.ReceiveDocument(sid, DocKeyword, docAttrJson);
        }

        public static JObject GetReceiveDocumentDefault(string sid, string DocKeyword)
        {
            return RecDocument.GetReceiveDocumentDefault(sid, DocKeyword);
        }

        public static JObject GetDraftRecognitionDefault(string sid, string ProjectKeyword)
        {
            return Recognition.GetDraftRecognitionDefault(sid, ProjectKeyword);
        }

        public static JObject DraftRecognition(string sid, string ProjectKeyword, string DocAttrJson, string ContentJson)
        {
            return Recognition.DraftRecognition(sid, ProjectKeyword, DocAttrJson, ContentJson);
        }

        public static JObject RecognitionStartWorkFlow(string sid, string docKeyword, string DocList, string UserList) {
            return Recognition.RecognitionStartWorkFlow(sid, docKeyword, DocList, UserList);
        }

        public static JObject GetRecDocumentNumber(string sid, string ProjectCode, string DocType, string SendCompany, string RecCompany)
        {
            return RecDocument.GetRecDocumentNumber( sid, ProjectCode, DocType, SendCompany,RecCompany);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="DocKeyword"></param>
        /// <param name="Operator">主办人</param>
        /// <param name="Coordinator">协办人</param>
        /// <returns></returns>
        public static JObject SendDistriProcess(string sid, string DocKeyword, string Operator, string Coordinator)
        {
            // return RecDocument.SendDistriProcess(sid, DocKeyword, Operator, Coordinator);

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

                Doc doc = dbsource.GetDocByKeyWord(DocKeyword);
                if (doc == null)
                {
                    reJo.msg = "错误的文档操作信息！指定的文档不存在！";
                    return reJo.Value;
                }
                WorkFlow flow;
                if ((flow = doc.WorkFlow) == null)
                {
                    reJo.msg = "错误的文档操作信息！指定的文档流程不存在！";
                    return reJo.Value;
                }

                #region 设置本部门文控（登录用户）到部门人员状态

                Server.Group group = new Server.Group();

                group.AddUser(curUser);

                DefWorkState defWorkStateCu = flow.DefWorkFlow.DefWorkStateList.Find(dwsx => dwsx.O_Code == "DEPARTMENTCONTROL");
                WorkState stateCu = flow.NewWorkState(defWorkStateCu);
                stateCu.SaveSelectUser(group);

                stateCu.IsRuning = true;

                stateCu.PreWorkState = flow.CuWorkState;
                stateCu.O_iuser5 = new int?(flow.CuWorkState.O_stateno);
                stateCu.Modify();

  
                #endregion



                #region 设置协办部门人员状态

                string[] strArry2 = Coordinator.Split(new char[] { ';' });
                
                //每个部门一个协办分支
                foreach (string op in strArry2)
                {
                    string strUser = op.IndexOf("_") >= 0 ? op.Substring(0, op.IndexOf("_")) : op;
                    Server.Group CoordinatorGroup = new Server.Group();
                    User CoordinatorUser = dbsource.GetUserByCode(strUser);
                    if (CoordinatorUser != null)
                    {
                        CoordinatorGroup.AddUser(CoordinatorUser);

                        DefWorkState defWorkStateCo = flow.DefWorkFlow.DefWorkStateList.Find(dwsx => dwsx.O_Code == "DEPARTMENTCONTROL");
                        WorkState state2 = flow.NewWorkState(defWorkStateCo);
                        state2.SaveSelectUser(CoordinatorGroup);

                        state2.IsRuning = true;

                        state2.PreWorkState = flow.CuWorkState;
                        state2.O_iuser5 = new int?(flow.CuWorkState.O_stateno);
                        state2.Modify();

                    }
                }

                #endregion

                WorkStateBranch branch = flow.CuWorkState.workStateBranchList.Find(wsb=>wsb.KeyWord== "TOCONTROL1");//[0];
                branch.NextStateAddGroup(group);

                ExReJObject GotoNextReJo = WebWorkFlowEvent.GotoNextStateAndSelectUser(flow.CuWorkState.workStateBranchList[0]);

                #region 设置本部门办理人状态
                string[] strArry = Operator.Split(new char[] { ';' });
                Server.Group OperatorGroup = new Server.Group();

                foreach (string op in strArry)
                {
                    User OperatorUser = dbsource.GetUserByCode(op);
                    if (OperatorUser != null)
                    {
                        OperatorGroup.AddUser(OperatorUser);
                    }
                }



                //放置本部门办理状态人员
                //WorkState state = flow.WorkStateList.Find(wsx => (wsx.Code == "MAINHANDLE") && 
                //    (wsx.CheckGroup.AllUserList.Count == 0));

                //if (state == null)
                //{
                DefWorkState defWorkState = flow.DefWorkFlow.DefWorkStateList.Find(dwsx => dwsx.O_Code == "MAINHANDLE");
                WorkState state = flow.NewWorkState(defWorkState);
                state.SaveSelectUser(OperatorGroup);

                state.IsRuning = true;

                state.PreWorkState = stateCu;
                state.O_iuser5 = new int?(stateCu.O_stateno);
                state.Modify();
                // } 
                #endregion


                //reJo.data = new JArray(new JObject(new JProperty("WorkFlowKeyword", doc.WorkFlow.KeyWord)));

                reJo.success = true;
                return reJo.Value;
            }
            catch (Exception exception)
            {
                //WebApi.CommonController.WebWriteLog(exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace);
                reJo.msg = "分发失败！" + exception.Message;
            }
            return reJo.Value;
        }

        //发文流程盖章
        public static JObject DocumenteSeal(string sid, string DocKeyword,string isSeal) {

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

                Doc doc = dbsource.GetDocByKeyWord(DocKeyword);
                if (doc == null)
                {
                    reJo.msg = "错误的文档操作信息！指定的文档不存在！";
                    return reJo.Value;
                }

                if (doc.WorkFlow == null) {
                    reJo.msg = "错误的文档操作信息！指定的文档流程不存在！";
                    return reJo.Value;
                }

                //录入数据进入表单
                //格式化日期
                DateTime inscribedate = Convert.ToDateTime(DateTime.Now);
                string strInscribedate = inscribedate.ToShortDateString().ToString().Replace("-", ".").Replace("/", ".");

                Hashtable htUserKeyWord = new Hashtable();

                string inscribe = "", sendercode = "";
                AttrData data;
                //获取发送方
                if ((data = doc.GetAttrDataByKeyWord("CA_SENDER")) != null)
                {
                    inscribe = data.ToString;
                }

                //获取发送方代码
                if ((data = doc.GetAttrDataByKeyWord("CA_SENDERCODE")) != null)
                {
                    sendercode = data.ToString;
                }

                if (isSeal == "true")
                {

                    if (!string.IsNullOrEmpty(sendercode))
                    {
                        //盖章图片名
                        htUserKeyWord["SEAL"] = sendercode + "SEAL";
                    }
                    else {
                        //盖章图片名
                        htUserKeyWord["SEAL"] = "BLANK" + "SEAL";
                    }

                }
                else {
                    //盖章图片名
                    htUserKeyWord["SEAL"] = "BLANK" + "SEAL";
                }


                //落款
                htUserKeyWord["INSCRIBE"] = inscribe;
                //落款时间
                htUserKeyWord["INSCRIBEDATE"] = strInscribedate;



                //获取即将生成的联系单文件路径
                string locFileName = doc.FullPathFile;

                FileInfo info = new FileInfo(locFileName);

                if (System.IO.File.Exists(locFileName))
                {

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
                        office.WriteDataToDocument(doc, locFileName, htUserKeyWord, htUserKeyWord);
                    }
                    catch { }
                    finally
                    {

                        //解锁
                        muxConsole.ReleaseMutex();
                    }
                }

                if (doc.WorkFlow.CuWorkState.Code == "APPROV" &&
                      dbsource.LoginUser.O_userno == doc.WorkFlow.CuWorkState.CuWorkUser.O_userno)
                {

                    doc.WorkFlow.O_suser3 = "approvpass";
                    doc.WorkFlow.Modify();

                    DBSourceController.refreshDBSource(sid);
                }

                reJo.data = new JArray(new JObject(new JProperty("WorkFlowKeyword", doc.WorkFlow.KeyWord)));

                reJo.success = true;
                return reJo.Value;
            }
            catch (Exception exception)
            {
                 //WebApi.CommonController.WebWriteLog(exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace);
                reJo.msg = "盖章失败！" + exception.Message;
            }
            return reJo.Value;
        }

        public static JObject ResetFileCode(string sid, string DocKeyword, string DocAttrJson)
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

                Doc doc = dbsource.GetDocByKeyWord(DocKeyword);
                if (doc == null)
                {
                    reJo.msg = "错误的文档操作信息！指定的文档不存在！";
                    return reJo.Value;
                }

                #region 获取信函参数内容

                //获取信函参数内容
                string mainFeeder = "", copyParty = "", sender = "",
                   sendCode = "", recCode = "", totalPages = "",
                   urgency = "", sendDate = "", seculevel = "",
                   secrTerm = "", needreply = "", replyDate = "",
                   title = "", content = "", approvpath = "",
                   nextStateUserList = "";

                JArray jaAttr = (JArray)JsonConvert.DeserializeObject(DocAttrJson);

                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    ////获取函件编号
                    //if (strName == "documentCode") documentCode = strValue.Trim();

                    //获取主送
                    if (strName == "mainFeeder") mainFeeder = strValue.Trim();

                    //获取抄送
                    else if (strName == "copyParty") copyParty = strValue;

                    //获取发送方
                    else if (strName == "sender") sender = strValue;

                    //获取发文编码
                    else if (strName == "sendCode") sendCode = strValue;

                    //获取发送日期
                    else if (strName == "sendDate") sendDate = strValue;

                }


                if (string.IsNullOrEmpty(sendCode))
                {
                    reJo.msg = "请填写函件编号！";
                    return reJo.Value;
                }
                //else if (string.IsNullOrEmpty(mainFeeder))
                //{
                //    reJo.msg = "请填写主送！";
                //    return reJo.Value;
                //}
                //else if (string.IsNullOrEmpty(sender))
                //{
                //    reJo.msg = "请选择发送方！";
                //    return reJo.Value;
                //}




                #endregion



                #region 填写发文信息进入word表单

                //获取即将函件的文件路径
                string locFileName = doc.FullPathFile;

                if (string.IsNullOrEmpty(locFileName))
                {
                    reJo.msg = "填写发文信息错误，获取函件文件失败！";
                    return reJo.Value;
                }
                if (!System.IO.File.Exists(locFileName))
                {
                    reJo.msg = "填写发文信息失败，函件文件不存在！";
                    return reJo.Value;
                }

                //格式化日期
                DateTime senddate = Convert.ToDateTime(sendDate);
                string strSenddate = senddate.ToShortDateString().ToString().Replace("-", ".").Replace("/", ".");
                string str = doc.O_filename.ToUpper();

                Hashtable htUserKeyWord = new Hashtable();
                //htUserKeyWord.Add("MAINFEEDER", mainFeeder);//主送
                //htUserKeyWord.Add("SENDER", sender);//发送方
                htUserKeyWord.Add("SENDCODE", sendCode);//发文编码
                //htUserKeyWord.Add("COPY", copyParty);//抄送
                //htUserKeyWord.Add("SENDDATE", strSenddate);//发送日期 

                if ((str.EndsWith(".DOC") || str.EndsWith(".DOCX")) || (str.EndsWith(".XLS") || str.EndsWith(".XLSX")))
                {
                    //线程锁 
                    muxConsole.WaitOne();
                    try
                    {
                        WebApi.CDMSWebOffice office = new WebApi.CDMSWebOffice
                        {
                            CloseApp = true,
                            VisibleApp = false
                        };
                        office.Release(true);
                        if (doc.WorkFlow != null)
                        {
                            enWorkFlowStatus status1 = doc.WorkFlow.O_WorkFlowStatus;
                        }
                        office.WriteDataToDocument(doc, locFileName, htUserKeyWord, htUserKeyWord);
                    }
                    catch (Exception ExOffice)
                    {
                        WebApi.CommonController.WebWriteLog(ExOffice.Message);
                    }
                    finally
                    {

                        //解锁
                        muxConsole.ReleaseMutex();
                    }
                }
                #endregion


                if (doc.WorkFlow != null)
                {
                        AttrData data;
                    string orgSendCode = "";
                    if ((data = doc.GetAttrDataByKeyWord("CA_SENDCODE")) != null)
                    {
                        orgSendCode = data.ToString;
                        //FileCode = FileCode + data.ToString + "-";

                    }

                    string orgName = doc.O_itemname;
                    foreach (Doc docItem in doc.WorkFlow.DocList)
                    {
                        if (docItem != doc)
                        {
                            docItem.O_itemname = docItem.O_itemname.Replace(orgSendCode, sendCode);
                            docItem.Modify();
                        }
                    }
                    doc.O_itemname = doc.O_itemname.Replace( orgSendCode, sendCode);
                    // doc.O_filename = regex.Replace(doc.O_filename, replacement);

                    string fullPathFile = doc.FullPathFile;
                    fullPathFile = fullPathFile.Replace('/', '\\');
                    string newFileName = doc.FullPathFile.Replace(orgSendCode, sendCode);

                    try
                    {
                        File.Move(fullPathFile, newFileName);
                    }
                    catch { }

                    doc.O_filename = doc.O_filename.Replace(orgSendCode, sendCode);

                    #region 设置信函文档附加属性



                    ////主送
                    //if ((data = doc.GetAttrDataByKeyWord("CA_MAINFEEDER")) != null)
                    //{
                    //    data.SetCodeDesc(mainFeeder);
                    //}
                    ////发送方
                    //if ((data = doc.GetAttrDataByKeyWord("CA_SENDER")) != null)
                    //{
                    //    data.SetCodeDesc(sender);
                    //}
                    //发文编码
                    if ((data = doc.GetAttrDataByKeyWord("CA_SENDCODE")) != null)
                    {
                        data.SetCodeDesc(sendCode);
                    }

                    ////抄送
                    //if ((data = doc.GetAttrDataByKeyWord("CA_COPY")) != null)
                    //{
                    //    data.SetCodeDesc(copyParty);
                    //}

                    //////发送日期
                    //if ((data = doc.GetAttrDataByKeyWord("CA_SENDDATE")) != null)
                    //{
                    //    data.SetCodeDesc(sendDate);
                    //}

                    ////保存项目属性，存进数据库
                    doc.AttrDataList.SaveData();

                    #endregion

                    doc.Modify();

                    doc.WorkFlow.O_suser3 = "pass";
                    doc.WorkFlow.Modify();

                    DBSourceController.refreshDBSource(sid);
                }



                reJo.data = new JArray(new JObject(new JProperty("WorkFlowKeyword", doc.WorkFlow.KeyWord)));

                reJo.success = true;
                return reJo.Value;
            }
            catch (Exception exception)
            {
                //WebApi.CommonController.WebWriteLog(exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace);
                reJo.msg = "填写发文信息失败，请手动填写发文信息！" + exception.Message;
            }
            return reJo.Value;

        }

        ////
        ///// <summary>
        ///// 获取文件编码流水号
        ///// </summary>
        ///// <param name="sid"></param>
        ///// <param name="ProjectKeyword"></param>
        ///// <param name="Prefix">文件编码前缀</param>
        ///// <returns></returns>
        //public static JObject GetFileCodeNumber(string sid,string ProjectKeyword,string Prefix)
        //{
        //    ExReJObject reJo = new ExReJObject();
        //    try
        //    {
        //        User curUser = DBSourceController.GetCurrentUser(sid);
        //        if (curUser == null)
        //        {
        //            reJo.msg = "登录验证失败！请尝试重新登录！";
        //            return reJo.Value;
        //        }

        //        DBSource dbsource = curUser.dBSource;
        //        if (dbsource == null)
        //        {
        //            reJo.msg = "登录验证失败！请尝试重新登录！";
        //            return reJo.Value;
        //        }

        //        Project m_Project = dbsource.GetProjectByKeyWord(ProjectKeyword);

        //        if (m_Project == null)
        //        {
        //            reJo.msg = "参数错误！文件夹不存在！";
        //            return reJo.Value;
        //        }

        //        List<Doc> docList = dbsource.SelectDoc(string.Format("select * from CDMS_Doc where o_itemname like '%{0}%' and o_dmsstatus !=10 order by o_itemname", Prefix));
        //        if (docList == null || docList.Count == 0)
        //        {
        //            reJo.data = new JArray(new JObject(
        //                new JProperty("Number", "001")));
        //            //new JProperty("EditTion","A")));
        //            reJo.success = true;
        //            return reJo.Value;
        //        }
        //        else
        //        {
        //            Doc doc = docList[docList.Count - 1];


        //            int tempNum = Convert.ToInt32(doc.O_itemname.Substring(Prefix.Length, 3));
      
        //            //3位数，不够位数补零
        //            string strNum = (tempNum + 1).ToString("d3");
        //            reJo.data = new JArray(new JObject(
        //                 new JProperty("Number", strNum)));

        //            reJo.success = true;
        //            return reJo.Value;
        //        }

        //        //reJo.success = true;
        //        //return reJo.Value;
        //    }
        //    catch (Exception e)
        //    {
        //        reJo.msg = e.Message;
        //        CommonController.WebWriteLog(reJo.msg);
        //    }
        //    return reJo.Value;

        //}

        /// <summary>
        /// 获取文件临时的后台编码
        /// </summary>
        /// <param name="dbsource"></param>
        /// <param name="RootProjectCode"></param>
        /// <param name="CommType">收文还是发文</param>
        /// <param name="docType"></param>
        /// <param name="strSendCompany"></param>
        /// <param name="strRecCompany"></param>
        /// <returns></returns>
         internal static string getDocTempNumber(DBSource dbsource, string RootProjectCode, string commType, string docType, string strSendCompany, string strRecCompany)
        {
            try
            {
                //string RootProjectCode = proj.GetValueByKeyWord("DESIGNPROJECT_CODE");

                //获取文档前缀
                string sendCompanyCode = strSendCompany.IndexOf("__") >= 0 ? strSendCompany.Substring(0, strSendCompany.IndexOf("__")) : strSendCompany;

                string recCompanyCode = strRecCompany.IndexOf("__") >= 0 ? strRecCompany.Substring(0, strRecCompany.IndexOf("__")) : strRecCompany;


                //编码前缀
                string strPrefix = "";

                strPrefix = commType + docType;

                //if (!string.IsNullOrEmpty(RootProjectCode))
                //{
                //    //项目管理类
                //    //strPrefix = RootProjectCode + "-" + sendCompanyCode + "-" + recCompanyCode + "-" + docType + "-" + commType + docType;
                //    strPrefix = RootProjectCode + "-" + commType + docType;
                //}
                //else
                //{
                //    //运营管理类
                //    strPrefix =  commType + docType;
                //}

                List<Doc> docList = new List<Doc>();
                //信函
                #region 查找信函流水号
               // if (commType == "S" && docType == "LET")
                {
                    string[] strArry = dbsource.DBExecuteSQL(string.Format(
                            "select lf.CA_FILEID from CDMS_Doc as cd inner join " +
                              "(select Itemno, CA_FILEID from User_CATALOGUING where CA_FILEID like '%{0}%') as lf " +
                              "on  cd.o_itemno = lf.Itemno " +
                              " where cd.o_dmsstatus != 10 order by lf.CA_FILEID ",
                        strPrefix));

                    if (strArry == null || strArry.Length == 0)
                    {
                        return commType + docType + "00001";
                    }
                    else
                    {
                        //5位数，不够位数补零
                        int tempNum = Convert.ToInt32((strArry[strArry.Length - 1]).Substring(strPrefix.Length, 5));

                        return commType + docType + (tempNum + 1).ToString("d5");
                    }
                } 
                #endregion

                //else
                //{
                //    docList = dbsource.SelectDoc(string.Format("select * from CDMS_Doc where o_itemname like '%{0}%' and o_dmsstatus !=10 order by o_itemname", strPrefix));


                //    if (docList == null || docList.Count == 0)
                //    {
                //        //return "SLET"+"00001"; "-" + docType + "-" + commType + docType;
                //        return commType + docType + "00001";
                //    }
                //    else
                //    {
                //        Doc doc = docList[docList.Count - 1];

                //        int tempNum = Convert.ToInt32(doc.O_itemname.Substring(strPrefix.Length, 5));
                //        //3位数，不够位数补零
                //        //return "SLET" + (tempNum + 1).ToString("d5");
                //        return commType + docType + (tempNum + 1).ToString("d5");
                //    }
                //}
            }
            catch
            {
                //return "SLET" + "00001";
                return commType + docType + "00001";
            }

        }

        internal static string getDocNumber(DBSource dbsource, string RootProjectCode,  string docType, string strSendCompany, string strRecCompany) {
            try
            {
                //string RootProjectCode = proj.GetValueByKeyWord("DESIGNPROJECT_CODE");

                //获取文档前缀
                string sendCompanyCode = strSendCompany.IndexOf("__") >= 0 ? strSendCompany.Substring(0, strSendCompany.IndexOf("__")) : strSendCompany;

                string recCompanyCode = strRecCompany.IndexOf("__") >= 0 ? strRecCompany.Substring(0, strRecCompany.IndexOf("__")) : strRecCompany;


                //编码前缀
                string strPrefix = "";
                if (!string.IsNullOrEmpty(RootProjectCode))
                {
                    //项目管理类
                    //strPrefix = RootProjectCode + "-" + sendCompanyCode + "-" + recCompanyCode + "-LET-" + "SLET";
                    strPrefix = RootProjectCode + "-" + sendCompanyCode + "-" + recCompanyCode + "-" + docType + "-" ;
                }
                else
                {
                    //运营管理类
                    //strPrefix = sendCompanyCode + "-" + recCompanyCode + "-LET-" + "SLET";
                    strPrefix = sendCompanyCode + "-" + recCompanyCode + "-" + docType + "-" ;
                }

                List<Doc> docList = dbsource.SelectDoc(string.Format(
                    "select * from CDMS_Doc where o_itemname like '%{0}[0-9]%' and o_dmsstatus !=10 order by o_itemname",
                    strPrefix));
                if (docList == null || docList.Count == 0)
                {
                    //return "SLET"+"00001"; "-" + docType + "-" + commType + docType;
                    return  "001";
                }
                else
                {
                    Doc doc = docList[docList.Count - 1];

                    int tempNum = Convert.ToInt32(doc.O_itemname.Substring(strPrefix.Length, 3));
                    //3位数，不够位数补零
                    //return "SLET" + (tempNum + 1).ToString("d5");
                    return (tempNum + 1).ToString("d3");
                }
            }
            catch
            {
                //return "SLET" + "00001";
                return  "001";
            }
        }

        //获取文件著录编码流水号
        public static JObject GetFileCodeNumber(string sid,string FileCodePerfix) {
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

                string runNum = "";

                string strSql = string.Format(
           "select lf.FILECODE from CDMS_Doc as cd inner join " +
             "(select Itemno, CA_FILECODE as FILECODE from User_CATALOGUING where CA_FILECODE like '%{0}%')" +
             " as lf " +
             "on  cd.o_itemno = lf.Itemno " +
             " where cd.o_dmsstatus != 10 order by lf.FILECODE ",
       FileCodePerfix);

                string[] strArry = dbsource.DBExecuteSQL(strSql);

                if (strArry == null || strArry.Length == 0 ||
                    (strArry[strArry.Length - 1]).Length < FileCodePerfix.Length + 3
                    )
                {
                    runNum=  "001";
                    
                }
                else
                {

                    //5位数，不够位数补零
                    int tempNum = Convert.ToInt32((strArry[strArry.Length - 1]).Substring(FileCodePerfix.Length, 3));

                    runNum = (tempNum + 1).ToString("d3");
                }
                reJo.success = true;
                reJo.data = new JArray(new JObject(new JProperty("RunNum", runNum)));
                return reJo.Value;
            }
            catch (Exception exception)
            {
                WebApi.CommonController.WebWriteLog(exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace);
                reJo.msg = "获取文件编号失败！" + exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace;
            }
            return reJo.Value;
        }

        /// <summary>
        /// 创建发文单后，发起发文流程
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="docKeyword"></param>
        /// <param name="DocList"></param>
        /// <returns></returns>
        public static JObject DocumentStartWorkFlow(string sid, string docKeyword, string DocList)
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

                Doc doc = dbsource.GetDocByKeyWord(docKeyword);
                if (doc == null)
                {
                    reJo.msg = "错误的文档操作信息！指定的文档不存在！";
                    return reJo.Value;
                }


            }
            catch (Exception exception)
            {
                WebApi.CommonController.WebWriteLog(exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace);
                //AssistFun.PopUpPrompt(exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace);
                reJo.msg = "启动流程失败！" + exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace;
            }
            return reJo.Value;

        }

        /// <summary>
        /// 处理拖拽文件到DocGrid控件的处理事件
        /// </summary>
        /// <returns></returns>
        public static JObject OnBeforeFileAddEvent(string sid, string ProjectKeyword) {
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

                Project m_Project = dbsource.GetProjectByKeyWord(ProjectKeyword);

                if (m_Project == null)
                {
                    reJo.msg = "参数错误！文件夹不存在！";
                    return reJo.Value;
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

        public static ExReJObject OnAfterCreateNewObject(object obj) {
            ExReJObject reJo = new ExReJObject();
            reJo.success = true;
            try
            {
                //判断是否设置了主设
                //查找设计阶段
                //找设计阶段
                Doc doc = (Doc)obj;
                Project m_Project = doc.Project;


                //收文目录可以使用
                if (m_Project.TempDefn == null || m_Project.TempDefn.KeyWord != "COM_COMTYPE")
                {
                    //当success返回true,msg返回""时，继续上传文件
                    reJo.success = true;
                    return reJo;
                }

                ////放置在函件单位下的分类目录下
                if (m_Project != null && m_Project.TempDefn != null && m_Project.TempDefn.KeyWord == "COM_COMTYPE" 
                    &&  (m_Project.ParentProject.Code == "收文" || m_Project.ParentProject.Description == "收文"))
                {
                    reJo.msg = "RecDocument";
                    reJo.data = new JArray(new JObject(
                        new JProperty("plugins", "HXEPC_Plugins"),
                        new JProperty("FuncName", "recDocument"),
                        new JProperty("DocKeyword", doc.KeyWord),
                        new JProperty("ProjectKeyword", doc.Project.KeyWord)
                        ));

                    //当返回false时，向客户端发送返回，返回为true时，就不向客户端返回
                    reJo.success = false;
                    return reJo;

                }
                reJo.success = true;
                return reJo;

            }
            catch (Exception e)
            {
                reJo.msg = e.Message;
                CommonController.WebWriteLog(reJo.msg);
            }
            return reJo;
        }


    }
}

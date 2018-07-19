namespace AVEVA.CDMS.HXEPC_Plugins
{
    //using AVEVA.CDMS.Common;
    using AVEVA.CDMS.Server;
    //using Microsoft.Win32;
    //using SEAGULL;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Mail;
    using System.Net.Mime;
    using System.Text;
    using System.Threading;
    //using System.Windows.Forms;


    public class CommonFunction
    {
        /// <summary>
        /// 获取项目对象
        /// </summary>
        /// <param name="obj">目录或者文档</param>
        /// <returns></returns>
        public static Project GetProject(object obj)
        {
            if (obj is Doc || obj is Project)
            {
                //或者对象
                Project p = null;
                if (obj is Doc)
                {
                    p = ((Doc)obj).Project;
                }
                else
                {
                    p = (Project)obj;
                }

                //查找
                while (p != null)
                {
                    if (p.TempDefn != null && p.TempDefn.KeyWord == "DESIGNPROJECT")
                    {
                        return p;
                    }
                    p = p.ParentProject;
                }
            }
            return null;
        }

        /// <summary>
        /// 获取阶段对象
        /// </summary>
        /// <param name="obj">目录或者文档</param>
        /// <returns></returns>
        public static Project GetDesign(object obj)
        {
            if (obj is Doc || obj is Project)
            {
                //或者对象
                Project p_Design = null;
                if (obj is Doc)
                {
                    p_Design = ((Doc)obj).Project;
                }
                else
                {
                    p_Design = (Project)obj;
                }


                //查找
                while (p_Design != null)
                {
                    if (p_Design.TempDefn != null && p_Design.TempDefn.KeyWord == "DESIGNPHASE")
                    {
                        return p_Design;
                    }
                    p_Design = p_Design.ParentProject;
                }
            }
            return null;
        }

        /// <summary>
        /// 获取专业对象
        /// </summary>
        /// <param name="obj">目录或者文档</param>
        /// <returns></returns>
        public static Project GetProfession(object obj)
        {
            if (obj is Doc || obj is Project)
            {
                //或者对象
                Project p_Pofession = null;
                if (obj is Doc)
                {
                    p_Pofession = ((Doc)obj).Project;
                }
                else
                {
                    p_Pofession = (Project)obj;
                }


                //查找
                while (p_Pofession != null)
                {
                    if (p_Pofession.TempDefn != null && p_Pofession.TempDefn.KeyWord == "PROFESSION")
                    {
                        return p_Pofession;
                    }
                    p_Pofession = p_Pofession.ParentProject;
                }
            }
            return null;
        }



        /// <summary>
        /// 查找模板
        /// </summary>
        /// <param name="obj">Doc， Project</param>
        /// <param name="keyword">模板关键字</param>
        /// <returns>TempDefn</returns>
        public static TempDefn GetTempDefn(object obj, string keyword)
        {
            if (obj is Doc || obj is Project)
            {
                //或者对象
                Project p = null;
                if (obj is Doc)
                {
                    p = ((Doc)obj).Project;
                }
                else
                {
                    p = (Project)obj;
                }


                //查找
                List<TempDefn> tdlist = p.dBSource.GetTempDefnByCode("keyword");
                if (tdlist != null && tdlist.Count > 0)
                {
                    return tdlist[0];
                }
            }

            return null;
        }


        //发文邮件
        /// <summary>
        /// 发邮件
        /// </summary>
        /// <param name="MailFrom">发件人邮箱地址</param>
        /// <param name="MailToList">收件人邮箱地址</param>
        /// <param name="MailTitle">主题</param>
        /// <param name="MailBody">邮件内容</param>
        /// <param name="FileList">附加文件</param>
        /// <param name="MailServer">服务器地址</param>
        /// <param name="UserName">邮箱登录账号</param>
        /// <param name="UserPw">邮箱登录密码</param>
        /// <returns></returns>
        private static bool SendEmail(string MailFrom, List<string> MailToList, string MailTitle, string MailBody, List<string> FileList, string MailServer, string UserName, string UserPw)
        {
            try
            {
                if ((MailToList == null) || (MailToList.Count < 1))
                {
                    return false;
                }
                MailMessage message = new MailMessage
                {
                    From = new MailAddress(MailFrom.Trim())
                };
                foreach (string str in MailToList)
                {
                    message.To.Add(str);
                }
                message.Subject = MailTitle;
                message.Body = MailBody;
                message.IsBodyHtml = true;
                message.BodyEncoding = Encoding.UTF8;
                message.Headers.Add("Disposition-Notification-To", MailFrom);
                if ((FileList != null) && (FileList.Count > 0))
                {
                    foreach (string str2 in FileList)
                    {
                        Attachment item = new Attachment(str2)
                        {
                            Name = Path.GetFileName(str2),
                            NameEncoding = Encoding.GetEncoding("gb2312"),
                            TransferEncoding = TransferEncoding.Base64
                        };
                        item.ContentDisposition.Inline = true;
                        item.ContentDisposition.DispositionType = "inline";
                        message.Attachments.Add(item);
                    }
                }
                new SmtpClient(MailServer) { Credentials = new NetworkCredential(UserName.Trim(), UserPw.Trim()) }.Send(message);
                return true;
            }
            catch (Exception exception)
            {
                //ErrorLog.WriteErrorLog(exception.ToString());
                WebApi.CommonController.WebWriteLog(exception.ToString());
                return false;
            }
        }

        public static Project GetProjectByDesc(Project curProject,string Desc)
        {
            if (curProject == null) return null;

            Project proj = curProject;
            foreach (Project prj in proj.ChildProjectList) {
               
                if (prj.Description == Desc) {
                    return prj;
                }
            }
            return null;
        }

        ///// <summary>
        ///// 流程里面发邮件
        ///// </summary>
        ///// <param name="wf"></param>
        //public static void SendMail(WorkFlow wf)
        //{
        //    try
        //    {
        //        Doc doc = wf.doc;
        //        string CompCode;
        //        string CompName;
        //        string CompEmail;
        //        string CompRecevier;
        //        string CompFaxNo;
        //        string CompEnclosure;

        //        string mailFrom = "CDMSAdmin@gedi.com.cn";
        //        List<string> mailToList = new List<string>();
        //        string mailTitle = doc.O_itemname;
        //        string mailBody = "";
        //        List<string> FileList = null;
        //        string mailServer = "smtp.gedi.com.cn";
        //        string userName = "CDMSAdmin";
        //        string userPw = "CDMS_12345";

        //        //查找设计阶段
        //        Project p = wf.doc.Project;
        //        while (p != null && p.ParentProject != null)
        //        {
        //            //找设计阶段
        //            if (p != null && p.TempDefn != null && p.TempDefn.KeyWord == "DESIGNPHASE")
        //            {
        //                break;
        //            }
        //            p = p.ParentProject;
        //        }

        //        //查找收文厂家
        //        Project profession = wf.CuWorkState.O_iuser4 != null ? p.dBSource.GetProjectByID((int)wf.CuWorkState.O_iuser4) : null;
        //        if (profession == null)
        //        {
        //            foreach (Project pp in p.ChildProjectList)
        //            {
        //                if (pp.Code == "收文")
        //                {
        //                    foreach (Project cj in pp.ChildProjectList)
        //                    {
        //                        //查找厂家
        //                        Project Comp = doc.Project.ParentProject;
        //                        if (cj.Code == Comp.Code)
        //                        {
        //                            //查找厂家数据
        //                            CompCode = cj.GetAttrDataByKeyWord("FC_COMPANYCODE").ToString;//厂家编码
        //                            CompName = cj.GetAttrDataByKeyWord("FC_COMPANYCHINESE").ToString;//厂家名称
        //                            CompEmail = cj.GetAttrDataByKeyWord("FC_EMAIL").ToString;//邮箱编号
        //                            CompRecevier = cj.GetAttrDataByKeyWord("FC_RECEIVER").ToString;//厂家收件人
        //                            CompFaxNo = cj.GetAttrDataByKeyWord("FC_FAXNO").ToString;//厂家传真号
        //                            CompEnclosure = doc.GetAttrDataByKeyWord("IF_SENDFILE").ToString;//文件附件

        //                            //分割附件
        //                            foreach (string enclosure in CompEnclosure.Split(new char[] { ';' }))
        //                            {
        //                                if (!string.IsNullOrEmpty(enclosure))
        //                                {
        //                                    mailToList.Add(enclosure);
        //                                }
        //                            }
        //                            if (DialogResult.Yes == MessageBox.Show("请核对邮件接收者地址,确认无误后点击确定进行发送(点取消请自行发送邮件)?\r\n接收方:" + CompName + "\r\n接收人:" + CompRecevier + "\r\n接收地址:" + CompEmail, "发送信息", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
        //                            {
        //                                try
        //                                {
        //                                    mailToList.Add(CompEmail);
        //                                    SendEmail(mailFrom, mailToList, mailTitle, mailBody, FileList, mailServer, userName, userPw);
        //                                    //MessageBox.Show(mailFrom + "\n\n发往\n\n" + CompEmail + "\n\n邮件发送成功   \r\n接口文件编号是：" + doc.Code);
        //                                    //ErrorLog.WriteErrorLog(mailFrom + "发往" + CompEmail + "邮件成功   接口文件编号是：" + doc.Code);
        //                                    WebApi.CommonController.WebWriteLog(mailFrom + "发往" + CompEmail + "邮件成功   接口文件编号是：" + doc.Code);
        //                                }
        //                                catch (Exception exception)
        //                                {
        //                                    //MessageBox.Show(mailFrom + "发往" + CompEmail + "邮件失败，请联系管理员。");
        //                                    //ErrorLog.WriteErrorLog(mailFrom + "发往" + CompEmail + "邮件失败" + exception.ToString() + "   接口文件编号是：" + doc.Code);
        //                                    WebApi.CommonController.WebWriteLog(mailFrom + "发往" + CompEmail + "邮件失败" + exception.ToString() + "   接口文件编号是：" + doc.Code);
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch { }
        //}

    
        /// <summary>
        /// 内部提资：专业主设人接收文件签名
        /// </summary>
        /// <param name="workflow"></param>
        public static void ExcSign(WorkFlow wf)
        {
            try
            {
                if (wf != null && wf.DocList != null && wf.DocList.Count > 0 && wf.O_WorkFlowStatus == enWorkFlowStatus.Finish)
                {

                    //取第一个文件
                    Doc doc = wf.DocList[0];

                    //文件全称
                    string str = doc.O_filename.ToUpper();

                    //流程状态不能为空
                    if ((wf.WorkStateList == null) || (wf.WorkStateList.Count <= 1))
                    {
                        return;
                    }

                    ////FTP对象
                    //FTPFactory fTP = null;
                    //if (doc.Storage.FTP != null)
                    //{
                    //    //文档存储位置
                    //    fTP = doc.Storage.FTP;
                    //}
                    //else
                    //{
                    //    fTP = new FTPFactory(doc.Storage);
                    //}
                    //if (fTP == null)
                    //{
                    //    return;
                    //}

                    //获取即将生成的互提单文件路径
                    string locFileName = doc.FullPathFile;

                    //文件路径不为空
                    if (!string.IsNullOrEmpty(locFileName))
                    {
                        if (System.IO.File.Exists(locFileName))
                        {
                            //文件位置
                            FileInfo info = new FileInfo(locFileName);

                            //文件是否只读
                            if (info.IsReadOnly)
                            {
                                info.IsReadOnly = false;
                                //info.Delete();
                            }
                        }

                        //读取
                        //fTP.download(doc.FullPathFile, doc.dBSource.LoginUser.WorkingPath + doc.O_filename, true);
                    }

                    //创建附加表单属性
                    Hashtable htUserKeyWord = new Hashtable();

                    //查找接收人
                    int indx = 1;
                    foreach (WorkState ws in wf.WorkStateList)
                    {
                        if (ws.DefWorkState.O_Code == "RECEIVE")
                        {
                            htUserKeyWord.Add("RECEIVER" + indx.ToString(), ws.WorkUserList[0].User.Code);            //接收用户
                            htUserKeyWord.Add("RENAME" + indx.ToString(), ws.O_suser3);                               //接收专业
                            htUserKeyWord.Add("RETIME" + indx.ToString(), DateTime.Now.ToString("yyyy-MM-dd"));       //接收时间
                            indx++;
                        }

                    }

                    //写表单
                    if (((htUserKeyWord.Count != 0) && str.EndsWith(".DOC")) || ((str.EndsWith(".DOCX") || str.EndsWith(".XLS")) || str.EndsWith(".XLSX")))
                    {
                        WebApi.CDMSWebOffice office = new WebApi.CDMSWebOffice
                        {
                            CloseApp = true,
                            VisibleApp = false
                        };

                        //释放文档
                        office.Release(true);
                        if ((doc.WorkFlow != null) && (doc.WorkFlow.O_WorkFlowStatus == enWorkFlowStatus.Finish))
                        {
                            office.IsFinial = true;
                        }

                        //写入文档
                        office.WriteDataToDocument(doc, locFileName, htUserKeyWord, htUserKeyWord);
                    }

                    ////上传文件
                    //if (fTP != null)
                    //{
                    //    fTP.upload(doc.dBSource.LoginUser.WorkingPath + doc.O_filename, doc.FullPathFile);
                    //    fTP.close();
                    //}
                }

            }
            catch (Exception ex)
            {
                //错误提示
                //AssistFun.PopUpPrompt("发生错误：" + ex.ToString());
                WebApi.CommonController.WebWriteLog("发生错误：" + ex.ToString());
            }
        }

        ////内部提资
        //public static bool InsertDocListAndOpenDoc(List<Doc> dl, Doc doc)
        //{
        //    try
        //    {
        //        if (ExMenu.callTheApp != null)
        //        {
        //            CallBackResult result2;
        //            CallBackParam param = new CallBackParam
        //            {
        //                callType = enCallBackType.UpdateDBSource,
        //                dbs = doc.dBSource
        //            };
        //            CallBackResult result = null;
        //            ExMenu.callTheApp(param, out result);
        //            CallBackParam param2 = new CallBackParam();
        //            if (doc == null)
        //            {
        //                return false;
        //            }
        //            param2.dList = new List<Doc>();
        //            param2.dList.Add(doc);
        //            param2.callType = enCallBackType.DocSimpleOpen;
        //            ExMenu.callTheApp(param2, out result2);
        //        }
        //        return true;
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}

        //线程锁 
        internal static Mutex muxConsole = new Mutex();


        /// <summary>
        /// 发文通信校审流程签名
        /// </summary>
        /// <param name="wf"></param>
        /// 签名公式：<AVEVA.CDMS.HXEPC_Plugins.CommonFunction><HXSign>
        public static void HXSign(WorkFlow wf)
        {
            try
            {
                if (((wf != null) && (wf.DocList != null)) && (wf.DocList.Count > 0))
                {
                    foreach (Doc doc in wf.DocList)
                    {
                        string str = doc.O_filename.ToUpper();
                        Hashtable htUserKeyWord = new Hashtable();
                        if ((wf.WorkStateList == null) || (wf.WorkStateList.Count <= 1))
                        {
                            continue;
                        }
                        int count = wf.WorkStateList.Count;
                        //FTPFactory fTP = null;
                        //if (doc.Storage.FTP != null)
                        //{
                        //    fTP = doc.Storage.FTP;
                        //}
                        //else
                        //{
                        //    fTP = new FTPFactory(doc.Storage);
                        //}
                        //if (fTP == null)
                        //{
                        //    return;
                        //}

                        //获取即将函件的文件路径
                        string locFileName = doc.FullPathFile;

                        if (string.IsNullOrEmpty(locFileName))
                        {
                            return;
                        }
                        if (!System.IO.File.Exists(locFileName))
                        {
                            return;
                        }

                        //if (!string.IsNullOrEmpty(locFileName))
                        //{
                        //    if (System.IO.File.Exists(locFileName))
                        //    {
                        //        FileInfo info = new FileInfo(locFileName);
                        //        if (info.IsReadOnly)
                        //        {
                        //            info.IsReadOnly = false;
                        //            //info.Delete();
                        //        }
                        //    }
                        //    //fTP.download(doc.FullPathFile, doc.dBSource.LoginUser.WorkingPath + doc.O_filename, true);
                        //}

                        string code = wf.CuWorkState.Code;
                        if ((doc.WorkFlow != null) && (doc.WorkFlow.O_WorkFlowStatus == enWorkFlowStatus.Finish))
                        {
                            code = "END";
                        }
                        switch (code)
                        {
                            //校核人
                            case "CHECK":
                                //htUserKeyWord.Add("PREPAREDSIGN1", doc.dBSource.LoginUser.O_username);
                                htUserKeyWord["CHECKPERSON"] = doc.WorkFlow.CuWorkState.WorkUserList[0].User.O_username;
                                htUserKeyWord["CHECKTIME"] = DateTime.Now.ToString("yyyy.MM.dd");
                                goto Label_04DA;

                            //审核人
                            case "AUDIT":
                                htUserKeyWord["AUDITPERSON"] = doc.WorkFlow.CuWorkState.WorkUserList[0].User.O_username;
                                htUserKeyWord["AUDITTIME"] = DateTime.Now.ToString("yyyy.MM.dd");
                                goto Label_04DA;

                            //审定人
                            case "AUDIT2":
                                htUserKeyWord["REVIEWER"] = doc.WorkFlow.CuWorkState.WorkUserList[0].User.O_username;
                                htUserKeyWord["AUDITTIME2"] = DateTime.Now.ToString("yyyy.MM.dd");
                                goto Label_04DA;
                            
                            //批准人
                            case "APPROV":
                                htUserKeyWord["APPROVER"] = doc.WorkFlow.CuWorkState.WorkUserList[0].User.O_username;
                                htUserKeyWord["APPROVETIME"] = DateTime.Now.ToString("yyyy.MM.dd");

                                //检查签名，如果没有签名的画斜线
                                //校核人
                                htUserKeyWord["CHECKPERSON"] = "slash";
                                //审核人
                                htUserKeyWord["AUDITPERSON"] = "slash";
                                //审定人
                                htUserKeyWord["REVIEWER"] = "slash";


                                goto Label_04DA;


                            case "INTERFACE":
                            case "END":
                                {
                                    string str3 = doc.dBSource.ParseExpression(doc, "$(PROJECTOWNER)");
                                    if (wf.DefWorkFlow.O_Code == "HTWORKFLOW")
                                    {
                                      //  htUserKeyWord["APPROVEPERSON"] = doc.dBSource.LoginUser.O_username;
                                    }
                                    else if (!string.IsNullOrEmpty(str3))
                                    {
                                       // htUserKeyWord["APPROVEPERSON"] = str3;
                                       // htUserKeyWord["APPROVETIME"] = DateTime.Now.ToString("yyyy.MM.dd");
                                    }
                                    goto Label_04DA;
                                }
                            default:
                                goto Label_04DA;
                        }
                        if (wf.CuWorkState.PreWorkState.Code == "CHECK")
                        {
                            htUserKeyWord["CHECKPERSON"] = doc.dBSource.LoginUser.O_username;
                        }
                        else
                        {
                            htUserKeyWord["AUDITPERSON"] = doc.dBSource.LoginUser.O_username;
                        }
                        Label_0385:
                        htUserKeyWord["AUDITTIME"] = DateTime.Now.ToString("yyyy.MM.dd");
                        Label_04DA:
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
                            catch (Exception ExOffice) {
                                WebApi.CommonController.WebWriteLog(ExOffice.Message);
                            }
                            finally
                            {

                                //解锁
                                muxConsole.ReleaseMutex();
                            }
                            if (((doc.WorkFlow != null) && (doc.WorkFlow.O_WorkFlowStatus == enWorkFlowStatus.Finish)) && (doc.GetValueByKeyWord("GEDI_INNERIFTYPE") != "提出资料"))
                            {
                                if (wf.DefWorkFlow.O_Code == "INTERFACEWORKFLOW")
                                {
                                    Doc doc2 = doc;
                                    DateTime minValue = DateTime.MinValue;
                                    if ((doc2.Project.DocList != null) && (doc2.Project.DocList.Count > 0))
                                    {
                                        doc2.Project.DocList.Sort(new Comparison<Doc>(CommonFunction.CompareByCode));
                                        Doc doc3 = doc2.Project.DocList.Last<Doc>(dx => dx.TempDefn != null);
                                        if ((doc3 != null) && ((doc3.TempDefn.Code == "GEDI_TRANSFERINGFORM") || (doc3.TempDefn.Code == "GEDI_VIEWFORM")))
                                        {
                                            if ((doc3.WorkFlow != null) && (doc3.WorkFlow.O_WorkFlowStatus == enWorkFlowStatus.Finish))
                                            {
                                                minValue = doc3.WorkFlow.WorkStateList.Last<WorkState>().O_FinishDate.Value;
                                            }
                                        }
                                        else if ((doc3 != null) && (doc3.TempDefn.Code == "IMPORTINTERFACEFILE"))
                                        {
                                            minValue = doc3.O_credatetime;
                                        }
                                    }
                                }
                            }
                            //if (fTP != null)
                            //{
                            //    fTP.upload(doc.dBSource.LoginUser.WorkingPath + doc.O_filename, doc.FullPathFile);
                            //    fTP.close();
                            //}
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                WebApi.CommonController.WebWriteLog(exception.ToString());
                //ErrorLog.WriteErrorLog(exception.ToString());
            }
        }

        //工作任务流程主任选择接收人时，自动创建用户任务
        public static void CreateUserTask(WorkFlow wf) {
            try
            {

                if (wf.CuWorkState.PreWorkState.Code == "DIRECTORSELECT" && (wf.DefWorkFlow.O_Code == "WORKTASK"))
                {
                    Project mProject = wf.Project;
                    if (mProject == null)
                    {
                        return;
                    }

                    ////添加用户任务
                    string strName = mProject.Code;//工作名称
                    string strWORKTEST = mProject.GetValueByKeyWord("WORKTEST");//工作内容
                    string strStartDate = mProject.GetValueByKeyWord("TASKPLANSTARTDATE");//开始时间
                    string strEndDate = mProject.GetValueByKeyWord("TASKPLANFINISHDATE"); //结束时间

                    DateTime dtStart = string.IsNullOrEmpty(strStartDate) ? DateTime.Now : Convert.ToDateTime(strStartDate);
                    DateTime dtEnd = string.IsNullOrEmpty(strEndDate) ? DateTime.Now : Convert.ToDateTime(strEndDate);

                    List<WorkUser> workUserList = wf.CuWorkState.WorkUserList;

                    foreach (WorkUser workUser in workUserList)
                    {
                        if (workUser != null && workUser.User != null)
                        {
                            User m_User = workUser.User;

                            if (m_User != null)
                            {
                                Task newTask = wf.dBSource.NewTask(enTaskLevel.Common, enTaskStatus.Runing, "任务", strName, strWORKTEST, m_User, null, dtStart, dtEnd, 0, null);
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                WebApi.CommonController.WebWriteLog(exception.ToString());
                //ErrorLog.WriteErrorLog(exception.ToString());
            }
        }

        ////
        /// <summary>
        /// 把用户字符串转换成用户列表
        /// </summary>
        /// <param name="dbsource"></param>
        /// <param name="userlist">用户列表，格式："用户代码1__用户名1，用户代码2__用户名2。。。"</param>
        /// <returns></returns>
        public static Group StrToGroup(DBSource dbsource,string userlist) {
            AVEVA.CDMS.Server.Group userGroup=new AVEVA.CDMS.Server.Group ();
            try
            {
                string[] strArray = (string.IsNullOrEmpty(userlist) ? "" : userlist).Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string struser in strArray)
                {
                    Object objuser = dbsource.GetObjectByKeyWord(struser);

                    if (objuser == null) break;
                    if (objuser is User)
                    {
                        userGroup.AddUser(objuser as User);
                        // userList.Add(user);
                    }

                }
            }
            catch { }
            return userGroup;
        }

        /// <summary>
        /// 转换用户字符串，获取只包含用户代码的字符串
        /// </summary>
        /// <param name="userlist">用户列表，格式："用户代码1__用户名1，用户代码2__用户名2。。。"</param>
        /// <returns>
        /// 返回格式："用户代码1，用户代码2。。。"
        /// </returns>
        public static string getUserCodelist(string userlist)
        {
            string strUserCodeList = "";

            string[] strArray = (string.IsNullOrEmpty(userlist) ? "" : userlist).Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string struser in strArray)
            {
                string userCode = struser.Substring(0, struser.IndexOf("__"));
                strUserCodeList = strUserCodeList + userCode+",";
            }

            if (strUserCodeList.EndsWith(","))
            {
                strUserCodeList = strUserCodeList.Remove(strUserCodeList.Length - 1);
            }

            return strUserCodeList;
        }
        public static int CompareByCode(Doc x, Doc y)
        {
            CaseInsensitiveComparer comparer = new CaseInsensitiveComparer();
            return comparer.Compare(x.O_credatetime, y.O_credatetime);
        }

        /// <summary>
        /// 根据模板获取父目录
        /// </summary>
        /// <param name="curProj"></param>
        /// <param name="TempDefnKeyWord"></param>
        /// <returns></returns>
        internal static Project getParentProjectByTempDefn(Project curProj,string TempDefnKeyWord) {
            //Project proj = null;

            #region 获取项目名称
            Project proj = curProj;
            Project rootProj = null;
            //string rootProjDesc = "";
            try
            {
                while (true)
                {
                    if (proj.TempDefn != null && proj.TempDefn.KeyWord == TempDefnKeyWord)
                    {
                        rootProj = proj;
                        //rootProjDesc = proj.Description;
                        break;
                    }
                    else
                    {
                        if (proj.ParentProject == null)
                        {
                            break;
                        }
                        else
                        {
                            proj = proj.ParentProject;
                        }
                    }

                }
            }
            catch (Exception ex){
                WebApi.CommonController.WebWriteLog(DateTime.Now.ToString() + ":" + "根据模板获取项目错误," +ex.Message);

            }
            #endregion
            return rootProj;

        }

        internal static string GetSecrearilManByUnitCode(DBSource dbsource,string UnitCode) {
            string secUserList = "";
            //从组织机构里面查找文控
            Server.Group gp = dbsource.GetGroupByName(UnitCode);
            if (gp == null)
            {
                return "";
            }
            foreach (Server.Group g in gp.AllGroupList)
            {
                if (g.Description == "文控")
                {

                    foreach (User user in g.AllUserList)
                    {
                        secUserList = user.ToString + ",";

                    }
                    if (!string.IsNullOrWhiteSpace(secUserList))
                    {
                        secUserList = secUserList.Substring(0, secUserList.Length - 1);
                    }
                    break;
                }
            }
            return secUserList;
        }

        internal static Group GetSecGroupByUnitCode(DBSource dbsource, string UnitCode)
        {
            string secUserList = "";
            //从组织机构里面查找文控
            Server.Group gp = dbsource.GetGroupByName(UnitCode);
            if (gp == null)
            {
                return null;
            }
            Server.Group resultGp = null;
            foreach (Server.Group g in gp.AllGroupList)
            {
                if (g.Description == "文控")
                {

                    resultGp = g;
                    break;
                }
            }
            return resultGp;
        }

        //获取项目部门的文控
        //internal static User GetSecUserByDepartmentCode(DBSource dbsource, string DepartmentCode) {
        internal static string GetSecUserByDepartmentCode(DBSource dbsource, string DepartmentCode)
        {
            try
            {
                //先再组织机构里面获取文控组
                Group gp = GetSecGroupByUnitCode(dbsource, DepartmentCode);
                if (gp != null)
                {
                    if (gp.AllUserList.Count > 0)
                    {
                       User user= gp.AllUserList[0];
                        return user.ToString;
                    }
                }
                else
                {
                    //如果组织机构里面没有，再从项目目录里面获取文控
                    Project m_RootProject = dbsource.RootLocalProjectList.Find(itemProj => itemProj.TempDefn.KeyWord == "PRODOCUMENTADMIN");
                    if (m_RootProject == null)
                    {
                        return null;
                    }

                    foreach (Project proj in m_RootProject.AllProjectList)
                    {
                        try
                        {
                            if (proj.TempDefn.KeyWord == "HXNY_DOCUMENTSYSTEM")
                            {//&& proj.Code== DepartmentCode) {
                                string strONSHORE = proj.GetAttrDataByKeyWord("RPO_ONSHORE").ToString;//通信代码
                                string strOFFSHORE = proj.GetAttrDataByKeyWord("RPO_OFFSHORE").ToString;//通信代码
                                if (strONSHORE == DepartmentCode || strOFFSHORE == DepartmentCode || proj.Code == DepartmentCode)
                                {
                                    string strUser = proj.GetAttrDataByKeyWord("SECRETARILMAN").ToString;//文件附件
                                    return strUser;
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
            return null;
        }

        internal static Group GetUserRootOrgGroup(User curUser)
        {//获取组织机构用户组
            try
            {
                foreach (AVEVA.CDMS.Server.Group groupOrg in curUser.dBSource.AllGroupList)
                {
                    if ((groupOrg.ParentGroup == null) && (groupOrg.O_grouptype == enGroupType.Organization))
                    {
                        if (groupOrg.AllUserList.Contains(curUser))
                        {
                            return groupOrg;
                        }
                    }
                }
            }
            catch { }
            return null;
        }
    }
}


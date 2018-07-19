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
    public class LetterCN
    {

        /// <summary>
        /// 获取创建信函表单的默认配置
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="ProjectKeyword"></param>
        /// <returns></returns>
        public static JObject GetDraftLetterCNDefault(string sid, string ProjectKeyword,string DraftOnProject)
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

                Project m_Project = dbsource.GetProjectByKeyWord(ProjectKeyword);

                //定位到发文目录
                //m_Project = LocalProject(m_Project);

                if (m_Project == null)
                {
                    reJo.msg = "参数错误！文件夹不存在！";
                    return reJo.Value;
                }

                //获取项目号
                string RootProjectCode = m_Project.GetValueByKeyWord("HXNY_DOCUMENTSYSTEM_CODE");
                if (RootProjectCode == null) RootProjectCode = "";

                //设计阶段目录
                Project designphase = CommonFunction.GetDesign(m_Project);

                //获取发文单位列表
                JObject joSendCompany = new JObject();

                //获取收文单位列表
                JObject joRecCompany = new JObject();

                //Dictionary<string, string> dictionary = new Dictionary<string, string>();
                List<DictData> dictDataList = dbsource.GetDictDataList("Communication");
                //[o_Code]:英文描述,[o_Desc]：中文描述,[o_sValue1]：通信代码
                //string str3 = m_Project.ExcuteDefnExpression("$(DESIGNPROJECT_CODE)")[0];
                foreach (DictData data6 in dictDataList)
                {
                    if (!string.IsNullOrEmpty(data6.O_sValue1))
                    {
                        joRecCompany.Add(new JProperty(data6.O_sValue1,data6.O_Desc ));
                        joSendCompany.Add(new JProperty(data6.O_sValue1, data6.O_Desc));
                    }
                }

                //获取根目录
                Project rootProj=CommonFunction.getParentProjectByTempDefn(m_Project, "HXNY_DOCUMENTSYSTEM");
                string strCompany = "";
                if (rootProj != null)
                {
                    strCompany = rootProj.GetAttrDataByKeyWord("PRO_COMPANY").ToString;
                }


                //string DocNumber = getDocNumber(m_Project, companyList[0].ToString);//设置编号

                string DocNumber = "";// 设置编号

                string groupCode = "", groupKeyword="", groupType="", sourceCompany="";

                if (DraftOnProject == "true")
                {
                    #region 获取项目的通信代码
                    AttrData data;
                    if ((data = m_Project.GetAttrDataByKeyWord("RPO_ONSHORE")) != null)
                    {
                        string strData = data.ToString;
                        if (!string.IsNullOrEmpty(strData))
                        {
                            sourceCompany = data.ToString;
                        }
                    }
                    if (string.IsNullOrEmpty(sourceCompany))
                    { 
                        if ((data = m_Project.GetAttrDataByKeyWord("RPO_OFFSHORE")) != null)
                        {
                            string strData = data.ToString;
                            if (!string.IsNullOrEmpty(strData))
                            {
                                sourceCompany = data.ToString;
                            }
                        }
                    }
                    #endregion
                }

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
                            if (string.IsNullOrEmpty(sourceCompany))
                            {
                                sourceCompany = groupOrg.Code;
                            }
                            break;
                        }
                    }
                }

  

                        JObject joData = new JObject(
                    new JProperty("RootProjectCode", RootProjectCode),
                    new JProperty("DocNumber", DocNumber),
                    new JProperty("RecCompanyList", joRecCompany),
                    new JProperty("SendCompanyList", joSendCompany),
                    new JProperty("SourceCompany", sourceCompany),//groupCode),//strCompany)
                     new JProperty("GroupKeyword", groupKeyword),
                    new JProperty("GroupType", groupType)
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

        /// <summary>
        /// 获取回复信函表单的默认配置
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="DocKeyword"></param>
        /// <returns></returns>
        public static JObject GetReplyLetterCNDefault(string sid, string DocKeyword)
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

                //Project pPrj = m_Doc.Project;
                Project recUnitProj = m_Doc.Project;

                //
                string recCode = m_Doc.Code.Split(new char[] { ' ' })[0];

                Project pcProj=CommonFunction.getParentProjectByTempDefn(m_Doc.Project, "PRO_COMMUNICATION");

                Project m_Project = null;// pcProj.GetProjectByName("发文").GetProjectByName("信函");

                Project sendProj = pcProj.GetProjectByName("发文");

                if (sendProj != null) {
                    m_Project= sendProj.GetProjectByName("信函");
                }

                if (m_Project == null) {
                    // m_Project = pcProj.GetProjectByName("发文").GetProjectByName("信函");
                    sendProj = CommonFunction.GetProjectByDesc(pcProj, "发文");
                    if (sendProj != null) {
                        m_Project = CommonFunction.GetProjectByDesc(sendProj, "信函");

                        if (m_Project == null)
                        {
                            m_Project = sendProj.GetProjectByName("信函");
                        }
                    }


                }

                    //定位到发文目录
                    //m_Project = LocalProject(m_Project);

                if (m_Project == null)
                {
                    reJo.msg = "参数错误！文件夹不存在！";
                    return reJo.Value;
                }

                //获取项目号
                string RootProjectCode = m_Project.GetValueByKeyWord("HXNY_DOCUMENTSYSTEM_CODE");
                if (RootProjectCode == null) RootProjectCode = "";

                //设计阶段目录
                Project designphase = CommonFunction.GetDesign(m_Project);

                //获取发文单位列表
                JObject joSendCompany = new JObject();

                //获取收文单位列表
                JObject joRecCompany = new JObject();

                //Dictionary<string, string> dictionary = new Dictionary<string, string>();
                List<DictData> dictDataList = dbsource.GetDictDataList("Communication");
                //[o_Code]:英文描述,[o_Desc]：中文描述,[o_sValue1]：通信代码
                //string str3 = m_Project.ExcuteDefnExpression("$(DESIGNPROJECT_CODE)")[0];
                foreach (DictData data6 in dictDataList)
                {
                    if (!string.IsNullOrEmpty(data6.O_sValue1))
                    {
                        joRecCompany.Add(new JProperty(data6.O_sValue1, data6.O_Desc));
                        joSendCompany.Add(new JProperty(data6.O_sValue1, data6.O_Desc));
                    }
                }

                //获取根目录
                Project rootProj = CommonFunction.getParentProjectByTempDefn(m_Project, "HXNY_DOCUMENTSYSTEM");
                string strCompany = "";
                if (rootProj != null)
                {
                    strCompany = rootProj.GetAttrDataByKeyWord("PRO_COMPANY").ToString;
                }


                //string DocNumber = getDocNumber(m_Project, companyList[0].ToString);//设置编号

                string DocNumber = "";// 设置编号

                JObject joData = new JObject(
                    new JProperty("RootProjectCode", RootProjectCode),
                    new JProperty("DocNumber", DocNumber),
                    new JProperty("RecCompanyList", joRecCompany),
                    new JProperty("SendCompanyList", joSendCompany),
                    new JProperty("SourceCompany", strCompany),
                    new JProperty("SendProjectKeyword", m_Project.KeyWord),
                    new JProperty("RecUnitCode", ""),//recUnitProj.Code), //设置收文单位
                    new JProperty("RecUnitDesc", ""),//recUnitProj.Description),
                    new JProperty("RecCode", recCode)//收文编码

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

        //线程锁 
        internal static Mutex muxConsole = new Mutex();
        public static JObject DraftLetterCN(string sid, string ProjectKeyword, string DocAttrJson, string FileListJson)
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


                Project m_Project = dbsource.GetProjectByKeyWord(ProjectKeyword);

                //定位到发文目录
                //m_Project = LocalProject(m_Project);

                if (m_Project == null)
                {
                    reJo.msg = "参数错误！文件夹不存在！";
                    return reJo.Value;
                }

                #region 获取信函参数内容

                //获取信函参数内容
                string fileCode = "", mainFeeder = "", copyParty = "", sender = "",
                   sendCode = "", recCode = "", totalPages = "",
                   urgency = "", sendDate = "", seculevel = "",
                   secrTerm = "", needreply = "", replyDate = "",
                   title = "", content = "", approvpath = "",
                   nextStateUserList = "", senderCode = "", mainFeederCode= "",
                   copyCode = "", fileId = "";

                JArray jaAttr = (JArray)JsonConvert.DeserializeObject(DocAttrJson);

                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    ////获取函件编号
                    //if (strName == "documentCode") documentCode = strValue.Trim();
                    //获取文件编码 
                    if (strName == "fileCode") fileCode = strValue.Trim();

                    //获取主送
                    else if (strName == "mainFeeder") mainFeeder = strValue.Trim();

                    //获取抄送
                    else if (strName == "copyParty") copyParty = strValue;

                    //获取发送方
                    else if (strName == "sender") sender = strValue;

                    //获取发文编码
                    else if (strName == "sendCode") sendCode = strValue;

                    //获取收文编码
                    else if (strName == "recCode") recCode = strValue;

                    //获取页数
                    else if (strName == "totalPages") totalPages = strValue;

                    //获取紧急程度
                    else if (strName == "urgency") urgency = strValue;

                    //获取发送日期
                    else if (strName == "sendDate") sendDate = strValue;

                    //获取保密等级
                    else if (strName == "seculevel") seculevel = strValue;

                    //获取保密期限
                    else if (strName == "secrTerm") secrTerm = strValue;

                    //获取是否需要回复
                    else if (strName == "needreply") needreply = strValue;

                    //获取回文日期
                    else if (strName == "replyDate") replyDate = strValue;

                    //获取标题
                    else if (strName == "title") title = strValue;

                    //获取正文内容
                    else if (strName == "content") content = strValue;

                    //获取审批路径
                    else if (strName == "approvpath") approvpath = strValue;

                    //获取下一状态人员
                    else if (strName == "nextStateUserList") nextStateUserList = strValue;

                    //获取发送方代码
                    else if (strName == "senderCode") senderCode = strValue;

                    //获取主送方代码
                    else if (strName == "mainFeederCode") mainFeederCode = strValue;

                    //获取抄送方代码
                    else if (strName == "copyCode") copyCode = strValue;

                    //获取文件ID
                    else if (strName == "fileId") fileId = strValue;

                }

                if (string.IsNullOrEmpty(fileCode))
                {
                    reJo.msg = "请填写文件编号！";
                    return reJo.Value;
                }
                if (string.IsNullOrEmpty(sendCode))
                {
                    reJo.msg = "请填写函件编号！";
                    return reJo.Value;
                }
                else if (string.IsNullOrEmpty(title))
                {
                    reJo.msg = "请填写函件标题！";
                    return reJo.Value;
                }
                else if (string.IsNullOrEmpty(nextStateUserList))
                {
                    reJo.msg = "请选择校审人员！";
                    return reJo.Value;
                }
                if (string.IsNullOrEmpty(fileId))
                {
                    reJo.msg = "请填写文件ID！";
                    return reJo.Value;
                }


                #endregion

                #region 获取文件列表
                List<LetterAttaFile> attaFiles = new List<LetterAttaFile>();
                if (!string.IsNullOrEmpty(FileListJson))
                {
                    int index = 0;
                    JArray jaFiles = (JArray)JsonConvert.DeserializeObject(FileListJson);

                    foreach (JObject joAttr in jaFiles)
                    {
                        string strFileName = joAttr["fn"].ToString();//文件名
                        string strFileCode = joAttr["fc"].ToString();//文件编码
                        string strFileDesc = joAttr["fd"].ToString();//文件题名
                        string strFilePage = joAttr["fp"].ToString();//页数
                        string strEdition = joAttr["fe"].ToString();//版次
                        string strSeculevel = joAttr["sl"].ToString();//密级
                        string strFileState = joAttr["fs"].ToString();//状态
                        string strRemark = joAttr["fr"].ToString();//状态

                        index++;
                        string strIndex = index.ToString();
                        LetterAttaFile afItem = new LetterAttaFile()
                        {
                            No = strIndex,
                            Name = strFileName,
                            Code = strFileCode,
                            Desc = strFileDesc,
                            Page = strFilePage,
                            Edition = strEdition,
                            Seculevel = strSeculevel,
                            Status = strFileState,
                            Remark = strRemark
                        };

                        attaFiles.Add(afItem);
                    }
                }
                #endregion


                #region 根据信函模板，生成信函文档

                //获取立项单文档所在的目录
                //Project m_Project = m_NewProject;

                List<TempDefn> docTempDefnByCode = m_Project.dBSource.GetTempDefnByCode("CATALOGUING");
                TempDefn docTempDefn = (docTempDefnByCode != null && docTempDefnByCode.Count > 0) ? docTempDefnByCode[0] : null;
                if (docTempDefn == null)
                {
                    reJo.msg = "没有与其相关的模板管理，创建无法正常完成";
                    return reJo.Value;
                }

                IEnumerable<string> source = from docx in m_Project.DocList select docx.Code;
                //string filename = sendCode + " " + title;
                string filename = sendCode + " " + title;

                if (source.Contains<string>(filename))
                {
                    for (int i = 1; i < 0x3e8; i++)
                    {
                        filename = sendCode + " " + title + i.ToString();
                        if (!source.Contains<string>(filename))
                        {
                            //reJo.msg = "新建信函出错！";
                            //return reJo.Value;
                            break;
                        }
                    }
                }

                //文档名称
                Doc docItem = m_Project.NewDoc(filename + ".docx", filename, "", docTempDefn);
                if (docItem == null)
                {
                    reJo.msg = "新建信函出错！";
                    return reJo.Value;
                }

                #endregion

                #region 设置信函文档附加属性


                AttrData data;

                //文文档模板
                if ((data = docItem.GetAttrDataByKeyWord("CA_ATTRTEMP")) != null)
                {
                    data.SetCodeDesc("LETTERFILE");
                }

                //文件编码
                if ((data = docItem.GetAttrDataByKeyWord("CA_FILECODE")) != null)
                {
                    data.SetCodeDesc(fileCode);
                }

                //主送
                if ((data = docItem.GetAttrDataByKeyWord("CA_MAINFEEDER")) != null)
                {
                    data.SetCodeDesc(mainFeeder);
                }
                //发送方
                if ((data = docItem.GetAttrDataByKeyWord("CA_SENDER")) != null)
                {
                    data.SetCodeDesc(sender);
                }
                //发文编码
                if ((data = docItem.GetAttrDataByKeyWord("CA_SENDCODE")) != null)
                {
                    data.SetCodeDesc(sendCode);
                }
                //收文编码
                if ((data = docItem.GetAttrDataByKeyWord("CA_RECEIPTCODE")) != null)
                {
                    data.SetCodeDesc(recCode);
                }

                //抄送
                if ((data = docItem.GetAttrDataByKeyWord("CA_COPY")) != null)
                {
                    data.SetCodeDesc(copyParty);
                }
                //页数
                if ((data = docItem.GetAttrDataByKeyWord("CA_PAGE")) != null)
                {
                    data.SetCodeDesc(totalPages);
                }
                //紧急程度
                if ((data = docItem.GetAttrDataByKeyWord("CA_URGENTDEGREE")) != null)
                {
                    data.SetCodeDesc(urgency);
                }
                //发送日期
                if ((data = docItem.GetAttrDataByKeyWord("CA_SENDDATE")) != null)
                {
                    data.SetCodeDesc(sendDate);
                }
                //保密等级
                if ((data = docItem.GetAttrDataByKeyWord("CA_SECRETGRADE")) != null)
                {
                    data.SetCodeDesc(seculevel);
                }
                //保密期限
                if ((data = docItem.GetAttrDataByKeyWord("CA_SECRETTERM")) != null)
                {
                    data.SetCodeDesc(secrTerm);
                }
                //获取是否需要回复
                if ((data = docItem.GetAttrDataByKeyWord("CA_IFREPLY")) != null)
                {
                    data.SetCodeDesc(needreply);
                }

                //回复日期
                if ((data = docItem.GetAttrDataByKeyWord("CA_REPLYDATE")) != null)
                {
                    data.SetCodeDesc(replyDate);
                }
                //标题
                if ((data = docItem.GetAttrDataByKeyWord("CA_TITLE")) != null)
                {
                    data.SetCodeDesc(title);
                }

                //文件列表
                if ((data = docItem.GetAttrDataByKeyWord("CA_ENCLOSURE")) != null)
                {
                    data.SetCodeDesc(FileListJson);
                }

                //校审级数（审批路径）
                if ((data = docItem.GetAttrDataByKeyWord("CA_SERIES")) != null)
                {
                    data.SetCodeDesc(approvpath);
                }

                //发送方代码
                if ((data = docItem.GetAttrDataByKeyWord("CA_SENDERCODE")) != null)
                {
                    data.SetCodeDesc(senderCode);
                }

                //主送方代码
                if ((data = docItem.GetAttrDataByKeyWord("CA_MAINFEEDERCODE")) != null)
                {
                    data.SetCodeDesc(mainFeederCode);
                }

                //抄送方代码
                if ((data = docItem.GetAttrDataByKeyWord("CA_COPYCODE")) != null)
                {
                    data.SetCodeDesc(copyCode);
                }

                //文件ID
                if ((data = docItem.GetAttrDataByKeyWord("CA_FILEID")) != null)
                {
                    data.SetCodeDesc(fileId);
                }
                ////保存项目属性，存进数据库
                docItem.AttrDataList.SaveData();

                #endregion

                #region 录入数据进入word表单

                string strDocList = "";//获取附件

                //录入数据进入表单
                Hashtable htUserKeyWord = new Hashtable();

                #region 添加勾选项

                #region 初始化勾选选项注释
                //////////////////紧急程度///////////////////////
                string URGENCY_NORMAL = "□";//一般
                string URGENCY_URGENT = "□";//紧急

                ////////////////保密等级////////////////////////
                string SECULEVEL_SECRET = "□";//商业秘密
                string SECULEVEL_LIMIT = "□";//受限
                string SECULEVEL_PUBLIC = "□";//公开

                string SECULEVEL_SECRET_TERM = "";//保密期限
                string SECULEVEL_LIMIT_TERM = "";//保密期限

                ///////////////是否需要回复/////////////////////
                string NEEDREPLY_YES = "□";//是
                string NEEDREPLY_NO = "□";//否

                #endregion

                #region 勾选选项
                if (urgency == "一般")
                {
                    URGENCY_NORMAL = "☑";

                }
                else if (urgency == "紧急")
                {
                    URGENCY_URGENT = "☑";

                }


                //勾选保密等级
                if (seculevel == "商业秘密")
                {
                    SECULEVEL_SECRET = "☑";
                    SECULEVEL_SECRET_TERM = secrTerm;
                }
                else if (seculevel == "受限")
                {
                    SECULEVEL_LIMIT = "☑";
                    SECULEVEL_LIMIT_TERM = secrTerm;
                }
                else if (seculevel == "公开")
                {
                    SECULEVEL_PUBLIC = "☑";
                }

                //勾选保密等级
                if (needreply == "是")
                {
                    NEEDREPLY_YES = "☑";
                }
                else if (needreply == "否")
                {
                    NEEDREPLY_NO = "☑";
                }


                #endregion

                #region 添加勾选项到哈希表
                htUserKeyWord.Add("URGENCY_NORMAL", URGENCY_NORMAL);//一般
                htUserKeyWord.Add("URGENCY_URGENT", URGENCY_URGENT);//紧急

                ////////////////保密等级////////////////////////
                htUserKeyWord.Add("SECULEVEL_SECRET", SECULEVEL_SECRET);//商业秘密
                htUserKeyWord.Add("SECULEVEL_LIMIT", SECULEVEL_LIMIT);//受限
                htUserKeyWord.Add("SECULEVEL_PUBLIC", SECULEVEL_PUBLIC);//公开

                htUserKeyWord.Add("SECULEVEL_SECRET_TERM", SECULEVEL_SECRET_TERM); //商业秘密保密期限
                htUserKeyWord.Add("SECULEVEL_LIMIT_TERM", SECULEVEL_LIMIT_TERM); //受限保密期限

                ///////////////是否需要回复/////////////////////
                htUserKeyWord.Add("NEEDREPLY_YES", NEEDREPLY_YES);//是
                htUserKeyWord.Add("NEEDREPLY_NO", NEEDREPLY_NO);//否



                #endregion
                #endregion

                //格式化日期
                DateTime senddate = Convert.ToDateTime(sendDate);
                DateTime replydate = Convert.ToDateTime(replyDate);
                string strSenddate = senddate.ToShortDateString().ToString().Replace("-", ".").Replace("/", ".");
                string strReplydate = replydate.ToShortDateString().ToString().Replace("-", ".").Replace("/", ".");

                


                htUserKeyWord.Add("MAINFEEDER", mainFeeder);//主送
                htUserKeyWord.Add("SENDER", sender);//发送方
                //htUserKeyWord.Add("SENDCODE", sendCode);//发文编码
                htUserKeyWord.Add("RECEIPTCODE", recCode);//收文编码
                htUserKeyWord.Add("COPY", copyParty);//抄送
                htUserKeyWord.Add("PAGE", totalPages);//页数
                htUserKeyWord.Add("SENDDATE", strSenddate);//发送日期
                htUserKeyWord.Add("REPLYDATE", strReplydate);//回复日期
                htUserKeyWord.Add("TITLE", title);//标题
                htUserKeyWord.Add("CONTENT", content);//信函正文

                //htUserKeyWord.Add("RHEADER", fileCode);//文件编码//页眉里面的编码

                //htUserKeyWord.Add("PREPAREDSIGN1", doc.dBSource.LoginUser.O_username);
                //电子签
                htUserKeyWord["PREPAREDSIGN"] = curUser.O_username;
                htUserKeyWord["DRAFTTIME"] = DateTime.Now.ToString("yyyy.MM.dd");


                #region 获取项目名称
                Project proj = docItem.Project;
                //Project rootProj = new Project();
                string rootProjDesc = "";
                Project rootProj = CommonFunction.getParentProjectByTempDefn(proj, "HXNY_DOCUMENTSYSTEM");
                //while (true)
                //{
                //    if (proj.TempDefn!=null && proj.TempDefn.KeyWord == "HXNY_DOCUMENTSYSTEM")
                //    {
                //        //rootProj = proj;
                //        rootProjDesc = proj.Description;
                //        break;
                //    }
                //    else
                //    {
                //        if (proj.ParentProject == null)
                //        {
                //            break;
                //        }
                //        else
                //        {
                //            proj = proj.ParentProject;
                //        }
                //    }

                //}
                #endregion

                string docClass = "";
                //项目管理类地址
                if (rootProj != null)
                {
                    docClass = "project";
                    string proAddress = rootProj.GetValueByKeyWord("PRO_ADDRESS").ToString();
                    string proTel = rootProj.GetValueByKeyWord("PRO_NUMBER").ToString();

                    htUserKeyWord.Add("PRO_ADDRESS", proAddress);//项目地址
                    htUserKeyWord.Add("PRO_TEL", proTel);//项目电话

                    htUserKeyWord["PROJECTDESC"] = "（" + rootProj.Description + "项目部）";
                }

                //运营管理类地址
                else if (rootProj == null)
                {
                    rootProj = CommonFunction.getParentProjectByTempDefn(proj, "OPERATEADMIN");
                    if (rootProj != null)
                    {
                        docClass = "operation";
                        string proAddress = rootProj.GetValueByKeyWord("OPE_ADDRESS").ToString();
                        string proTel = rootProj.GetValueByKeyWord("OPE_NUMBER").ToString();

                        htUserKeyWord.Add("PRO_ADDRESS", proAddress);//项目地址
                        htUserKeyWord.Add("PRO_TEL", proTel);//项目电话

                        //htUserKeyWord["PROJECTDESC"] = "（" + rootProj.Description + "项目部）";
                    }
                }

                //添加附件
                List<string> list3 = new List<string>();
                foreach (LetterAttaFile file in attaFiles)
                {

                    //string remark = string.IsNullOrEmpty(file.Seculevel) ? "" : (file.Seculevel+",");
                    //remark = remark+(string.IsNullOrEmpty(file.Status) ? "" : (file.Status + ","));
                    //remark = remark + (string.IsNullOrEmpty(file.Remark) ? "" : (file.Remark + ","));

                    //if (remark.Substring(remark.Length - 1, 1) == ","&& (!string.IsNullOrEmpty(remark))) {
                    //    remark=remark.Substring(0, remark.Length-1);
                    //}

                    list3.Add(file.No);
                    list3.Add(file.Code);
                    list3.Add(file.Desc);
                    list3.Add(file.Page);
                    list3.Add(file.Edition);
                    list3.Add(file.Seculevel);
                    list3.Add(file.Status);
                    list3.Add(file.Remark);
                }

                //用htAuditDataList传送附件列表到word
                Hashtable htAuditDataList = new Hashtable();
                //word里面表格关键字的设置公式(不需要加"$()") ：表格关键字+":"+已画好表格线的行数+":"+表格列数
                //例如关键字是"DRAWING",画了一行表格线，从第二行起画表格线,每行有6列，则公式是："DRAWING:1:6"
                htAuditDataList.Add("DRAWING", list3);

                //attrDataByKeyWord = mShortDoc.GetAttrDataByKeyWord("IFR_NOTE");
                //if ((attrDataByKeyWord != null) && (this.m_doc != null))
                //{
                //    List<Doc> list4 = this.m_doc.WorkFlow.DocList.Distinct<Doc>().ToList<Doc>();
                //    string docids = "";
                //    list4.ForEach((Action<Doc>)(d => docids = docids + d.O_itemno + ","));
                //    docids = docids.TrimEnd(new char[] { ',' });
                //    attrDataByKeyWord.SetCodeDesc(docids);
                //    mShortDoc.AttrDataList.SaveData();
                //}

                string workingPath = m_Project.dBSource.LoginUser.WorkingPath;


                try
                {
                    //上传下载文档
                    string exchangfilename = "信函中文模板";

                    if (attaFiles.Count<=0)
                        exchangfilename = "信函中文模板无附件";

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
                            office.WriteDataToDocument(docItem, locFileName, htUserKeyWord, htAuditDataList);
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


                    if (string.IsNullOrEmpty(strDocList))
                    {
                        strDocList = docItem.KeyWord;
                    }
                    else
                    {
                        strDocList = docItem.KeyWord + "," + strDocList;
                    }

                    //这里刷新数据源，否则创建流程的时候获取不了专业字符串
                    DBSourceController.RefreshDBSource(sid);

                    reJo.success = true;
                    reJo.data = new JArray(new JObject(new JProperty("ProjectKeyword", docItem.Project.KeyWord),
                        new JProperty("DocKeyword", docItem.KeyWord), new JProperty("DocList", strDocList),
                        new JProperty("DocCode", docItem.Code)));
                    return reJo.Value;
                }
                catch { }
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
        /// 创建发文单后，发起发文流程
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="docKeyword"></param>
        /// <param name="DocList"></param>
        /// <returns></returns>
        public static JObject LetterStartWorkFlow(string sid, string docKeyword, string DocList,
            string ApprovPath,string UserList, string SendUnitCode)
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

                #region 获取项目文控
                Doc doc = dbsource.GetDocByKeyWord(docKeyword);
                if (doc == null)
                {
                    reJo.msg = "错误的文档操作信息！指定的文档不存在！";
                    return reJo.Value;
                }

                string secUserList = "";

                //记录是运营管理类还是项目管理类
                bool isProjectDoc = true;
                Project rootProj = CommonFunction.getParentProjectByTempDefn(doc.Project, "HXNY_DOCUMENTSYSTEM");
                if (rootProj == null)
                {
                    //运营管理类
                    //reJo.msg = "获取项目根目录失败!";
                    //return reJo.Value;

                    isProjectDoc = false;

                    //获取部门文控
                    if (string.IsNullOrEmpty(SendUnitCode))
                    {
                        
                        reJo.msg = "部门文控未设置！";
                        return reJo.Value;
                    }
                    List<DictData> dictDataList = dbsource.GetDictDataList("Communication");
                    //[o_Code]:英文描述,[o_Desc]：中文描述,[o_sValue1]：通信代码
                    //string str3 = m_Project.ExcuteDefnExpression("$(DESIGNPROJECT_CODE)")[0];
                    foreach (DictData data6 in dictDataList)
                    {
                        if (!string.IsNullOrEmpty(data6.O_sValue1) && data6.O_sValue1 == SendUnitCode)
                        {
                            secUserList = data6.O_sValue4;
                            break;
                        }
                    }

                    if (string.IsNullOrEmpty(secUserList))
                    {
                        //从组织机构里面查找文控
                        Server.Group gp= dbsource.GetGroupByName(SendUnitCode);
                        if (gp == null)
                        {
                            reJo.msg = "部门文控未设置！";
                            return reJo.Value;
                        }
                        foreach (Server.Group g in gp.AllGroupList) {
                            if (g.Description == "文控") {
                                
                                foreach (User user in g.AllUserList)
                                {
                                    secUserList = user.ToString + ",";
                                    
                                }
                                if (!string.IsNullOrWhiteSpace(secUserList)) {
                                    secUserList = secUserList.Substring(0, secUserList.Length-1);
                                }
                                break;
                            }
                        }
                        if (string.IsNullOrEmpty(secUserList))
                        {
                            reJo.msg = "部门文控未设置！";
                            return reJo.Value;
                        }
                    }
                }
                else
                {

                    AttrData secData;
                    if ((secData = rootProj.GetAttrDataByKeyWord("SECRETARILMAN")) != null)
                    {
                        secUserList = secData.ToString;
                    }

                    if (string.IsNullOrEmpty(secUserList))
                    {
                        reJo.msg = "项目文控未设置！";
                        return reJo.Value;


                    }

                }


                if (string.IsNullOrEmpty(secUserList)) {
                    reJo.msg = "文控未设置！";
                    return reJo.Value;
                }

                string[] secUserArray = (string.IsNullOrEmpty(secUserList) ? "" : secUserList).Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                Server.Group secGroup = new Server.Group();


                foreach (string strObj in secUserArray)
                {
                    string strUser = strObj.IndexOf("__") >= 0 ? strObj.Substring(0, strObj.IndexOf("__")) : strObj;

                    object obj = dbsource.GetUserByName(strUser);

                    if (obj is User)
                    {
                        //m_UserList.Add((User)obj);
                        secGroup.AddUser((User)obj);
                    }
                }
                if (secGroup.UserList.Count <= 0)
                {
                    reJo.msg = "获取项目文控错误，自动启动流程失败！请手动启动流程";
                    return reJo.Value;
                }
                #endregion

                #region 获取下一状态用户
                string[] userArray = (string.IsNullOrEmpty(UserList) ? "" : UserList).Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                Server.Group group = new Server.Group();
                //List<User> m_UserList = new List<User>();
                //启动工作流程
                //反转列表
                //m_UserList.Reverse();
                foreach (string strObj in userArray)
                {
                    object obj = dbsource.GetObjectByKeyWord(strObj);

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

                //if (MessageBox.Show("是否启动校审流程？", "工作流启动", MessageBoxButtons.YesNo) != DialogResult.No)
                //{
                //    Thread.Sleep(300);
                //    if (((doc.OperateDocStatus == enDocStatus.OUT) && (doc.FLocker == doc.dBSource.LoginUser)) || (doc.OperateDocStatus == enDocStatus.COMING_IN))
                //    {
                //        MessageBox.Show("文档处于检出状态，请先保存文档，并关闭该文档");
                //        this.StartWorkFlow(doc, defWFCode);
                //    }
                //     else
                {


                    #region 获取文档列表
                    string[] strArray = (string.IsNullOrEmpty(DocList) ? "" : DocList).Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    List<Doc> m_DocList = new List<Doc>();
                    //启动工作流程
                    m_DocList.Reverse();
                    foreach (string strObj in strArray)
                    {
                        object obj = dbsource.GetObjectByKeyWord(strObj);

                        if (obj is Doc)
                        {
                            m_DocList.Add((Doc)obj);
                        }
                    }
                    #endregion

                    WorkFlow flow = dbsource.NewWorkFlow(m_DocList, "COMMUNICATIONWORKFLOW");
                    //if (flow == null || flow.CuWorkState == null || flow.CuWorkState.workStateBranchList == null || (flow.CuWorkState.workStateBranchList.Count <= 0))
                    if (flow == null)
                    {
                        reJo.msg = "自动启动流程失败!请手动启动";
                        return reJo.Value;
                    }



                    //获取下一状态
                    //me.approvpathdata = [{ text: "二级-编批", value: "二级-编批" }, { text: "三级-编审批", value: "三级-编审批" },
                    //{ text: "四级-编审定批", value: "四级-编审定批" }, { text: "五级-编校审定批", value: "五级-编校审定批" }];
                    string strCheckStateCode = "";

                    WorkState ws = new WorkState();
                    if (ApprovPath == "二级-编批")
                    {
                        strCheckStateCode = "APPROV";
                    }
                    else if (ApprovPath == "三级-编审批" || ApprovPath == "四级-编审定批")
                    {
                        strCheckStateCode = "AUDIT";
                    }
                    else if (ApprovPath == "五级-编校审定批")
                    {
                        strCheckStateCode = "CHECK";

                    }
                    else {
                        flow.Delete();
                        flow.Delete();

                        reJo.msg = "审批路径参数错误，自动启动流程失败！请手动启动流程";
                        return reJo.Value;
                    }

                    //放置校核状态人员
                    WorkState state = flow.WorkStateList.Find(wsx => (wsx.Code == strCheckStateCode) && (wsx.CheckGroup.AllUserList.Count == 0));
                    if (state == null)
                    {
                        DefWorkState defWorkState = flow.DefWorkFlow.DefWorkStateList.Find(dwsx => dwsx.O_Code == strCheckStateCode);
                        state = flow.NewWorkState(defWorkState);
                        state.SaveSelectUser(group);

                        state.IsRuning = false;

                        state.PreWorkState = flow.CuWorkState;
                        state.O_iuser5 = new int?(flow.CuWorkState.O_stateno);
                        state.Modify();
                    }

                    //放置第二个文控状态人员
                    state = flow.WorkStateList.Find(wsx => (wsx.Code == "SECRETARILMAN") && (wsx.CheckGroup.AllUserList.Count == 0));
                    if (state == null)
                    {
                        DefWorkState defWorkState = flow.DefWorkFlow.DefWorkStateList.Find(dwsx => dwsx.O_Code == "SECRETARILMAN");
                        state = flow.NewWorkState(defWorkState);
                        state.SaveSelectUser(secGroup);

                        state.IsRuning = false;

                        state.PreWorkState = flow.CuWorkState;
                        state.O_iuser5 = new int?(flow.CuWorkState.O_stateno);
                        state.Modify();
                    }

                    //如果是运营管理类
                    if (isProjectDoc == false) {
                        AttrData data;
                        //主送方代码
                        if ((data = doc.GetAttrDataByKeyWord("CA_MAINFEEDERCODE")) != null)
                        {
                            string recerCode=data.ToString;

                            //Server.Group recGroup = dbsource.GetGroupByName(recerCode);
                            //根据部门编码获取文控
                            Server.Group recGroup = CommonFunction.GetSecGroupByUnitCode(dbsource, recerCode);

                            if (recGroup != null)
                            {
                                //放置收文文控状态人员
                                state = flow.WorkStateList.Find(wsx => (wsx.Code == "RECUNIT") && (wsx.CheckGroup.AllUserList.Count == 0));
                                if (state == null)
                                {
                                    DefWorkState defWorkState = flow.DefWorkFlow.DefWorkStateList.Find(dwsx => dwsx.O_Code == "RECUNIT");
                                    state = flow.NewWorkState(defWorkState);
                                    state.SaveSelectUser(recGroup);

                                    state.IsRuning = false;

                                    state.PreWorkState = flow.CuWorkState;
                                    state.O_iuser5 = new int?(flow.CuWorkState.O_stateno);
                                    state.Modify();
                                }
                            }
                        }


                    }

                    //foreach (User user in group.UserList)
                    //{
                    //    ws.group.AddUser(user);
                    //}
                    //flow.WorkStateList.Add(ws);
                    //flow.Modify();

                    ////查找主任
                    //AttrData ad = doc.GetAttrDataByKeyWord("PROFESSIONMANAGER");
                    //if (ad == null || ad.group == null || ad.group.UserList.Count <= 0)
                    //{
                    //    //AssistFun.PopUpPrompt("本专业没有设置主设，不能启动流程，请设置主设后再启动流程！");
                    //    reJo.msg = "本专业没有设置主设，不能启动流程，请设置主设后再启动流程！";
                    //    return reJo.Value;
                    //}

                    ////启动流程
                    WorkStateBranch branch = flow.CuWorkState.workStateBranchList[0];
                    branch.NextStateAddGroup(secGroup);

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

                    return GotoNextReJo.Value;
                    //if (ExMenu.callTheApp != null)
                    //{
                    //    CallBackParam param = new CallBackParam
                    //    {
                    //        callType = enCallBackType.UpdateDBSource,
                    //        dbs = flow.dBSource
                    //    };
                    //    CallBackResult result = null;
                    //    ExMenu.callTheApp(param, out result);
                    //}
                }
                //   }

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
        /// 流程流转到文控时，文控填写收发文单位和发文编码
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="DocKeyword"></param>
        /// <returns></returns>
        public static JObject LetterCNFillInfo(string sid, string DocKeyword, string DocAttrJson) {
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
                else if (string.IsNullOrEmpty(mainFeeder))
                {
                    reJo.msg = "请填写主送！";
                    return reJo.Value;
                }
                else if (string.IsNullOrEmpty(sender))
                {
                    reJo.msg = "请选择发送方！";
                    return reJo.Value;
                }




                #endregion

                #region 设置信函文档附加属性


                AttrData data;
                //主送
                if ((data = doc.GetAttrDataByKeyWord("CA_MAINFEEDER")) != null)
                {
                    data.SetCodeDesc(mainFeeder);
                }
                //发送方
                if ((data = doc.GetAttrDataByKeyWord("CA_SENDER")) != null)
                {
                    data.SetCodeDesc(sender);
                }
                //发文编码
                if ((data = doc.GetAttrDataByKeyWord("CA_SENDCODE")) != null)
                {
                    data.SetCodeDesc(sendCode);
                }

                //抄送
                if ((data = doc.GetAttrDataByKeyWord("CA_COPY")) != null)
                {
                    data.SetCodeDesc(copyParty);
                }

                ////发送日期
                if ((data = doc.GetAttrDataByKeyWord("CA_SENDDATE")) != null)
                {
                    data.SetCodeDesc(sendDate);
                }

                ////保存项目属性，存进数据库
                doc.AttrDataList.SaveData();

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
                htUserKeyWord.Add("MAINFEEDER", mainFeeder);//主送
                htUserKeyWord.Add("SENDER", sender);//发送方
                //htUserKeyWord.Add("SENDCODE", sendCode);//发文编码
                htUserKeyWord.Add("COPY", copyParty);//抄送
                htUserKeyWord.Add("SENDDATE", strSenddate);//发送日期 

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
                    doc.WorkFlow.O_suser3 = "pass";
                    doc.WorkFlow.Modify();
                }

                reJo.data = new JArray(new JObject(new JProperty("WorkFlowKeyword",doc.WorkFlow.KeyWord)));
                
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

        public static JObject GetLetterCNNumber(string sid, string ProjectCode, string SendCompany, string RecCompany)
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


                //string runNum = getDocNumber(dbsource,ProjectCode, SendCompany, RecCompany);
                string runNum =Document. getDocTempNumber(dbsource, ProjectCode,"S","LET", SendCompany, RecCompany);
                if (string.IsNullOrEmpty(runNum)) runNum = "SLET" + "00001";
                    reJo.success = true;
                    reJo.data=new JArray(new JObject(new JProperty("RunNum", runNum)));
   

             
            }
            catch (Exception exception)
            {
                WebApi.CommonController.WebWriteLog(exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace);
                //AssistFun.PopUpPrompt(exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace);
                reJo.msg = "获取信函编号失败！" + exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace;
            }
            return reJo.Value;
        }

        //获取文档最大编号
        private static string getDocNumber(DBSource dbsource,string RootProjectCode, string strSendCompany, string strRecCompany)
        {
            try
            {
                //string RootProjectCode = proj.GetValueByKeyWord("DESIGNPROJECT_CODE");

                //获取文档前缀
                string sendCompanyCode = strSendCompany.IndexOf("__") >= 0 ? strSendCompany.Substring(0, strSendCompany.IndexOf("__")) : strSendCompany;

                string recCompanyCode = strRecCompany.IndexOf("__") >= 0 ? strRecCompany.Substring(0, strRecCompany.IndexOf("__")) : strRecCompany;


                //编码前缀
                string strPrefix = RootProjectCode + "-" + sendCompanyCode + "-" + recCompanyCode + "-LET-" ;

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



        //定义信函附件结构体
        internal struct LetterAttaFile
        {
            // 文件序号
            public string No { get; set; }
            // 文件名称
            public string Name { get; set; }
            
            //文件编码
            public string Code { get; set; }
            
            //文件描述
            public string Desc { get; set; }

            //页数
            public string Page { get; set; }

            //版次
            public string Edition { get; set; }


            //密级
            public string Seculevel { get; set; }

            //状态
            public string Status { get; set; }

            //备注
            public string Remark { get; set; }
        }


    }
}

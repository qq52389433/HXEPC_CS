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
    /// <summary>
    /// 文件传递单
    /// </summary>
    public class TransmittalCN
    {

        /// <summary>
        /// 获取创建信函表单的默认配置
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="ProjectKeyword"></param>
        /// <returns></returns>
        public static JObject GetDraftTransmittalCNDefault(string sid, string ProjectKeyword)
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
                    //if (!(data6.O_sValue2 != str3))
                    if (!string.IsNullOrEmpty(data6.O_sValue1))
                    {
                        joRecCompany.Add(new JProperty(data6.O_sValue1, data6.O_Desc));
                        joSendCompany.Add(new JProperty(data6.O_sValue1, data6.O_Desc));
                        //dictionary.Add(data6.O_Desc, data6.O_sValue1);
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
                    new JProperty("SourceCompany", strCompany)
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
        /// 获取文件传递单编号
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="ProjectCode"></param>
        /// <param name="SendCompany"></param>
        /// <param name="RecCompany"></param>
        /// <returns></returns>
        public static JObject GetTransmittalCNNumber(string sid, string ProjectCode, string SendCompany, string RecCompany)
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


                string runNum = getDocTempNumber(dbsource, ProjectCode, SendCompany, RecCompany);
                if (string.IsNullOrEmpty(runNum)) runNum = "STRA" + "00001";
                reJo.success = true;
                reJo.data = new JArray(new JObject(new JProperty("RunNum", runNum)));

            }
            catch (Exception exception)
            {
                WebApi.CommonController.WebWriteLog(exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace);
                //AssistFun.PopUpPrompt(exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace);
                reJo.msg = "获取文件传递单编号失败！" + exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace;
            }
            return reJo.Value;
        }

        //获取文档最大编号
        private static string getDocNumber(DBSource dbsource, string RootProjectCode, string strSendCompany, string strRecCompany)
        {
            try
            {
                //string RootProjectCode = proj.GetValueByKeyWord("DESIGNPROJECT_CODE");

                //获取文档前缀
                string sendCompanyCode = strSendCompany.IndexOf("__") >= 0 ? strSendCompany.Substring(0, strSendCompany.IndexOf("__")) : strSendCompany;

                string recCompanyCode = strRecCompany.IndexOf("__") >= 0 ? strRecCompany.Substring(0, strRecCompany.IndexOf("__")) : strRecCompany;


                //编码前缀
                string strPrefix = RootProjectCode + "-" + sendCompanyCode + "-" + recCompanyCode + "-TRA-";

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


        //获取文档最大编号
        private static string getDocTempNumber(DBSource dbsource, string RootProjectCode, string strSendCompany, string strRecCompany)
        {
            try
            {
                //string RootProjectCode = proj.GetValueByKeyWord("DESIGNPROJECT_CODE");

                //获取文档前缀
                string sendCompanyCode = strSendCompany.IndexOf("__") >= 0 ? strSendCompany.Substring(0, strSendCompany.IndexOf("__")) : strSendCompany;

                string recCompanyCode = strRecCompany.IndexOf("__") >= 0 ? strRecCompany.Substring(0, strRecCompany.IndexOf("__")) : strRecCompany;


                //编码前缀
               // string strPrefix = RootProjectCode + "-" + sendCompanyCode + "-" + recCompanyCode + "-TRA-"+ "STRA";
                string strPrefix = "";
                if (!string.IsNullOrEmpty(RootProjectCode))
                {
                    //项目管理类
                    strPrefix = RootProjectCode + "-" + sendCompanyCode + "-" + recCompanyCode + "-TRA-" + "STRA";
                }
                else
                {
                    //运营管理类
                    strPrefix = sendCompanyCode + "-" + recCompanyCode + "-TRA-" + "STRA";
                }

                List<Doc> docList = dbsource.SelectDoc(string.Format("select * from CDMS_Doc where o_itemname like '%{0}%' and o_dmsstatus !=10 order by o_itemname", strPrefix));
                if (docList == null || docList.Count == 0)
                {
                    return "STRA"+"00001";
                }
                else
                {
                    Doc doc = docList[docList.Count - 1];
                    //int docCodeLength = 0;
                    //docCodeLength = strPrefix.IndexOf("-LET-");
                    //docCodeLength = docCodeLength > 0 ? 0 : docCodeLength + 3 + 1;
                    //int tempNum = Convert.ToInt32(doc.O_itemname.Substring(docCodeLength, 3));

                    int tempNum = Convert.ToInt32(doc.O_itemname.Substring(strPrefix.Length, 5));
                    //3位数，不够位数补零
                    return (tempNum + 1).ToString("d5");
                }
            }
            catch
            {
                return "00001";
            }

        }


        //线程锁 
        internal static Mutex muxConsole = new Mutex();
        public static JObject DraftTransmittalCN(string sid, string ProjectKeyword, string DocAttrJson, string FileListJson)
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
                string fileCode = "", documentCode ="",mainFeeder = "", copyParty = "",
                    sender = "", sendDate = "", totalPages = "",
                    seculevel = "", secrTerm = "", needreply = "",
                    replyDate = "", transmode="", transmodeSupp="",
                    purpose="", purposeSupp="",title = "", 
                    content = "", approvpath="";



                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();
  
                    //获取文件编码 
                    if (strName == "fileCode") fileCode = strValue.Trim();

                    //获取函件编号
                    if (strName == "documentCode") documentCode = strValue.Trim();

                    //获取主送
                    if (strName == "mainFeeder") mainFeeder = strValue.Trim();
                    //获取发送方
                    else if (strName == "sender") sender = strValue;

                    //获取发送日期
                    else if (strName == "sendDate") sendDate = strValue;
                    //获取页数
                    else if (strName == "totalPages") totalPages = strValue;

                    //获取保密等级
                    else if (strName == "seculevel") seculevel = strValue;
                    //获取保密期限
                    else if (strName == "secrTerm") secrTerm = strValue;

                    //获取是否需要回复
                    else if (strName == "needreply") needreply = strValue;
                    //获取回文日期
                    else if (strName == "replyDate") replyDate = strValue;

                    //获取传递方式
                    else if (strName == "transmode") transmode = strValue;
                    //获取传递方式补充说明
                    else if (strName == "transmodeSupp") transmodeSupp = strValue;

                    //获取提交目的
                    else if (strName == "purpose") purpose = strValue;
                    //获取提交目的补充说明
                    else if (strName == "purposeSupp") purposeSupp = strValue;

                    //获取标题
                    else if (strName == "title") title = strValue;
                    //获取正文内容
                    else if (strName == "content") content = strValue;

                    else if (strName == "approvpath") approvpath = strValue;
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
                else if (string.IsNullOrEmpty(mainFeeder))
                {
                    reJo.msg = "请填写主送单位！";
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

                #region 根据文件传递单模板，生成文件传递单文档

                //获取立项单文档所在的目录
                //Project m_Project = m_NewProject;

                List<TempDefn> docTempDefnByCode = m_Project.dBSource.GetTempDefnByCode("FILETRANSMIT");
                TempDefn docTempDefn = (docTempDefnByCode != null && docTempDefnByCode.Count > 0) ? docTempDefnByCode[0] : null;
                if (docTempDefn == null)
                {
                    reJo.msg = "没有与其相关的模板管理，创建无法正常完成";
                    return reJo.Value;
                }

                IEnumerable<string> source = from docx in m_Project.DocList select docx.Code;
                string filename = documentCode + " " + title;
                if (source.Contains<string>(filename))
                {
                    for (int i = 1; i < 0x3e8; i++)
                    {
                        filename = documentCode + " " + title + i.ToString();
                        if (!source.Contains<string>(filename))
                        {
                            //reJo.msg = "新建文件传递单出错！";
                            //return reJo.Value;
                            break;
                        }
                    }
                }

                //文档名称
                Doc docItem = m_Project.NewDoc(filename + ".xlsx", filename, "", docTempDefn);
                if (docItem == null)
                {
                    reJo.msg = "新建文件传递单出错！";
                    return reJo.Value;
                }

                #endregion

                #region 设置文件传递单文档附加属性


                AttrData data;

                //文件编码
                if ((data = docItem.GetAttrDataByKeyWord("FI_FILECODE")) != null)
                {
                    data.SetCodeDesc(fileCode);
                }

                //主送
                if ((data = docItem.GetAttrDataByKeyWord("FI_MAINFEEDER")) != null)
                {
                    data.SetCodeDesc(mainFeeder);
                }
                //发送方
                if ((data = docItem.GetAttrDataByKeyWord("FI_SENDER")) != null)
                {
                    data.SetCodeDesc(sender);
                }

                //发文日期
                if ((data = docItem.GetAttrDataByKeyWord("FI_SENDDATE")) != null)
                {
                    data.SetCodeDesc(sendDate);
                }

                //页数
                if ((data = docItem.GetAttrDataByKeyWord("FI_PAGE")) != null)
                {
                    data.SetCodeDesc(totalPages);
                }

                //保密等级
                if ((data = docItem.GetAttrDataByKeyWord("FI_SECRETGRADE")) != null)
                {
                    data.SetCodeDesc(seculevel);
                }
                //保密期限
                if ((data = docItem.GetAttrDataByKeyWord("FI_SECRETTERM")) != null)
                {
                    data.SetCodeDesc(secrTerm);
                }

                //获取是否需要回复
                if ((data = docItem.GetAttrDataByKeyWord("FI_IFREPLY")) != null)
                {
                    data.SetCodeDesc(needreply);
                }
                //回复日期
                if ((data = docItem.GetAttrDataByKeyWord("FI_REPLYDATE")) != null)
                {
                    data.SetCodeDesc(replyDate);
                }

                //传递方式
                if ((data = docItem.GetAttrDataByKeyWord("FI_TRANSMITMETHOD")) != null)
                {
                    data.SetCodeDesc(transmode);
                }
                //传递方式的补充说明
                if ((data = docItem.GetAttrDataByKeyWord("FI_TRANSMITMETHODSUPP")) != null)
                {
                    data.SetCodeDesc(transmodeSupp);
                }

                //提交目的
                if ((data = docItem.GetAttrDataByKeyWord("FI_SUBMISSIONOBJ")) != null)
                {
                    data.SetCodeDesc(purpose);
                }
                //提交目的补充说明
                if ((data = docItem.GetAttrDataByKeyWord("FI_SUBMISSIONOBJSUPP")) != null)
                {
                    data.SetCodeDesc(purposeSupp);
                }

                //标题
                if ((data = docItem.GetAttrDataByKeyWord("FI_TRANSMITTITLE")) != null)
                {
                    data.SetCodeDesc(title);
                }

                //摘要
                if ((data = docItem.GetAttrDataByKeyWord("FI_ABSTRACT")) != null)
                {
                    data.SetCodeDesc(content);
                }

                //校审级数（审批路径）
                if ((data = docItem.GetAttrDataByKeyWord("CA_SERIES")) != null)
                {
                    data.SetCodeDesc(approvpath);

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

                ////////////////保密等级////////////////////////
                string SECULEVEL_SECRET = "□";//商业秘密
                string SECULEVEL_LIMIT = "□";//受限
                string SECULEVEL_PUBLIC = "□";//公开

                string SECULEVEL_SECRET_TERM = "";//保密期限
                string SECULEVEL_LIMIT_TERM = "";//保密期限

                ///////////////是否需要回复/////////////////////
                string NEEDREPLY_YES = "□";//是
                string NEEDREPLY_NO = "□";//否

                ///////////////传递方式/////////////////////
                string TRANSMITMETHOD_POST = "□";//邮寄
                string TRANSMITMETHOD_ONFACE = "□";//当面递交
                string TRANSMITMETHOD_MAIL = "□";//邮件
                string TRANSMITMETHOD_OA = "□";//OA
                string TRANSMITMETHOD_CDMS = "□";//CDMS
                string TRANSMITMETHOD_OTHER = "□";//其他

                string TRANSMITMETHOD_SUPP = "";//传递方式补充说明

                //////////////提交目的/////////////////////
                string SUBMISSIONOBJ_SUBMI = "□";//提交
                string SUBMISSIONOBJ_DEMAND = "□";//按需求提交
                string SUBMISSIONOBJ_CHECK = "□";//审查
                string SUBMISSIONOBJ_RECORD = "□";//备案
                string SUBMISSIONOBJ_PURCHASE = "□";//采购
                string SUBMISSIONOBJ_SUPPLY = "□";//供货
                string SUBMISSIONOBJ_CONS = "□";//施工
                string SUBMISSIONOBJ_DEBUG = "□";//调试
                string SUBMISSIONOBJ_INFORM = "□";//告知
                string SUBMISSIONOBJ_HANDOVER = "□";//交工资料
                string SUBMISSIONOBJ_OTHER = "□";//其他

                string SUBMISSIONOBJ_SUPP = ""; //提交目的补充说明

                #endregion

                #region 勾选选项


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

                //勾选传递方式
                if (transmode == "邮寄")
                {
                    TRANSMITMETHOD_POST = "☑";
                }
                else if (transmode == "当面递交")
                {
                    TRANSMITMETHOD_ONFACE = "☑";
                }
                else if (transmode == "邮件")
                {
                    TRANSMITMETHOD_MAIL = "☑";
                }
                else if (transmode == "OA")
                {
                    TRANSMITMETHOD_OA = "☑";
                }
                else if (transmode == "CDMS")
                {
                    TRANSMITMETHOD_CDMS = "☑";
                }
                else if (transmode == "其他")
                {
                    TRANSMITMETHOD_OTHER = "☑";
                    TRANSMITMETHOD_SUPP = transmodeSupp;//传递方式补充说明
                }


                //勾选提交目的
                if (purpose == "提交")
                {
                    SUBMISSIONOBJ_SUBMI = "☑";
                }
                else if (purpose == "按需求提交")
                {
                    SUBMISSIONOBJ_DEMAND = "☑";
                }
                else if (purpose == "审查")
                {
                    SUBMISSIONOBJ_CHECK = "☑";
                }
                else if (purpose == "备案")
                {
                    SUBMISSIONOBJ_RECORD = "☑";
                }
                else if (purpose == "采购")
                {
                    SUBMISSIONOBJ_PURCHASE = "☑";
                }
                else if (purpose == "供货")
                {
                    SUBMISSIONOBJ_SUPPLY = "☑";
                }
                else if (purpose == "施工")
                {
                    SUBMISSIONOBJ_CONS = "☑";
                }
                else if (purpose == "调试")
                {
                    SUBMISSIONOBJ_DEBUG = "☑";
                }
                else if (purpose == "告知")
                {
                    SUBMISSIONOBJ_INFORM = "☑";
                }
                else if (purpose == "交工资料")
                {
                    SUBMISSIONOBJ_HANDOVER = "☑";
                }
                else if (purpose == "其他")
                {
                    SUBMISSIONOBJ_OTHER = "☑";
                    SUBMISSIONOBJ_SUPP = purposeSupp;//提交目的补充说明
                }


                #endregion

                #region 添加勾选项到哈希表

                ////////////////保密等级////////////////////////
                htUserKeyWord.Add("C101", SECULEVEL_SECRET);//商业秘密
                htUserKeyWord.Add("C102", SECULEVEL_LIMIT);//受限
                htUserKeyWord.Add("C103", SECULEVEL_PUBLIC);//公开

                htUserKeyWord.Add("T001", SECULEVEL_SECRET_TERM); //商业秘密保密期限
                htUserKeyWord.Add("T002", SECULEVEL_LIMIT_TERM); //受限保密期限

                ///////////////是否需要回复/////////////////////
                htUserKeyWord.Add("C201", NEEDREPLY_YES);//是
                htUserKeyWord.Add("C202", NEEDREPLY_NO);//否

                ///////////////传递方式/////////////////////
                htUserKeyWord.Add("C301", TRANSMITMETHOD_POST);//邮寄
                htUserKeyWord.Add("C302", TRANSMITMETHOD_ONFACE);//当面递交
                htUserKeyWord.Add("C303", TRANSMITMETHOD_MAIL);//邮件
                htUserKeyWord.Add("C304", TRANSMITMETHOD_OA);//OA
                htUserKeyWord.Add("C305", TRANSMITMETHOD_CDMS);//CDMS
                htUserKeyWord.Add("C306", TRANSMITMETHOD_OTHER);//其他

                htUserKeyWord.Add("T003", TRANSMITMETHOD_SUPP);//传递方式补充说明

                //////////////提交目的/////////////////////
                htUserKeyWord.Add("C401", SUBMISSIONOBJ_SUBMI);//提交
                htUserKeyWord.Add("C402", SUBMISSIONOBJ_DEMAND);//按需求提交
                htUserKeyWord.Add("C403", SUBMISSIONOBJ_CHECK);//审查
                htUserKeyWord.Add("C404", SUBMISSIONOBJ_RECORD);//备案
                htUserKeyWord.Add("C405", SUBMISSIONOBJ_PURCHASE);//采购
                htUserKeyWord.Add("C406", SUBMISSIONOBJ_SUPPLY);//供货
                htUserKeyWord.Add("C407", SUBMISSIONOBJ_CONS);//施工
                htUserKeyWord.Add("C408", SUBMISSIONOBJ_DEBUG);//调试
                htUserKeyWord.Add("C409", SUBMISSIONOBJ_INFORM);//告知
                htUserKeyWord.Add("C410", SUBMISSIONOBJ_HANDOVER);//交工资料
                htUserKeyWord.Add("C411", SUBMISSIONOBJ_OTHER);//其他

                htUserKeyWord.Add("T004", SUBMISSIONOBJ_SUPP);//提交目的补充说明


                #endregion
                #endregion

                //格式化日期
                DateTime senddate = Convert.ToDateTime(sendDate);
                DateTime replydate = Convert.ToDateTime(replyDate);
                string strSenddate = senddate.ToShortDateString().ToString().Replace("-", ".").Replace("/", ".");
                string strReplydate = replydate.ToShortDateString().ToString().Replace("-", ".").Replace("/", ".");

                //htUserKeyWord.Add("DOCUMENTCODE", documentCode);//函件编号
                //htUserKeyWord.Add("MAINFEEDER", mainFeeder);//主送
                //htUserKeyWord.Add("SENDER", sender);//发送方
                //htUserKeyWord.Add("SENDDATE", strSenddate);//发送日期
                //htUserKeyWord.Add("PAGE", totalPages);//页数

                //htUserKeyWord.Add("REPLYDATE", strReplydate);//回复日期

                //htUserKeyWord.Add("TITLE", title);//标题
                //htUserKeyWord.Add("CONTENT", content);//摘要正文

                //htUserKeyWord.Add("T005", documentCode);//函件编号
                htUserKeyWord.Add("T006", mainFeeder);//主送
                htUserKeyWord.Add("T007", sender);//发送方
                htUserKeyWord.Add("T008", strSenddate);//发送日期
                htUserKeyWord.Add("T009", totalPages);//页数

                htUserKeyWord.Add("T010", strReplydate);//回复日期

                htUserKeyWord.Add("T011", title);//标题
                htUserKeyWord.Add("T012", content);//摘要正文

                //htUserKeyWord.Add("RIGHT__HEADER", "编码："+ fileCode);//设置右边页眉

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

                string workingPath = m_Project.dBSource.LoginUser.WorkingPath;


                try
                {
                    //上传下载文档
                    string exchangfilename = "文件传递单中文模板";

                    //获取网站路径
                    string sPath = System.Web.HttpContext.Current.Server.MapPath("/ISO/HXEPC/");

                    //获取模板文件路径
                    string modelFileName = sPath + exchangfilename + ".xlsx";

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
                        new JProperty("DocKeyword", docItem.KeyWord), new JProperty("DocList", strDocList)));
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

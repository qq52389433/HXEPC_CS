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
    public class MeetMinutesCN
    {
        //GetDraftMeetMinutesCNDefault

        /// <summary>
        /// 获取创建会议纪要表单的默认配置
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="ProjectKeyword"></param>
        /// <returns></returns>
        public static JObject GetDraftMeetMinutesCNDefault(string sid, string ProjectKeyword)
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

        //线程锁 
        internal static Mutex muxConsole = new Mutex();

        /// <summary>
        /// 起草会议纪要
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="ProjectKeyword"></param>
        /// <param name="DocAttrJson"></param>
        /// <returns></returns>
        public static JObject DraftMeetMinutesCN(string sid, string ProjectKeyword, string DocAttrJson)
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
                string fileCode = "", sendCode = "",//documentCode = "",
                    sendDate = "", totalPages = "",
                    mainFeeder = "", copyParty = "", title = "",
                    meetTime = "", meetPlace = "", hostUnit = "",
                    moderator = "", participants = "", content = "",
                    approvpath = "", nextStateUserList = "";



                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    //获取文件编码
                    if (strName == "fileCode") fileCode = strValue.Trim();

                    //获取发文编号
                    else if (strName == "sendCode") sendCode = strValue.Trim();

                    //获取发文日期
                    else if (strName == "sendDate") sendDate = strValue;
                    //获取页数
                    else if (strName == "totalPages") totalPages = strValue;

                    //获取主送
                    if (strName == "mainFeeder") mainFeeder = strValue.Trim();
                    //获取抄送
                    else if (strName == "copyParty") copyParty = strValue;

                    //获取会议主题
                    else if (strName == "title") title = strValue;
                    //获取会议时间
                    else if (strName == "meetTime") meetTime = strValue;
                    //获取会议地点
                    else if (strName == "meetPlace") meetPlace = strValue;

                    //获取主办单位
                    else if (strName == "hostUnit") hostUnit = strValue;
                    //获取主持人
                    else if (strName == "moderator") moderator = strValue;
                    //获取参会单位与人员
                    else if (strName == "participants") participants = strValue;

                    //获取会议内容
                    else if (strName == "content") content = strValue;

                    //获取审批路径
                    else if (strName == "approvpath") approvpath = strValue;

                    //获取下一状态人员
                    else if (strName == "nextStateUserList") nextStateUserList = strValue;

                }


                if (string.IsNullOrEmpty(fileCode))
                {
                    reJo.msg = "请填写文件编号！";
                    return reJo.Value;
                }

                else if (string.IsNullOrEmpty(sendCode))
                {
                    reJo.msg = "请填写发文编码！";
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

                #endregion


                #region 根据会议纪要模板，生成会议纪要文档

                //获取立项单文档所在的目录
                //Project m_Project = m_NewProject;

                List<TempDefn> docTempDefnByCode = m_Project.dBSource.GetTempDefnByCode("MEETINGSUMMARY");
                TempDefn docTempDefn = (docTempDefnByCode != null && docTempDefnByCode.Count > 0) ? docTempDefnByCode[0] : null;
                if (docTempDefn == null)
                {
                    reJo.msg = "没有与其相关的模板管理，创建无法正常完成";
                    return reJo.Value;
                }

                IEnumerable<string> source = from docx in m_Project.DocList select docx.Code;
                string filename = sendCode + " " + title;
                if (source.Contains<string>(filename))
                {
                    for (int i = 1; i < 0x3e8; i++)
                    {
                        filename = sendCode + " " + title + i.ToString();
                        if (!source.Contains<string>(filename))
                        {
                            //reJo.msg = "新建会议纪要出错！";
                            //return reJo.Value;
                            break;
                        }
                    }
                }

                //文档名称
                Doc docItem = m_Project.NewDoc(filename + ".docx", filename, "", docTempDefn);
                if (docItem == null)
                {
                    reJo.msg = "新建会议纪要出错！";
                    return reJo.Value;
                }

                #endregion

                #region 设置文会议纪要文档附加属性


                AttrData data;


                //函件编号
                if ((data = docItem.GetAttrDataByKeyWord("ME_FILECODE")) != null)
                {
                    data.SetCodeDesc(fileCode);
                }

                //会议主题
                if ((data = docItem.GetAttrDataByKeyWord("ME_TITLE")) != null)
                {
                    data.SetCodeDesc(title);
                }

                //会议时间
                if ((data = docItem.GetAttrDataByKeyWord("ME_TIME")) != null)
                {
                    data.SetCodeDesc(meetTime);
                }

                //函件编号
                if ((data = docItem.GetAttrDataByKeyWord("ME_SENDCODE")) != null)
                {
                    data.SetCodeDesc(sendCode);
                }

                //发文日期
                if ((data = docItem.GetAttrDataByKeyWord("ME_SENDDATE")) != null)
                {
                    data.SetCodeDesc(sendDate);
                }

                //主送
                if ((data = docItem.GetAttrDataByKeyWord("ME_MAINSEND")) != null)
                {
                    data.SetCodeDesc(mainFeeder);
                }

                //页数
                if ((data = docItem.GetAttrDataByKeyWord("ME_PAGE")) != null)
                {
                    data.SetCodeDesc(totalPages);
                }

                //抄送
                if ((data = docItem.GetAttrDataByKeyWord("ME_COPY")) != null)
                {
                    data.SetCodeDesc(copyParty);
                }


                ////保存项目属性，存进数据库
                docItem.AttrDataList.SaveData();

                #endregion

                #region 录入数据进入word表单

                string strDocList = "";//获取附件

                //录入数据进入表单
                Hashtable htUserKeyWord = new Hashtable();


                //格式化日期
                DateTime senddate = Convert.ToDateTime(sendDate);

                string strSenddate = senddate.ToShortDateString().ToString().Replace("-", ".").Replace("/", ".");


                htUserKeyWord.Add("HEADERCODE", fileCode);//页眉里面的发文编码

                htUserKeyWord.Add("SENDDATE", strSenddate);//发送日期
                htUserKeyWord.Add("PAGE", totalPages);//页数

                htUserKeyWord.Add("MAINFEEDER", mainFeeder);//主送
                //htUserKeyWord.Add("DOCUMENTCODE", documentCode);//发文编码

                htUserKeyWord.Add("COPY", copyParty);//抄送方

                htUserKeyWord.Add("TITLE", title);//标题
                htUserKeyWord.Add("MEETTIME", meetTime);//会议时间
                htUserKeyWord.Add("MEETPLACE", meetPlace);//会议地点

                htUserKeyWord.Add("HOSTUNIT", hostUnit);//主办单位
                htUserKeyWord.Add("MODERATOR", moderator);//主持人
                htUserKeyWord.Add("PARTICIPANTS", participants);//参会单位与人员

                htUserKeyWord.Add("CONTENT", content);//会议内容


                string workingPath = m_Project.dBSource.LoginUser.WorkingPath;


                try
                {
                    //上传下载文档
                    string exchangfilename = "会议纪要中文模板";

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

        public static JObject GetMeetMinutesCNNumber(string sid, string ProjectCode, string SendCompany, string RecCompany)
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
                string runNum = Document.getDocTempNumber(dbsource, ProjectCode, "S", "MOM", SendCompany, RecCompany);
                if (string.IsNullOrEmpty(runNum)) runNum = "SMOM" + "00001";
                reJo.success = true;
                reJo.data = new JArray(new JObject(new JProperty("RunNum", runNum)));



            }
            catch (Exception exception)
            {
                WebApi.CommonController.WebWriteLog(exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace);
                //AssistFun.PopUpPrompt(exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace);
                reJo.msg = "获取会议纪要编号失败！" + exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace;
            }
            return reJo.Value;
        }
    }
}

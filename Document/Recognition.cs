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
    public class Recognition
    {
        /// <summary>
        /// 获取创建认质认价表单的默认配置
        /// </summary>
        /// <param name="sid"></param>
        /// <param name="ProjectKeyword"></param>
        /// <returns></returns>
        public static JObject GetDraftRecognitionDefault(string sid, string ProjectKeyword)
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
        public static JObject DraftRecognition(string sid, string ProjectKeyword, string DocAttrJson, string ContentJson)
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


                if (m_Project == null)
                {
                    reJo.msg = "参数错误！文件夹不存在！";
                    return reJo.Value;
                }

                #region 获取信函参数内容

                //获取信函参数内容
                string fileCode = "", projectName = "", contractCode = "",
                   sendDate = "", sendCode = "", recCode = "",
                   recType = "", materialType = "";

                JArray jaAttr = (JArray)JsonConvert.DeserializeObject(DocAttrJson);

                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    ////获取函件编号
                    //if (strName == "documentCode") documentCode = strValue.Trim();

                    //获取文件编码 
                    if (strName == "fileCode") fileCode = strValue.Trim();

                    //获取项目名称 
                    else if (strName == "projectName") projectName = strValue;

                    //获取合同号 
                    else if (strName == "contractCode") contractCode = strValue;

                    //获取发文日期
                    else if (strName == "sendDate") sendDate = strValue;

                    //获取发文编码
                    else if (strName == "sendCode") sendCode = strValue;

                    //获取收文编码
                    else if (strName == "recCode") recCode = strValue;

                    //获取来文类型
                    else if (strName == "recType") recType = strValue;

                    //获取物资类型
                    else if (strName == "materialType") materialType = strValue;

                }


                if (string.IsNullOrEmpty(fileCode))
                {
                    reJo.msg = "请填写文件编号！";
                    return reJo.Value;
                }
                else if (string.IsNullOrEmpty(projectName))
                {
                    reJo.msg = "请填写工程名称！";
                    return reJo.Value;
                }


                #endregion

                #region 获取材料或设备列表
                List<Material> materialList = new List<Material>();
                if (!string.IsNullOrEmpty(ContentJson))
                {
                    int index = 0;
                    JArray jaFiles = (JArray)JsonConvert.DeserializeObject(ContentJson);

                    foreach (JObject joAttr in jaFiles)
                    {
                        string strMatName = joAttr["matName"].ToString();//材料（设备）名称
                        string strSpec = joAttr["spec"].ToString();//规格型号
                        string strMeaUnit = joAttr["meaUnit"].ToString();//计量单位
                        string strDesignNum = joAttr["designNum"].ToString();//设计图号
                        string strBrand = joAttr["brand"].ToString();//品牌
                        string strQuantity = joAttr["quantity"].ToString();//报审数量
                        string strAudit = joAttr["audit"].ToString();//意见
                        string strPrice = joAttr["price"].ToString();//报审单价
                        string strCostPrice = joAttr["costPrice"].ToString();//造价员单价
                        string strCenterPrice = joAttr["centerPrice"].ToString();//财务中心单价
                        string strTenderPrice = joAttr["tenderPrice"].ToString();//招标部单价
                        string strAuditPrice = joAttr["auditPrice"].ToString();//审核 合价
                        string strRemark = joAttr["remark"].ToString();//备注

                       if (string.IsNullOrEmpty(strMatName)) { continue; }

                        index++;
                        string strIndex = index.ToString();
                        Material afItem = new Material()
                        {
                            No = strIndex,
                            MatName = strMatName,
                            Spec = strSpec,
                            MeaUnit = strMeaUnit,
                            DesignNum = strDesignNum,
                            Brand = strBrand,
                            Quantity = strQuantity,
                            Audit = strAudit,
                            Price = strPrice,
                            CostPrice = strCostPrice,
                            CenterPrice = strCenterPrice,
                            TenderPrice = strTenderPrice,
                            AuditPrice = strAuditPrice,
                            Remark = strRemark
                        };

                        materialList.Add(afItem);
                    }
                }
                #endregion


                #region 根据信函模板，生成信函文档

                //获取立项单文档所在的目录
                //Project m_Project = m_NewProject;

                List<TempDefn> docTempDefnByCode = m_Project.dBSource.GetTempDefnByCode("PRICEFILE");
                TempDefn docTempDefn = (docTempDefnByCode != null && docTempDefnByCode.Count>0) ? docTempDefnByCode[0] : null;
                if (docTempDefn == null)
                {
                    reJo.msg = "没有与其相关的模板管理，创建无法正常完成";
                    return reJo.Value;
                }

                IEnumerable<string> source = from docx in m_Project.DocList select docx.Code;
                string filename = fileCode + " " + projectName+"认质认价报审单";
                if (source.Contains<string>(filename))
                {
                    for (int i = 1; i < 0x3e8; i++)
                    {
                        filename = fileCode + " " + projectName + "认质认价报审单" + i.ToString();
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
                    reJo.msg = "新建认质认价单出错！";
                    return reJo.Value;
                }

                #endregion

                #region 设置信函文档附加属性


                AttrData data;
                //文件编码
                if ((data = docItem.GetAttrDataByKeyWord("PR_FILECODE")) != null)
                {
                    data.SetCodeDesc(fileCode);
                }
                //工程名称
                if ((data = docItem.GetAttrDataByKeyWord("PR_PROJECTNAME")) != null)
                {
                    data.SetCodeDesc(projectName);
                }
                //合同号
                if ((data = docItem.GetAttrDataByKeyWord("PR_CONTRACTCODE")) != null)
                {
                    data.SetCodeDesc(contractCode);
                }
                //发文日期
                if ((data = docItem.GetAttrDataByKeyWord("PR_SENDDATE")) != null)
                {
                    data.SetCodeDesc(sendDate);
                }

                //发文编号
                if ((data = docItem.GetAttrDataByKeyWord("PR_SENDCODE")) != null)
                {
                    data.SetCodeDesc(sendCode);
                }
                //来文编号
                if ((data = docItem.GetAttrDataByKeyWord("PR_RECCODE")) != null)
                {
                    data.SetCodeDesc(recCode);
                }
                //来文类型
                if ((data = docItem.GetAttrDataByKeyWord("PR_RECTYPE")) != null)
                {
                    data.SetCodeDesc(recType);
                }
                //采购类型
                if ((data = docItem.GetAttrDataByKeyWord("PR_MATERIALTYPE")) != null)
                {
                    data.SetCodeDesc(materialType);
                }
                //材料列表
                if ((data = docItem.GetAttrDataByKeyWord("PR_CONTENT")) != null)
                {
                    data.SetCodeDesc(ContentJson);
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
 

                htUserKeyWord.Add("FILECODE", fileCode);//文件编码
                htUserKeyWord.Add("PROJECTNAME", projectName);//发送方
                htUserKeyWord.Add("CONTRACTCODE", contractCode);//合同号
                htUserKeyWord.Add("SENDDATE", strSenddate);//日期
                htUserKeyWord.Add("SENDCODE", sendCode);//编号
                htUserKeyWord.Add("RECCODE", recCode);//来文编号
                htUserKeyWord.Add("RECTYPE", recType);//来文类型
                htUserKeyWord.Add("MATERIALTYPE", materialType);//采购类型 （材料或者设备）
                

                //htUserKeyWord.Add("RHEADER", sendCode);//页眉里面的编码


                //htUserKeyWord["PREPAREDSIGN"] = curUser.O_username;
                //htUserKeyWord["DRAFTTIME"] = DateTime.Now.ToString("yyyy.MM.dd");


                #region 获取项目名称
                Project proj = docItem.Project;
                //Project rootProj = new Project();
                string rootProjDesc = "";
                Project rootProj = CommonFunction.getParentProjectByTempDefn(proj, "HXNY_DOCUMENTSYSTEM");

                #endregion
                string proSource = "";
                if (rootProj != null)
                {
                    proSource = rootProj.GetValueByKeyWord("PRO_COMPANY").ToString();

                    //    //string proAddress = rootProj.GetValueByKeyWord("PRO_ADDRESS").ToString();
                    //    //string proTel = rootProj.GetValueByKeyWord("PRO_NUMBER").ToString();

                    //    //htUserKeyWord.Add("PRO_ADDRESS", proAddress);//项目地址
                    //    //htUserKeyWord.Add("PRO_TEL", proTel);//项目电话

                    //    //htUserKeyWord["PROJECTDESC"] = "（" + rootProj.Description + "项目部）";
                }

                    //添加附件
                    List<string> list3 = new List<string>();
                foreach (Material mat in materialList)
                {
                    list3.Add(mat.No);
                    list3.Add(mat.MatName);
                    list3.Add(mat.Spec);
                    list3.Add(mat.MeaUnit);
                    list3.Add(mat.DesignNum);
                    list3.Add(mat.Brand);
                    list3.Add(mat.Quantity);
                    list3.Add(mat.Audit);
                    list3.Add(mat.Price);
                    list3.Add(mat.CostPrice);
                    list3.Add(mat.CenterPrice);
                    list3.Add(mat.TenderPrice);
                    list3.Add(mat.AuditPrice);
                    list3.Add(mat.Remark);
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
                    string exchangfilename = (proSource == "CWEC") ? "认质认价工程有限" : "认质认价工业股份";


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
        public static JObject RecognitionStartWorkFlow(string sid, string docKeyword, string DocList,  string UserList)
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

                    WorkFlow flow = dbsource.NewWorkFlow(m_DocList, "RECOGNITION");
                    //if (flow == null || flow.CuWorkState == null || flow.CuWorkState.workStateBranchList == null || (flow.CuWorkState.workStateBranchList.Count <= 0))
                    if (flow == null)
                    {
                        reJo.msg = "自动启动流程失败!请手动启动";
                        return reJo.Value;
                    }



                    //获取下一状态
                    //me.approvpathdata = [{ text: "二级-编批", value: "二级-编批" }, { text: "三级-编审批", value: "三级-编审批" },
                    //{ text: "四级-编审定批", value: "四级-编审定批" }, { text: "五级-编校审定批", value: "五级-编校审定批" }];

                    WorkState ws = new WorkState();

                    DefWorkState dws = flow.DefWorkFlow.DefWorkStateList.Find(s => s.KeyWord == "APPROV");// CHECK");
                    ws.DefWorkState = dws;

                    //if (ApprovPath == "二级-编批")
                    //{
                    //    DefWorkState dws = flow.DefWorkFlow.DefWorkStateList.Find(s => s.KeyWord == "APPROV");// CHECK");
                    //    ws.DefWorkState = dws;
                    //}
                    //else if (ApprovPath == "三级-编审批" || ApprovPath == "四级-编审定批")
                    //{
                    //    DefWorkState dws = flow.DefWorkFlow.DefWorkStateList.Find(s => s.KeyWord == "AUDIT");// CHECK");
                    //    ws.DefWorkState = dws;
                    //}
                    //else if (ApprovPath == "五级-编校审定批")
                    //{
                    //    DefWorkState dws = flow.DefWorkFlow.DefWorkStateList.Find(s => s.KeyWord == "CHECK");// CHECK");
                    //    ws.DefWorkState = dws;

                    //}
                    //else
                    //{
                    //    flow.Delete();
                    //    flow.Delete();

                    //    reJo.msg = "审批路径参数错误，自动启动流程失败！请手动启动流程";
                    //    return reJo.Value;
                    //}

                    //foreach (User user in group.UserList)
                    //{
                    //    ws.group.AddUser(user);
                    //}
                    //flow.WorkStateList.Add(ws);
                    //flow.Modify();


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

                    return GotoNextReJo.Value;

                }

            }
            catch (Exception exception)
            {
                WebApi.CommonController.WebWriteLog(exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace);
                reJo.msg = "启动流程失败！" + exception.Message + "\r\n" + exception.Source + "\r\n" + exception.StackTrace;
            }
            return reJo.Value;

        }


        //定义信函附件结构体
        internal struct Material
        {
            // 文件序号
            public string No { get; set; }
            // 材料设备 名称
            public string MatName { get; set; }

            //规格
            public string Spec { get; set; }

            //计量单位 
            public string MeaUnit { get; set; }

            //设计图号 
            public string DesignNum { get; set; }

            //品牌
            public string Brand { get; set; }


            //报审数量
            public string Quantity { get; set; }

            //意见 
            public string Audit { get; set; }

            //报审价格
            public string Price { get; set; }
            //价格
            public string CostPrice { get; set; }
            //价格
            public string CenterPrice { get; set; }
            //价格
            public string TenderPrice { get; set; }

            //审核合价
            public string AuditPrice { get; set; }
            //备注
            public string Remark { get; set; }
        }
    }
}

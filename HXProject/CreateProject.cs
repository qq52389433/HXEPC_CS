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
    public class CreateProject
    {
        /// <summary>
        /// 创建项目时，获取表单默认值，填充到combo
        /// </summary>
        /// <param name="sid">sid</param>
        /// <returns></returns>
        //public static JObject GetCreateRootProjectDefault(string sid)
         public static JObject GetCreateProjectListingDefault(string sid)
            
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

                int listIndex = 0;
                ////设计阶段
                //JObject joDesignPhase = new JObject();
                //List<DictData> dictDataList = dbsource.GetDictDataList("DesignPhase");
                //foreach (DictData d in dictDataList)
                //{
                //    //this.designPhase.Items.Add(d.O_Code + "__" + d.O_Desc);
                //    listIndex = listIndex + 1;
                //    joDesignPhase.Add(new JProperty("v" + listIndex.ToString(), d.O_Code + "__" + d.O_Desc));
                //}

                //合同模式
                JObject joContractModel = new JObject();
                listIndex = 0;
                List<DictData> dictDataList = dbsource.GetDictDataList("ProjectContractModel");
                foreach (DictData d in dictDataList)
                {
                    //this.ContractModel.Items.Add(d.O_Code);
                    listIndex = listIndex + 1;
                    joContractModel.Add(new JProperty("v" + listIndex.ToString(), d.O_Code));
                }

                //开发阶段
                JObject joProjectSource = new JObject();
                listIndex = 0;
                dictDataList = dbsource.GetDictDataList("ProSource");
                foreach (DictData d in dictDataList)
                {
                    //this.ProjectSource.Items.Add(d.O_Code);
                    listIndex = listIndex + 1;
                    joProjectSource.Add(new JProperty("v" + listIndex.ToString(), d.O_Code));
                }

                //工程实施性质
                JObject joEngineeringProperties = new JObject();
                listIndex = 0;
                dictDataList = dbsource.GetDictDataList("EngineeringProperties");
                foreach (DictData d in dictDataList)
                {
                    //this.DesignDevelopment.Items.Add(d.O_Code);
                    listIndex = listIndex + 1;
                    joEngineeringProperties.Add(new JProperty("v" + listIndex.ToString(), d.O_Code));
                }

                //项目类型的选择
                JObject joProjectType = new JObject();
                listIndex = 0;
                dictDataList = dbsource.GetDictDataList("ProjectType");
                string key = "";
                foreach (DictData d in dictDataList)
                {

                    if (key.Contains(d.O_Code))
                    {
                        continue;
                    }

                    //this.PROJECTTYPE.Items.Add(d.O_Code);
                    listIndex = listIndex + 1;
                    joProjectType.Add(new JProperty("v" + listIndex.ToString(), d.O_Code));

                    key = d.O_Code;
                }

                reJo.data = new JArray(
                    new JObject(
                    //    new JProperty("DesignPhase", joDesignPhase),
                    new JProperty("ContractType", joContractModel), //合同类型（合同模式）
                    new JProperty("EngineeringProperties", joEngineeringProperties),//项目实施性质
                    new JProperty("ProjectSource", joProjectSource),

                    new JProperty("ProjectType", joProjectType)
                    ));
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
        /// 获取二级项目类型
        /// </summary>
        /// <param name="sid">sid</param>
        /// <param name="ProjectType">一级项目类型</param>
        /// <returns></returns>
        public static JObject GetProjectTypeII(string sid, string ProjectType)
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

                //根据一级项目类型获取二级项目类型
                JObject joData = new JObject();
                List<DictData> dictDataList = dbsource.GetDictDataList("ProjectType");

                int dicIndex = 0;
                foreach (DictData d in dictDataList)
                {
                    if (d.O_Code.Contains(ProjectType))
                    {
                        dicIndex = dicIndex + 1;
                        joData.Add(new JProperty("v" + dicIndex.ToString(), d.O_sValue1));
                    }
                }

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
        //项目立项Web接口
        public static JObject CreateRootProject(string sid, string projectAttrJson)
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
                JArray jaAttr = (JArray)JsonConvert.DeserializeObject(projectAttrJson);

                //reJo.Value = AVEVA.CDMS.HXEPC_Common.EnterPoint.CreateRootProjectX(dbsource, jaAttr);
                reJo = CreateRootProjectX(dbsource, jaAttr);//, sid);


                //刷新数据源
                if (reJo.success == true || reJo.msg == "创建/获取设计阶段失败，请联系管理员！")
                {
                    DBSourceController.RefreshDBSource(sid);
                }

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

        //项目立项
        public static ExReJObject CreateRootProjectX(DBSource dbsource, JArray jaAttr)//, string sid)
        {
            ExReJObject reJo = new ExReJObject();

            try
            {

                #region 获取项目参数项目

                //JArray jaAttr = (JArray)JsonConvert.DeserializeObject(projectAttrJson);

                //获取项目参数项目
                string projectCode = "", projectDescCN = "",sourceUnit="",
                    sourceDesc="",sourceType="", projectAddr = "", projectTel = "",
                    secretarilMan ="";

                //合同号
                string projectNomber = "";
                //备注
                string Remarks = "";

                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    //获取项目代码
                    if (strName == "projectCode") projectCode = strValue.Trim();

                    //获取项目英文名称
                    //else if (strName == "projectDescEN") projectDescEN = strValue;
                    
                    //获取项目中文名称
                    else if (strName == "projectDescCN") projectDescCN = strValue;

                    //获取所属公司
                    else if (strName=="sourceUnit") sourceUnit = strValue;

                    //获取所属公司
                    else if (strName == "sourceDesc") sourceDesc = strValue;

                    //获取国际还是国内项目
                    else if (strName == "sourceType") sourceType = strValue;

                    //获取项目地址
                    else if (strName == "projectAddr") projectAddr = strValue;

                    //获取项目电话
                    else if (strName == "projectTel") projectTel = strValue;

                    //获取项目文控
                    else if (strName == "secretarilMan") secretarilMan = strValue;

                }

                string projectDesc= projectDescCN;
                //string projectDesc = string.IsNullOrEmpty(projectDescEN) ?
                //    projectDescCN
                //    : projectDescCN + " " + projectDescEN;

                if (string.IsNullOrEmpty(projectCode))
                {
                    reJo.msg = "请填写项目编号！";
                    return reJo;
                }
                else if (string.IsNullOrEmpty(projectDesc))
                {
                    reJo.msg = "请填写项目名称！";
                    return reJo;
                }

                #endregion


                //  根据名称查找项目模板(根目录)对象
                List<TempDefn> tempDefnByCode = dbsource.GetTempDefnByCode("HXNY_DOCUMENTSYSTEM");
                TempDefn mTempDefn = (tempDefnByCode != null) ? tempDefnByCode[0] : null;
                if (mTempDefn == null)
                {
                    reJo.msg = "欲创建的项目关联的模板不存在！不能完成创建";
                    return reJo;
                }
                else
                {

                    //获取DBSource的虚拟Local目录
                    //Project m_LocalProject = dbsource.NewProject(enProjectType.Local);

                    Project m_RootProject = dbsource.RootLocalProjectList.Find(itemProj => itemProj.TempDefn.KeyWord == "PRODOCUMENTADMIN");
                    if (m_RootProject == null)
                    {
                        reJo.msg = "[项目管理类文件目录]不存在,或者未添加目录模板！不能完成创建";
                        return reJo;
                    }

                    //查找项目是否已经创建
                    Project findProj= m_RootProject.AllProjectList.Find(itemProj=>itemProj.Code== projectCode);
                    if (findProj != null )
                    {
                        reJo.msg = "项目[" + projectCode + "]已存在！不能完成创建";
                        return reJo;
                    }

                    //检查单位部门数据表，是否有项目来源，没有就添加
                    DictData dictData = null;

                    List<DictData> dictDataList = dbsource.GetDictDataList("Communication");

                    foreach (DictData data6 in dictDataList)
                    {
                        if (data6.O_sValue1 == sourceUnit)
                        {
                            dictData = data6;

                        }
                    }

                    if (dictData == null && string.IsNullOrEmpty(sourceDesc))
                    {
                        reJo.msg = "请输入项目来源名称！";
                        return reJo;

                    }
                    else if (dictData == null)
                    {
                        ///添加到数据字典
                        string format = "insert CDMS_DictData (" +
    "o_parentno,o_datatype,o_ikey,o_skey,o_Code,o_Desc,o_sValue1,o_sValue2,o_sValue3,o_sValue4,o_sValue5,o_iValue1 ,o_iValue2)" +
    " values ({0},{1},{2},'{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}',{11},{12}" + ")";

                        format = string.Format(format, new object[] {
                    0,2,0,"Communication","",sourceDesc,sourceUnit,"","","","",0,0
                     });
                        dbsource.DBExecuteSQL(format);
                       // DBSourceController.refreshDBSource(sid);
                    }

                    //foreach (Project findProj in m_Project.AllProjectList)//.ChildProjectList)
                    //{
                    //    if (findProj != null && findProj.Code == projectCode)
                    //    {
                    //        reJo.msg = "项目[" + projectCode + "]已存在！不能完成创建";
                    //        return reJo;
                    //    }
                    //}

                    //创建项目
                    Project m_NewProject = m_RootProject.NewProject(projectCode, projectDesc, null, mTempDefn);


                    if (m_NewProject == null)
                    {
                        reJo.msg = "新建项目失败";
                        return reJo;

                    }
                    else
                    {
                    


                        #region 设置项目属性


                        AttrData data;

                        //所属公司
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PRO_COMPANY")) != null)
                        {
                            data.SetCodeDesc(sourceUnit);
                        }
                        
                        //国际或国内项目
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PRO_FROM")) != null)
                        {
                            data.SetCodeDesc(sourceType);
                        }

                        //项目地址
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PRO_ADDRESS")) != null)
                        {
                            data.SetCodeDesc(projectAddr);
                        }

                        //项目电话
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PRO_NUMBER")) != null)
                        {
                            data.SetCodeDesc(projectTel);
                        }

                        //项目电话
                        if ((data = m_NewProject.GetAttrDataByKeyWord("SECRETARILMAN")) != null)
                        {
                            data.SetCodeDesc(secretarilMan);
                        }

                        //string[] strArray = (string.IsNullOrEmpty(userlist) ? "" : userlist).Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        //foreach (string struser in strArray)
                        //{
                        //    string userCode = struser.Substring(0, struser.IndexOf("__"));
                        //    strUserCodeList = strUserCodeList + userCode + ",";
                        //}

                        //保存项目属性，存进数据库
                        m_NewProject.AttrDataList.SaveData();

                        #endregion




                        //增加管理员组权限
                        AVEVA.CDMS.Server.Group group = dbsource.GetGroupByName("AdminGroup");
                        if (group != null)
                        {
                            m_NewProject.acceData.AddAcce(group, 32735); //rItemData->iMask
                        }

                        if (!m_NewProject.acceData.Save())
                        {

                        }

                            reJo.data = new JArray(new JObject(new JProperty("projectKeyword", m_NewProject.KeyWord)));
                        reJo.success = true;
                        return reJo;

                        //AVEVA.CDMS.WebApi.DBSourceController.RefreshDBSource(sid);

                    }
                }
            }
            catch (Exception e)
            {
                reJo.msg = e.Message;
                CommonController.WebWriteLog(reJo.msg);
            }

            return reJo;
        }


        public static JObject CreateProjectListing(string sid, string ProjectKeyword, string projectAttrJson)
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
                JArray jaAttr = (JArray)JsonConvert.DeserializeObject(projectAttrJson);

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
                string projectCode = "", projectDescEN = "", projectDescCN = "",
                    projectShortDescEN = "", projectShortDescCN = "", unintName = "", //projectNo = "",
                    projectAddr = "", projectOverview = "", projectType = "",
                    projectTypeII = "", projectTypeSupp = "", projectNature = "",
                    projectNatureSupp = "", projectSource = "", projectSourceSupp = "",
                    contractType = "", contractTypeSupp = "", projectAmount = "",
                    projectDuration = "", planStartDate = "", planEndDate = "",
                    communicationCode = "", intcommunicationCode= "", email = "";

                //合同号
                string projectNomber = "";
                //备注
                string Remarks = "";

                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    //获取项目代码
                    if (strName == "projectCode") projectCode = strValue.Trim();

                    //获取项目英文名称
                    else if (strName == "projectDescEN") projectDescEN = strValue;

                    //获取项目中文名称
                    else if (strName == "projectDescCN") projectDescCN = strValue;

                    //获取项目英文简称
                    else if (strName == "projectShortDescEN") projectShortDescEN = strValue;

                    //获取项目中文简称
                    else if (strName == "projectShortDescCN") projectShortDescCN = strValue;

                    //获取业主
                    else if (strName == "unintName") unintName = strValue;

                    //获取项目地址
                    else if (strName == "projectAddr") projectAddr = strValue;

                    //获取项目概况
                    else if (strName == "projectOverview") projectOverview = strValue;

                    //获取项目类型
                    else if (strName == "projectType") projectType = strValue;

                    //获取项目二级类型
                    else if (strName == "projectTypeII") projectTypeII = strValue;

                    //获取项目类型补充说明
                    else if (strName == "projectTypeSupp") projectTypeSupp = strValue;

                    //获取合同额
                    else if (strName == "projectAmount") projectAmount = strValue;

                    //获取项目实施性质
                    else if (strName == "projectNature") projectNature = strValue;

                    //项目实施性质补充说明
                    else if (strName == "projectNatureSupp") projectNatureSupp = strValue;

                    //获取项目来源
                    else if (strName == "projectSource") projectSource = strValue;

                    //获取项目来源补充说明
                    else if (strName == "projectSourceSupp") projectSourceSupp = strValue;

                    //获取合同类型
                    else if (strName == "contractType") contractType = strValue;

                    //获取合同类型补充
                    else if (strName == "contractTypeSupp") contractTypeSupp = strValue;

                    //获取预估金额
                    else if (strName == "projectAmount") projectAmount = strValue;

                    //获取计划工期
                    else if (strName == "projectDuration") projectDuration = strValue;

                    //获取计划开始时间
                    else if (strName == "planStartDate") planStartDate = strValue;

                    //获取计划结束时间
                    else if (strName == "planEndDate") planEndDate = strValue;

                    //通信代码
                    else if (strName == "communicationCode") communicationCode = strValue;

                    //通信代码
                    else if (strName == "intcommunicationCode") intcommunicationCode = strValue;

                    //通信邮箱
                    else if (strName == "email") email = strValue;

                }

                string projectDesc = string.IsNullOrEmpty(projectDescEN) ?
                    projectDescCN
                    : projectDescCN + " " + projectDescEN;

                if (string.IsNullOrEmpty(projectCode))
                {
                    reJo.msg = "请填写项目编号！";
                    return reJo.Value;
                }
                else if (string.IsNullOrEmpty(projectDesc))
                {
                    reJo.msg = "请填写项目名称！";
                    return reJo.Value;
                }

                #endregion


                #region 根据立项单模板，生成立项单文档

                //获取立项单文档所在的目录
                //Project m_Project = m_NewProject;

                List<TempDefn> docTempDefnByCode = m_Project.dBSource.GetTempDefnByCode("CREATERPROJECT");
                TempDefn docTempDefn = (docTempDefnByCode != null && docTempDefnByCode.Count > 0) ? docTempDefnByCode[0] : null;
                if (docTempDefn == null)
                {
                    reJo.msg = "没有与其相关的模板管理，创建无法正常完成";
                    return reJo.Value;
                }

                IEnumerable<string> source = from docx in m_Project.DocList select docx.Code;
                string filename = projectDesc + "项目立项单";
                if (source.Contains<string>(filename))
                {
                    for (int i = 1; i < 0x3e8; i++)
                    {
                        filename = projectDesc + "项目立项单" + i.ToString();
                        if (!source.Contains<string>(filename))
                        {
                            reJo.msg = "新建项目立项单出错！";
                            return reJo.Value;
                        }
                    }
                }

                //文档名称
                Doc docItem = m_Project.NewDoc(filename + ".doc", filename, "", docTempDefn);
                if (docItem == null)
                {
                    reJo.msg = "新建项目立项单出错！";
                    return reJo.Value;
                }

                #endregion

                #region 设置立项单文档附加属性


                AttrData data;

                //项目代码
                if ((data = docItem.GetAttrDataByKeyWord("PROCODE")) != null)
                {
                    data.SetCodeDesc(projectCode);
                }

                //项目名称（中文）
                if ((data = docItem.GetAttrDataByKeyWord("PRONAME")) != null)
                {
                    data.SetCodeDesc(projectDescCN);
                }
                //项目名称（英文）
                if ((data = docItem.GetAttrDataByKeyWord("PROENGLISH")) != null)
                {
                    data.SetCodeDesc(projectDescEN);
                }
                //项目简称（中文）
                if ((data = docItem.GetAttrDataByKeyWord("ABBREVIATION")) != null)
                {
                    data.SetCodeDesc(projectShortDescCN);
                }
                //项目简称（英文）
                if ((data = docItem.GetAttrDataByKeyWord("ABBREVIATIONENG")) != null)
                {
                    data.SetCodeDesc(projectShortDescEN);
                }

                //通信代码
                if ((data = docItem.GetAttrDataByKeyWord("DOMCOMMUNICATIONCODE")) != null)
                {
                    data.SetCodeDesc(communicationCode);
                }
                //通信代码
                if ((data = docItem.GetAttrDataByKeyWord("INTCOMMUNICATIONCODE")) != null)
                {
                    data.SetCodeDesc(intcommunicationCode);
                }
                //业主
                if ((data = docItem.GetAttrDataByKeyWord("OWNER")) != null)
                {
                    data.SetCodeDesc(unintName);
                }
                //项目实施地址
                if ((data = docItem.GetAttrDataByKeyWord("PROADDRESS")) != null)
                {
                    data.SetCodeDesc(projectAddr);
                }
                //项目实施性质
                if ((data = docItem.GetAttrDataByKeyWord("PRONATURE")) != null)
                {
                    data.SetCodeDesc(projectNature);
                }
                //项目实施性质（说明）
                if ((data = docItem.GetAttrDataByKeyWord("PRONATUREREMARKS")) != null)
                {
                    data.SetCodeDesc(projectNatureSupp);
                }
                //项目类型
                if ((data = docItem.GetAttrDataByKeyWord("PROTYPE")) != null)
                {
                    data.SetCodeDesc(projectType);
                }
                //项目分类(二级)
                if ((data = docItem.GetAttrDataByKeyWord("PROTYPECLASS")) != null)
                {
                    data.SetCodeDesc(projectTypeII);
                }

                //项目类型（说明）
                if ((data = docItem.GetAttrDataByKeyWord("PROTYPEREMARKS")) != null)
                {
                    data.SetCodeDesc(projectTypeSupp);
                }
                //项目来源
                if ((data = docItem.GetAttrDataByKeyWord("PROSOURCE")) != null)
                {
                    data.SetCodeDesc(projectSource);
                }
                //项目来源（说明）
                if ((data = docItem.GetAttrDataByKeyWord("PROSOURCEREMARKS")) != null)
                {
                    data.SetCodeDesc(projectSourceSupp);
                }
                //合同类型
                if ((data = docItem.GetAttrDataByKeyWord("CONTRACTTYPE")) != null)
                {
                    data.SetCodeDesc(contractType);
                }
                //合同类型（说明）
                if ((data = docItem.GetAttrDataByKeyWord("CONTRACTTYPEREMARKS")) != null)
                {
                    data.SetCodeDesc(contractTypeSupp);
                }
                //预估金额
                if ((data = docItem.GetAttrDataByKeyWord("ESTIMATEMONEY")) != null)
                {
                    data.SetCodeDesc(projectAmount);
                }
                //计划工期
                if ((data = docItem.GetAttrDataByKeyWord("PLANTIME")) != null)
                {
                    data.SetCodeDesc(projectDuration);
                }
                //计划开始时间
                if ((data = docItem.GetAttrDataByKeyWord("PLANSTARTDATE")) != null)
                {
                    data.SetCodeDesc(planStartDate);
                }
                //计划结束时间
                if ((data = docItem.GetAttrDataByKeyWord("PLANENDDATE")) != null)
                {
                    data.SetCodeDesc(planEndDate);
                }
                //项目概况
                if ((data = docItem.GetAttrDataByKeyWord("REMARKS")) != null)
                {
                    data.SetCodeDesc(projectOverview);
                }
                //通信邮箱
                if ((data = docItem.GetAttrDataByKeyWord("EMAIL")) != null)
                {
                    data.SetCodeDesc(email);
                }

                //立项人
                if ((data = docItem.GetAttrDataByKeyWord("WRITEMAN")) != null)
                {
                    data.SetCodeDesc(dbsource.LoginUser.ToString);
                }

                ////保存项目属性，存进数据库
                docItem.AttrDataList.SaveData();

                #endregion

                #region 录入数据进入word表单

                string strDocList = "";//获取附件

                //录入数据进入表单
                Hashtable htUserKeyWord = new Hashtable();

                //if (strPriImport == "一般资料")
                //{
                //    htUserKeyWord.Add("DATA_NOR", "☑");//一般
                //    htUserKeyWord.Add("DTAT_IMP", "□");//重要
                //}
                //else
                //{
                //    htUserKeyWord.Add("DATA_NOR", "□");//一般
                //    htUserKeyWord.Add("DTAT_IMP", "☑");//重要
                //}

                #region 添加勾选项
                #region 勾选选项注释
                //PROTYPE_POWER     //发电工程
                //PROTYPE_ENVI      //环保工程
                //PROTYPE_PUBLIC    //市政工程
                //PROTYPE_HEAT      //热力工程
                //PROTYPE_ESTATE    //地产项目
                //PROTYPE_OTHER     //其他

                //PROTYPE_POWER_COAL    //燃煤发电
                //PROTYPE_POWER_BIOM    //生物质发电
                //PROTYPE_POWER_PLANT   //燃机发电
                //PROTYPE_POWER_DOME    //生活垃圾发电
                //PROTYPE_POWER_OPTO    //光热发电
                //PROTYPE_POWER_PV      //光伏发电
                //PROTYPE_POWER_WIND    //风力发电
                //PROTYPE_POWER_NUCL    //核电工程 
                //PROTYPE_POWER_GEOT    //地热发电
                //PROTYPE_POWER_COGE    //余热发电
                //PROTYPE_POWER_OILGAS  //油·气发电
                //PROTYPE_POWER_GASC    //燃气联合循环发电

                //PROTYPE_ENVI_DESU     //烟气脱硫工程
                //PROTYPE_ENVI_DENI     //烟气脱硝工程
                //PROTYPE_ENVI_DUST     //烟气除尘工程
                //PROTYPE_ENVI_KITC     //厨余垃圾处理工程
                //PROTYPE_ENVI_VENO     //静脉产业园工程
                //PROTYPE_ENVI_WATER    //水处理工程
                //PROTYPE_ENVI_SLUD     //污泥处理工程
                //PROTYPE_ENVI_FORE     //农林废弃物综合利用
                //PROTYPE_ENVI_OTHER    //其他

                //PROTYPE_PUBLIC_MUNI   //市政设施工程
                //PROTYPE_PUBLIC_INDU   //工业设施工程
                //PROTYPE_PUBLIC_CIVIL  //民用设施工程
                //PROTYPE_PUBLIC_PUBL   //公共设施项目工程
                //PROTYPE_PUBLIC_OTHER  //其他

                //PROTYPE_HEAT_URBAN    //城市供热工程
                //PROTYPE_HEAT_INDU     //工业供热工程
                //PROTYPE_HEAT_OTHER    //其他

                //PROTYPE_ESTATE_RESI   //住宅项目
                //PROTYPE_ESTATE_BUSI   //商业项目
                //PROTYPE_ESTATE_COMP   //综合体项目
                //PROTYPE_ESTATE_OTHER  //其他

                ////////////////项目实施性质////////////////
                //PRONATURE_NEW         //新建
                //PRONATURE_EXP         //扩建
                //PRONATURE_RENO        //改造
                //PRONATURE_OTHER       //其他

                ///////////////项目来源/////////////////
                //PROSOURCE_INTE    //华西工程国际销售公司
                //PROSOURCE_DOME    //华西工程国内销售公司
                //PROSOURCE_PPP     //华西工程PPP销售公司
                //PROSOURCE_DESI    //华西工程设计研究院
                //PROSOURCE_OPER    //华西工程运维公司
                //PROSOURCE_SERV    //华西工程项目服务部
                //PROSOURCE_STOCK   //股份公司
                //PROSOURCE_INDU    //华西工业
                //PROSOURCE_OTHER   //其他

                ///////////////合约类型///////////////////
                //CONTRACTTYPE_E    
                //CONTRACTTYPE_P
                //CONTRACTTYPE_C    
                //CONTRACTTYPE_PC    
                //CONTRACTTYPE_EP
                //CONTRACTTYPE_EPC
                //CONTRACTTYPE_EPCC
                //CONTRACTTYPE_PPP
                //CONTRACTTYPE_BOT
                //CONTRACTTYPE_BOO
                //CONTRACTTYPE_OTHER   //其他 

                #endregion

                #region 初始化勾选选项注释

                //////////////////项目类型//////////////////////
                string PROTYPE_POWER = "□";     //发电工程
                string PROTYPE_ENVI = "□";      //环保工程
                string PROTYPE_PUBLIC = "□";    //市政工程
                string PROTYPE_HEAT = "□";      //热力工程
                string PROTYPE_ESTATE = "□";    //地产项目
                string PROTYPE_OTHER = "□";    //其他

                string PROTYPE_POWER_COAL = "□";   //燃煤发电
                string PROTYPE_POWER_BIOM = "□";   //生物质发电
                string PROTYPE_POWER_PLANT = "□";  //燃机发电
                string PROTYPE_POWER_DOME = "□";     //生活垃圾发电
                string PROTYPE_POWER_OPTO = "□";    //光热发电
                string PROTYPE_POWER_PV = "□";     //光伏发电
                string PROTYPE_POWER_WIND = "□";     //风力发电
                string PROTYPE_POWER_NUCL = "□";    //核电工程 
                string PROTYPE_POWER_GEOT = "□";    //地热发电
                string PROTYPE_POWER_COGE = "□";    //余热发电
                string PROTYPE_POWER_OILGAS = "□";   //油·气发电
                string PROTYPE_POWER_GASC = "□";   //燃气联合循环发电

                string PROTYPE_ENVI_DESU = "□";    //烟气脱硫工程
                string PROTYPE_ENVI_DENI = "□";    //烟气脱硝工程
                string PROTYPE_ENVI_DUST = "□";    //烟气除尘工程
                string PROTYPE_ENVI_KITC = "□";     //厨余垃圾处理工程
                string PROTYPE_ENVI_VENO = "□";     //静脉产业园工程
                string PROTYPE_ENVI_WATER = "□";   //水处理工程
                string PROTYPE_ENVI_SLUD = "□";    //污泥处理工程
                string PROTYPE_ENVI_FORE = "□";     //农林废弃物综合利用
                string PROTYPE_ENVI_OTHER = "□";   //其他

                string PROTYPE_PUBLIC_MUNI = "□";   //市政设施工程
                string PROTYPE_PUBLIC_INDU = "□";   //工业设施工程
                string PROTYPE_PUBLIC_CIVIL = "□";   //民用设施工程
                string PROTYPE_PUBLIC_PUBL = "□";   //公共设施项目工程
                string PROTYPE_PUBLIC_OTHER = "□";   //其他

                string PROTYPE_HEAT_URBAN = "□";    //城市供热工程
                string PROTYPE_HEAT_INDU = "□";     //工业供热工程
                string PROTYPE_HEAT_OTHER = "□";    //其他

                string PROTYPE_ESTATE_RESI = "□";  //住宅项目
                string PROTYPE_ESTATE_BUSI = "□";   //商业项目
                string PROTYPE_ESTATE_COMP = "□";   //综合体项目
                string PROTYPE_ESTATE_OTHER = "□";  //其他

                ////////////////项目实施性质////////////////
                string PRONATURE_NEW = "□";         //新建
                string PRONATURE_EXP = "□";         //扩建
                string PRONATURE_RENO = "□";       //改造
                string PRONATURE_OTHER = "□";      //其他

                ///////////////项目来源/////////////////
                string PROSOURCE_INTE = "□";   //华西工程国际销售公司
                string PROSOURCE_DOME = "□";    //华西工程国内销售公司
                string PROSOURCE_PPP = "□";    //华西工程PPP销售公司
                string PROSOURCE_DESI = "□";    //华西工程设计研究院
                string PROSOURCE_OPER = "□";   //华西工程运维公司
                string PROSOURCE_SERV = "□";     //华西工程项目服务部
                string PROSOURCE_STOCK = "□";   //股份公司
                string PROSOURCE_INDU = "□";    //华西工业
                string PROSOURCE_OTHER = "□";   //其他

                ///////////////合约类型///////////////////
                string CONTRACTTYPE_E = "□";
                string CONTRACTTYPE_P = "□";
                string CONTRACTTYPE_C = "□";
                string CONTRACTTYPE_PC = "□";
                string CONTRACTTYPE_EP = "□";
                string CONTRACTTYPE_EPC = "□";
                string CONTRACTTYPE_EPCC = "□";
                string CONTRACTTYPE_PPP = "□";
                string CONTRACTTYPE_BOT = "□";
                string CONTRACTTYPE_BOO = "□";
                string CONTRACTTYPE_OTHER = "□";   //其他 


                string PROTYPE_ENVI_SUPP = ""; //环保工程 补充说明
                string PROTYPE_PUBLIC_SUPP = "";//市政工程 补充说明
                string PROTYPE_HEAT_SUPP = "";//热力工程 补充说明
                string PROTYPE_ESTATE_SUPP = "";//地产项目 补充说明
                string PROTYPE_SUPP = "";//项目类型 补充说明

                string PRONATURE_SUPP = "";//项目实施性质 补充说明
                string PROSOURCE_SUPP = "";//项目来源 补充说明
                string CONTRACTTYPE_SUPP = "";//合约类型 补充说明
                #endregion

                #region 勾选项目类型
                if (projectType == "发电工程")
                {
                    PROTYPE_POWER = "☑";

                    if (projectTypeII == "燃煤发电")
                        PROTYPE_POWER_COAL = "☑";
                    else if (projectTypeII == "生物质发电")
                        PROTYPE_POWER_BIOM = "☑";
                    else if (projectTypeII == "燃机发电")
                        PROTYPE_POWER_PLANT = "☑";
                    else if (projectTypeII == "生活垃圾发电")
                        PROTYPE_POWER_DOME = "☑";
                    else if (projectTypeII == "光热发电")
                        PROTYPE_POWER_OPTO = "☑";
                    else if (projectTypeII == "光伏发电")
                        PROTYPE_POWER_PV = "☑";
                    else if (projectTypeII == "风力发电")
                        PROTYPE_POWER_WIND = "☑";
                    else if (projectTypeII == "核电工程")
                        PROTYPE_POWER_NUCL = "☑";
                    else if (projectTypeII == "地热发电")
                        PROTYPE_POWER_GEOT = "☑";
                    else if (projectTypeII == "余热发电")
                        PROTYPE_POWER_COGE = "☑";
                    else if (projectTypeII == "油·气发电")
                        PROTYPE_POWER_OILGAS = "☑";
                    else if (projectTypeII == "燃气联合循环发电")
                        PROTYPE_POWER_GASC = "☑";
                }
                else if (projectType == "环保工程")
                {
                    PROTYPE_ENVI = "☑";

                    if (projectTypeII == "烟气脱硫工程")
                        PROTYPE_ENVI_DESU = "☑";
                    else if (projectTypeII == "烟气脱硝工程")
                        PROTYPE_ENVI_DENI = "☑";
                    else if (projectTypeII == "烟气除尘工程")
                        PROTYPE_ENVI_DUST = "☑";
                    else if (projectTypeII == "厨余垃圾处理工程")
                        PROTYPE_ENVI_KITC = "☑";
                    else if (projectTypeII == "静脉产业园工程")
                        PROTYPE_ENVI_VENO = "☑";
                    else if (projectTypeII == "水处理工程")
                        PROTYPE_ENVI_WATER = "☑";
                    else if (projectTypeII == "污泥处理工程")
                        PROTYPE_ENVI_SLUD = "☑";
                    else if (projectTypeII == "农林废弃物综合利用")
                        PROTYPE_ENVI_FORE = "☑";
                    else if (projectTypeII == "其他")
                    {
                        PROTYPE_ENVI_OTHER = "☑";
                        PROTYPE_ENVI_SUPP = projectTypeSupp;
                    }

                }
                else if (projectType == "市政工程")
                {
                    PROTYPE_PUBLIC = "☑";

                    if (projectTypeII == "市政设施工程")
                        PROTYPE_PUBLIC_MUNI = "☑";
                    else if (projectTypeII == "工业设施工程")
                        PROTYPE_PUBLIC_INDU = "☑";
                    else if (projectTypeII == "民用设施工程")
                        PROTYPE_PUBLIC_CIVIL = "☑";
                    else if (projectTypeII == "公共设施项目工程")
                        PROTYPE_PUBLIC_PUBL = "☑";
                    else if (projectTypeII == "其他")
                    {
                        PROTYPE_PUBLIC_OTHER = "☑";
                        PROTYPE_PUBLIC_SUPP = projectTypeSupp;
                    }
                }
                else if (projectType == "热力工程")
                {
                    PROTYPE_HEAT = "☑";

                    if (projectTypeII == "城市供热工程")
                        PROTYPE_HEAT_URBAN = "☑";
                    else if (projectTypeII == "工业供热工程")
                        PROTYPE_HEAT_INDU = "☑";
                    else if (projectTypeII == "其他")
                    {
                        PROTYPE_HEAT_OTHER = "☑";
                        PROTYPE_HEAT_SUPP = projectTypeSupp;
                    }

                }
                else if (projectType == "地产项目")
                {
                    PROTYPE_ESTATE = "☑";

                    if (projectTypeII == "住宅项目")
                        PROTYPE_ESTATE_RESI = "☑";
                    else if (projectTypeII == "商业项目")
                        PROTYPE_ESTATE_BUSI = "☑";
                    else if (projectTypeII == "综合体项目")
                        PROTYPE_ESTATE_COMP = "☑";
                    else if (projectTypeII == "其他")
                    {
                        PROTYPE_ESTATE_OTHER = "☑";
                        PROTYPE_ESTATE_SUPP = projectTypeSupp;
                    }

                }
                else if (projectType == "其他")
                {
                    PROTYPE_OTHER = "☑";
                    PROTYPE_SUPP = projectTypeSupp;
                }

                #endregion


                #region 勾选项目实施性质，项目来源和合约类型

                //勾选项目实施性质
                if (projectNature == "新建")
                {
                    PRONATURE_NEW = "☑";
                }
                else if (projectNature == "扩建")
                {
                    PRONATURE_EXP = "☑";
                }
                else if (projectNature == "改造")
                {
                    PRONATURE_RENO = "☑";
                }
                else if (projectNature == "其他")
                {
                    PRONATURE_OTHER = "☑";
                    PRONATURE_SUPP = projectNatureSupp;
                }

                //勾选项目来源
                if (projectSource == "华西工程国际销售公司")
                {
                    PROSOURCE_INTE = "☑";
                }
                else if (projectSource == "华西工程国内销售公司")
                {
                    PROSOURCE_DOME = "☑";
                }
                else if (projectSource == "华西工程PPP销售公司")
                {
                    PROSOURCE_PPP = "☑";
                }
                else if (projectSource == "华西工程设计研究院")
                {
                    PROSOURCE_DESI = "☑";
                }
                else if (projectSource == "华西工程运维公司")
                {
                    PROSOURCE_OPER = "☑";
                }
                else if (projectSource == "华西工程项目服务部")
                {
                    PROSOURCE_SERV = "☑";
                }
                else if (projectSource == "股份公司")
                {
                    PROSOURCE_STOCK = "☑";
                }
                else if (projectSource == "华西工业")
                {
                    PROSOURCE_INDU = "☑";
                }
                else if (projectSource == "其他")
                {
                    PROSOURCE_OTHER = "☑";
                    PROSOURCE_SUPP = projectSourceSupp;
                }

                //勾选合约类型
                if (contractType == "E")
                {
                    CONTRACTTYPE_E = "☑";
                }
                else if (contractType == "P")
                {
                    CONTRACTTYPE_P = "☑";
                }
                else if (contractType == "C")
                {
                    CONTRACTTYPE_C = "☑";
                }
                else if (contractType == "PC")
                {
                    CONTRACTTYPE_PC = "☑";
                }
                else if (contractType == "EP")
                {
                    CONTRACTTYPE_EP = "☑";
                }
                else if (contractType == "EPC")
                {
                    CONTRACTTYPE_EPC = "☑";
                }
                else if (contractType == "EPCC")
                {
                    CONTRACTTYPE_EPCC = "☑";
                }
                else if (contractType == "PPP")
                {
                    CONTRACTTYPE_PPP = "☑";
                }
                else if (contractType == "BOT")
                {
                    CONTRACTTYPE_BOT = "☑";
                }
                else if (contractType == "BOO")
                {
                    CONTRACTTYPE_BOO = "☑";
                }
                else if (contractType == "其他")
                {
                    CONTRACTTYPE_OTHER = "☑";
                    CONTRACTTYPE_SUPP = contractTypeSupp;
                }

                #endregion


                #region 添加勾选项到哈希表
                htUserKeyWord.Add("PROTYPE_POWER", PROTYPE_POWER);//发电工程
                htUserKeyWord.Add("PROTYPE_ENVI", PROTYPE_ENVI);//环保工程
                htUserKeyWord.Add("PROTYPE_PUBLIC", PROTYPE_PUBLIC);//环保工程
                htUserKeyWord.Add("PROTYPE_HEAT", PROTYPE_HEAT);//热力工程
                htUserKeyWord.Add("PROTYPE_ESTATE", PROTYPE_ESTATE);//地产项目
                htUserKeyWord.Add("PROTYPE_OTHER", PROTYPE_OTHER);//其他

                htUserKeyWord.Add("PROTYPE_POWER_COAL", PROTYPE_POWER_COAL);//燃煤发电
                htUserKeyWord.Add("PROTYPE_POWER_BIOM", PROTYPE_POWER_BIOM);//生物质发电
                htUserKeyWord.Add("PROTYPE_POWER_PLANT", PROTYPE_POWER_PLANT);//燃机发电
                htUserKeyWord.Add("PROTYPE_POWER_DOME", PROTYPE_POWER_DOME);//生活垃圾发电
                htUserKeyWord.Add("PROTYPE_POWER_OPTO", PROTYPE_POWER_OPTO);//光热发电
                htUserKeyWord.Add("PROTYPE_POWER_PV", PROTYPE_POWER_PV);//光伏发电
                htUserKeyWord.Add("PROTYPE_POWER_WIND", PROTYPE_POWER_WIND);//风力发电
                htUserKeyWord.Add("PROTYPE_POWER_NUCL", PROTYPE_POWER_NUCL);//核电工程 
                htUserKeyWord.Add("PROTYPE_POWER_GEOT", PROTYPE_POWER_GEOT);//地热发电
                htUserKeyWord.Add("PROTYPE_POWER_COGE", PROTYPE_POWER_COGE);//余热发电
                htUserKeyWord.Add("PROTYPE_POWER_OILGAS", PROTYPE_POWER_OILGAS);//油·气发电
                htUserKeyWord.Add("PROTYPE_POWER_GASC", PROTYPE_POWER_GASC);//燃气联合循环发电

                htUserKeyWord.Add("PROTYPE_ENVI_DESU", PROTYPE_ENVI_DESU);//烟气脱硫工程
                htUserKeyWord.Add("PROTYPE_ENVI_DENI", PROTYPE_ENVI_DENI);//烟气脱硝工程
                htUserKeyWord.Add("PROTYPE_ENVI_DUST", PROTYPE_ENVI_DUST);//烟气除尘工程
                htUserKeyWord.Add("PROTYPE_ENVI_KITC", PROTYPE_ENVI_KITC);//厨余垃圾处理工程
                htUserKeyWord.Add("PROTYPE_ENVI_VENO", PROTYPE_ENVI_VENO);//静脉产业园工程
                htUserKeyWord.Add("PROTYPE_ENVI_WATER", PROTYPE_ENVI_WATER);//水处理工程
                htUserKeyWord.Add("PROTYPE_ENVI_SLUD", PROTYPE_ENVI_SLUD);//污泥处理工程
                htUserKeyWord.Add("PROTYPE_ENVI_FORE", PROTYPE_ENVI_FORE);//农林废弃物综合利用
                htUserKeyWord.Add("PROTYPE_ENVI_OTHER", PROTYPE_ENVI_OTHER);//其他

                htUserKeyWord.Add("PROTYPE_PUBLIC_MUNI", PROTYPE_PUBLIC_MUNI);//市政设施工程
                htUserKeyWord.Add("PROTYPE_PUBLIC_INDU", PROTYPE_PUBLIC_INDU);//工业设施工程
                htUserKeyWord.Add("PROTYPE_PUBLIC_CIVIL", PROTYPE_PUBLIC_CIVIL);//民用设施工程
                htUserKeyWord.Add("PROTYPE_PUBLIC_PUBL", PROTYPE_PUBLIC_PUBL);//公共设施项目工程
                htUserKeyWord.Add("PROTYPE_PUBLIC_OTHER", PROTYPE_PUBLIC_OTHER);//其他

                htUserKeyWord.Add("PROTYPE_HEAT_URBAN", PROTYPE_HEAT_URBAN);//城市供热工程
                htUserKeyWord.Add("PROTYPE_HEAT_INDU", PROTYPE_HEAT_INDU);//工业供热工程
                htUserKeyWord.Add("PROTYPE_HEAT_OTHER", PROTYPE_HEAT_OTHER);//其他

                htUserKeyWord.Add("PROTYPE_ESTATE_RESI", PROTYPE_ESTATE_RESI);//住宅项目
                htUserKeyWord.Add("PROTYPE_ESTATE_BUSI", PROTYPE_ESTATE_BUSI);//商业项目
                htUserKeyWord.Add("PROTYPE_ESTATE_COMP", PROTYPE_ESTATE_COMP);//综合体项目
                htUserKeyWord.Add("PROTYPE_ESTATE_OTHER", PROTYPE_ESTATE_OTHER);//其他

                ////////////////项目实施性质////////////////
                htUserKeyWord.Add("PRONATURE_NEW", PRONATURE_NEW);//新建
                htUserKeyWord.Add("PRONATURE_EXP", PRONATURE_EXP);//扩建
                htUserKeyWord.Add("PRONATURE_RENO", PRONATURE_RENO);//改造
                htUserKeyWord.Add("PRONATURE_OTHER", PRONATURE_OTHER);//其他

                ///////////////项目来源/////////////////
                htUserKeyWord.Add("PROSOURCE_INTE", PROSOURCE_INTE);//华西工程国际销售公司
                htUserKeyWord.Add("PROSOURCE_DOME", PROSOURCE_DOME);//华西工程国内销售公司
                htUserKeyWord.Add("PROSOURCE_PPP", PROSOURCE_PPP);//华西工程PPP销售公司
                htUserKeyWord.Add("PROSOURCE_DESI", PROSOURCE_DESI);//华西工程设计研究院
                htUserKeyWord.Add("PROSOURCE_OPER", PROSOURCE_OPER);//华西工程运维公司
                htUserKeyWord.Add("PROSOURCE_SERV", PROSOURCE_SERV);//华西工程项目服务部
                htUserKeyWord.Add("PROSOURCE_STOCK", PROSOURCE_STOCK);//股份公司
                htUserKeyWord.Add("PROSOURCE_INDU", PROSOURCE_INDU);//华西工业
                htUserKeyWord.Add("PROSOURCE_OTHER", PROSOURCE_OTHER);//其他


                ///////////////合约类型///////////////////
                htUserKeyWord.Add("CONTRACTTYPE_E", CONTRACTTYPE_E);//E
                htUserKeyWord.Add("CONTRACTTYPE_P", CONTRACTTYPE_P);//P
                htUserKeyWord.Add("CONTRACTTYPE_C", CONTRACTTYPE_C);//C
                htUserKeyWord.Add("CONTRACTTYPE_PC", CONTRACTTYPE_PC);//PC
                htUserKeyWord.Add("CONTRACTTYPE_EP", CONTRACTTYPE_EP);//EP
                htUserKeyWord.Add("CONTRACTTYPE_EPC", CONTRACTTYPE_EPC);//EPC
                htUserKeyWord.Add("CONTRACTTYPE_EPCC", CONTRACTTYPE_EPCC);//EPCC
                htUserKeyWord.Add("CONTRACTTYPE_PPP", CONTRACTTYPE_PPP);//PPP
                htUserKeyWord.Add("CONTRACTTYPE_BOT", CONTRACTTYPE_BOT);//BOT
                htUserKeyWord.Add("CONTRACTTYPE_BOO", CONTRACTTYPE_BOO);//BOO
                htUserKeyWord.Add("CONTRACTTYPE_OTHER", CONTRACTTYPE_OTHER);//其他 


                htUserKeyWord.Add("PROTYPE_ENVI_SUPP", PROTYPE_ENVI_SUPP);//环保工程 补充说明 
                htUserKeyWord.Add("PROTYPE_PUBLIC_SUPP", PROTYPE_PUBLIC_SUPP);//市政工程 补充说明
                htUserKeyWord.Add("PROTYPE_HEAT_SUPP", PROTYPE_HEAT_SUPP);//热力工程 补充说明
                htUserKeyWord.Add("PROTYPE_ESTATE_SUPP", PROTYPE_ESTATE_SUPP);//地产项目 补充说明
                htUserKeyWord.Add("PROTYPE_SUPP", PROTYPE_SUPP);//项目类型 补充说明

                htUserKeyWord.Add("PRONATURE_SUPP", PRONATURE_SUPP);//项目实施性质 补充说明
                htUserKeyWord.Add("PROSOURCE_SUPP", PROSOURCE_SUPP);//项目来源 补充说明
                htUserKeyWord.Add("CONTRACTTYPE_SUPP", CONTRACTTYPE_SUPP);//合约类型 补充说明

                #endregion
                #endregion

                if (!string.IsNullOrEmpty(communicationCode) && !string.IsNullOrEmpty(intcommunicationCode))
                { communicationCode = communicationCode +","+ intcommunicationCode; }
                else if (string.IsNullOrEmpty(communicationCode) && !string.IsNullOrEmpty(intcommunicationCode))
                {
                    communicationCode = intcommunicationCode;
                }

                htUserKeyWord.Add("PROCODE", projectCode);//项目代码
                htUserKeyWord.Add("PRONAME", projectDescCN);//项目名称（中文）
                htUserKeyWord.Add("PROENGLISH", projectDescEN);//项目名称（英文）
                htUserKeyWord.Add("ABBREVIATION", projectShortDescCN);//项目简称（中文）
                htUserKeyWord.Add("ABBREVIATIONENG", projectShortDescEN);//项目简称（英文）
                htUserKeyWord.Add("COMMUNICATIONCODE", communicationCode);//通信代码
                htUserKeyWord.Add("INTCOMMUNICATIONCODE", intcommunicationCode);//国际通信代码
                htUserKeyWord.Add("OWNER", unintName);//业主
                htUserKeyWord.Add("PROADDRESS", projectAddr);//项目实施地址
                htUserKeyWord.Add("ESTIMATEMONEY", projectAmount); //预估金额
                htUserKeyWord.Add("PLANTIME", projectDuration);//计划工期
                htUserKeyWord.Add("PLANSTARTDATE", planStartDate);//计划开始时间
                htUserKeyWord.Add("PLANENDDATE", planEndDate);//计划结束时间
                htUserKeyWord.Add("REMARKS", projectOverview);//项目概况
                htUserKeyWord.Add("EMAIL", email);
                //htUserKeyWord.Add("FORM", strFormQuantity);


                //this.Cursor = Cursors.WaitCursor;
                string workingPath = m_Project.dBSource.LoginUser.WorkingPath;
                //AttrData ad = m_Project.GetAttrDataByKeyWord("ISSAVE");
                //if (ad != null)
                //{
                //    ad.SetCodeDesc("");
                //}
                //m_Project.AttrDataList.SaveData();


                try
                {
                    //上传下载文档
                    string exchangfilename = "项目立项单";

                    //获取网站路径
                    string sPath = System.Web.HttpContext.Current.Server.MapPath("/ISO/HXEPC/");

                    //获取模板文件路径
                    string modelFileName = sPath + exchangfilename + ".doc";

                    //获取即将生成的联系单文件路径
                    string locFileName = docItem.FullPathFile;

                    //FTPFactory factory = m_Project.Storage.FTP ?? new FTPFactory(m_Project.Storage);
                    //string locFileName = m_Project.dBSource.LoginUser.WorkingPath + docItem.Code + ".doc";
                    //factory.download(@"\ISO\" + exchangfilename + ".doc", locFileName, false);
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
                    //base.DialogResult = DialogResult.OK;
                    //this.Cursor = Cursors.Default;
                    //base.Close();
                    //CommonFunction.InsertDocListAndOpenDoc(this.m_DocList, docItem);

                    if (string.IsNullOrEmpty(strDocList))
                    {
                        strDocList = docItem.KeyWord;
                    }
                    else
                    {
                        strDocList = docItem.KeyWord + "," + strDocList;
                    }

                    ////增加管理员
                    //AVEVA.CDMS.Server.Group gp = dbsource.GetGroupByName("AdminGroup");
                    //if (gp != null)
                    //{
                    //    m_NewProject.groupList[0].AddGroup(gp);
                    //    m_NewProject.groupList[0].Modify();
                    //}

                    //这里刷新数据源，否则创建流程的时候获取不了专业字符串
                    DBSourceController.RefreshDBSource(sid);

                    reJo.success = true;
                    reJo.data = new JArray(new JObject(new JProperty("ProjectKeyword", docItem.Project.KeyWord),
                        new JProperty("DocKeyword", docItem.KeyWord), new JProperty("DocList", strDocList)));
                    return reJo.Value;
                }
                catch { }
                #endregion


                ////增加领导组
                //AVEVA.CDMS.Server.Group gp = dbsource.GetGroupByName("Manager");
                //if (gp != null)
                //{
                //    m_NewProject.groupList[0].AddGroup(gp);
                //    m_NewProject.groupList[0].Modify();
                //}

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

        //立项函数，立项时生成立项单
        public static ExReJObject CreateRootProjectX2(DBSource dbsource, JArray jaAttr, string sid)
        {
            ExReJObject reJo = new ExReJObject();

            try
            {

                #region 获取项目参数项目

                //JArray jaAttr = (JArray)JsonConvert.DeserializeObject(projectAttrJson);

                //获取项目参数项目
                string projectCode = "", projectDescEN = "", projectDescCN = "",
                    projectShortDescEN = "", projectShortDescCN = "", unintName = "", //projectNo = "",
                    projectAddr = "", projectOverview = "", projectType = "",
                    projectTypeII = "", projectTypeSupp = "", projectNature = "",
                    projectNatureSupp = "", projectSource = "", projectSourceSupp = "",
                    contractType = "", contractTypeSupp = "", projectAmount = "",
                    projectDuration = "", planStartDate = "", planEndDate = "",
                    communicationCode = "", email = "";

                //合同号
                string projectNomber = "";
                //备注
                string Remarks = "";

                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    //获取项目代码
                    if (strName == "projectCode") projectCode = strValue.Trim();

                    //获取项目英文名称
                    else if (strName == "projectDescEN") projectDescEN = strValue;

                    //获取项目中文名称
                    else if (strName == "projectDescCN") projectDescCN = strValue;

                    //获取项目英文简称
                    else if (strName == "projectShortDescEN") projectShortDescEN = strValue;

                    //获取项目中文简称
                    else if (strName == "projectShortDescCN") projectShortDescCN = strValue;

                    //获取业主
                    else if (strName == "unintName") unintName = strValue;

                    //获取项目地址
                    else if (strName == "projectAddr") projectAddr = strValue;

                    //获取项目概况
                    else if (strName == "projectOverview") projectOverview = strValue;

                    //获取项目类型
                    else if (strName == "projectType") projectType = strValue;

                    //获取项目二级类型
                    else if (strName == "projectTypeII") projectTypeII = strValue;

                    //获取项目类型补充说明
                    else if (strName == "projectTypeSupp") projectTypeSupp = strValue;

                    //获取合同额
                    else if (strName == "projectAmount") projectAmount = strValue;

                    //获取项目实施性质
                    else if (strName == "projectNature") projectNature = strValue;

                    //项目实施性质补充说明
                    else if (strName == "projectNatureSupp") projectNatureSupp = strValue;

                    //获取项目来源
                    else if (strName == "projectSource") projectSource = strValue;

                    //获取项目来源补充说明
                    else if (strName == "projectSourceSupp") projectSourceSupp = strValue;

                    //获取合同类型
                    else if (strName == "contractType") contractType = strValue;

                    //获取合同类型补充
                    else if (strName == "contractTypeSupp") contractTypeSupp = strValue;

                    //获取预估金额
                    else if (strName == "projectAmount") projectAmount = strValue;

                    //获取计划工期
                    else if (strName == "projectDuration") projectDuration = strValue;

                    //获取计划开始时间
                    else if (strName == "planStartDate") planStartDate = strValue;

                    //获取计划结束时间
                    else if (strName == "planEndDate") planEndDate = strValue;

                    //通信代码
                    else if (strName == "communicationCode") communicationCode = strValue;

                    //通信邮箱
                    else if (strName == "email") email = strValue;

                }

                string projectDesc = string.IsNullOrEmpty(projectDescEN) ?
                    projectDescCN
                    : projectDescCN + " " + projectDescEN;

                if (string.IsNullOrEmpty(projectCode))
                {
                    reJo.msg = "请填写项目编号！";
                    return reJo;
                }
                else if (string.IsNullOrEmpty(projectDesc))
                {
                    reJo.msg = "请填写项目名称！";
                    return reJo;
                }

                #endregion


                //  根据名称查找项目模板(根目录)对象
                List<TempDefn> tempDefnByCode = dbsource.GetTempDefnByCode("HXNY_DOCUMENTSYSTEM");
                TempDefn mTempDefn = (tempDefnByCode != null) ? tempDefnByCode[0] : null;
                if (mTempDefn == null)
                {
                    reJo.msg = "欲创建的项目关联的模板不存在！不能完成创建";
                    return reJo;
                }
                else
                {

                    //获取DBSource的虚拟Local目录
                    Project m_LocalProject = dbsource.NewProject(enProjectType.Local);

                    //查找项目是否已经创建
                    Project findProj = dbsource.RootLocalProjectList.Find(itemProj => itemProj.Code == projectCode);
                    if (findProj != null)
                    {
                        reJo.msg = "项目[" + projectCode + "]已存在！不能完成创建";
                        return reJo;
                    }

                    //foreach (Project findProj in m_Project.AllProjectList)//.ChildProjectList)
                    //{
                    //    if (findProj != null && findProj.Code == projectCode)
                    //    {
                    //        reJo.msg = "项目[" + projectCode + "]已存在！不能完成创建";
                    //        return reJo;
                    //    }
                    //}

                    //创建项目
                    Project m_NewProject = m_LocalProject.NewProject(projectCode, projectDesc, null, mTempDefn);


                    if (m_NewProject == null)
                    {
                        reJo.msg = "新建项目失败";
                        return reJo;

                    }
                    else
                    {



                        #region 设置项目属性

                        //设置项目属性
                        //foreach (JObject joAttr in jaAttr)
                        //{
                        //    string strName = joAttr["name"].ToString();
                        //    string strValue = joAttr["value"].ToString();

                        //    AttrData data;
                        //    if ((data = m_NewProject.GetAttrDataByKeyWord(strName)) != null)
                        //    {
                        //        data.SetCodeDesc(strValue);
                        //    }

                        //}

                        AttrData data;
                        //项目名称（中文）
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PRO_NAME")) != null)
                        {
                            data.SetCodeDesc(projectDescCN);
                        }
                        //项目名称（英文）
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PRO_ENGLISH")) != null)
                        {
                            data.SetCodeDesc(projectDescEN);
                        }
                        //项目简称（中文）
                        if ((data = m_NewProject.GetAttrDataByKeyWord("ABBREVIATION")) != null)
                        {
                            data.SetCodeDesc(projectShortDescCN);
                        }
                        //项目简称（英文）
                        if ((data = m_NewProject.GetAttrDataByKeyWord("ABBREVIATIONENG")) != null)
                        {
                            data.SetCodeDesc(projectShortDescEN);
                        }

                        //通信代码
                        if ((data = m_NewProject.GetAttrDataByKeyWord("COMMUNICATIONCODE")) != null)
                        {
                            data.SetCodeDesc(communicationCode);
                        }
                        //业主
                        if ((data = m_NewProject.GetAttrDataByKeyWord("OWNER")) != null)
                        {
                            data.SetCodeDesc(unintName);
                        }
                        //项目实施地址
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PROADDRESS")) != null)
                        {
                            data.SetCodeDesc(projectAddr);
                        }
                        //项目实施性质
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PRONATURE")) != null)
                        {
                            data.SetCodeDesc(projectNature);
                        }
                        //项目实施性质（说明）
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PRONATUREREMARKS")) != null)
                        {
                            data.SetCodeDesc(projectNatureSupp);
                        }
                        //项目类型
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PROTYPE")) != null)
                        {
                            data.SetCodeDesc(projectType);
                        }
                        //项目分类(二级)
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PROTYPECLASS")) != null)
                        {
                            data.SetCodeDesc(projectTypeII);
                        }

                        //项目类型（说明）
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PROTYPEREMARKS")) != null)
                        {
                            data.SetCodeDesc(projectTypeSupp);
                        }
                        //项目来源
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PROSOURCE")) != null)
                        {
                            data.SetCodeDesc(projectSource);
                        }
                        //项目来源（说明）
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PROSOURCEREMARKS")) != null)
                        {
                            data.SetCodeDesc(projectSourceSupp);
                        }
                        //合同类型
                        if ((data = m_NewProject.GetAttrDataByKeyWord("CONTRACTTYPE")) != null)
                        {
                            data.SetCodeDesc(contractType);
                        }
                        //合同类型（说明）
                        if ((data = m_NewProject.GetAttrDataByKeyWord("CONTRACTTYPEREMARKS")) != null)
                        {
                            data.SetCodeDesc(contractTypeSupp);
                        }
                        //预估金额
                        if ((data = m_NewProject.GetAttrDataByKeyWord("ESTIMATEMONEY")) != null)
                        {
                            data.SetCodeDesc(projectAmount);
                        }
                        //计划工期
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PLANTIME")) != null)
                        {
                            data.SetCodeDesc(projectDuration);
                        }
                        //计划开始时间
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PLANSTARTDATE")) != null)
                        {
                            data.SetCodeDesc(planStartDate);
                        }
                        //计划结束时间
                        if ((data = m_NewProject.GetAttrDataByKeyWord("PLANENDDATE")) != null)
                        {
                            data.SetCodeDesc(planEndDate);
                        }
                        //项目概况
                        if ((data = m_NewProject.GetAttrDataByKeyWord("REMARKS")) != null)
                        {
                            data.SetCodeDesc(projectOverview);
                        }
                        //通信邮箱
                        if ((data = m_NewProject.GetAttrDataByKeyWord("EMAIL")) != null)
                        {
                            data.SetCodeDesc(email);
                        }

                        //立项人
                        if ((data = m_NewProject.GetAttrDataByKeyWord("WRITEMAN")) != null)
                        {
                            data.SetCodeDesc(dbsource.LoginUser.ToString);
                        }

                        ////保存项目属性，存进数据库
                        m_NewProject.AttrDataList.SaveData();

                        #endregion


                        //#region 创建“通信文件”目录
                        //Project CommProject = CreateChildProject(m_NewProject, "通信文件", "", "PRO_COMMUNICATION");
                        //if (CommProject != null)
                        //{
                        //    Project CommProjectItem = CreateChildProject(CommProject, "红头文", "", "COM_COMTYPE");
                        //    if (CommProjectItem != null)
                        //    {
                        //        Project SubDocProject = CreateChildProject(CommProjectItem, "发文", "", "COM_SUBDOCUMENT");
                        //        SubDocProject = CreateChildProject(CommProjectItem, "收文", "", "COM_SUBDOCUMENT");
                        //    }
                        //    CommProjectItem = CreateChildProject(CommProject, "会议纪要", "", "COM_COMTYPE");
                        //    if (CommProjectItem != null)
                        //    {
                        //        Project SubDocProject = CreateChildProject(CommProjectItem, "发文", "", "COM_SUBDOCUMENT");
                        //        SubDocProject = CreateChildProject(CommProjectItem, "收文", "", "COM_SUBDOCUMENT");
                        //    }
                        //    CommProjectItem = CreateChildProject(CommProject, "内部工作联系单", "", "COM_COMTYPE");
                        //    if (CommProjectItem != null)
                        //    {
                        //        Project SubDocProject = CreateChildProject(CommProjectItem, "发文", "", "COM_SUBDOCUMENT");
                        //        SubDocProject = CreateChildProject(CommProjectItem, "收文", "", "COM_SUBDOCUMENT");
                        //    }
                        //    CommProjectItem = CreateChildProject(CommProject, "外部工作联系单", "", "COM_COMTYPE");
                        //    if (CommProjectItem != null)
                        //    {
                        //        Project SubDocProject = CreateChildProject(CommProjectItem, "发文", "", "COM_SUBDOCUMENT");
                        //        SubDocProject = CreateChildProject(CommProjectItem, "收文", "", "COM_SUBDOCUMENT");
                        //    }
                        //    CommProjectItem = CreateChildProject(CommProject, "文件传递单", "", "COM_COMTYPE");
                        //    if (CommProjectItem != null)
                        //    {
                        //        Project SubDocProject = CreateChildProject(CommProjectItem, "发文", "", "COM_SUBDOCUMENT");
                        //        SubDocProject = CreateChildProject(CommProjectItem, "收文", "", "COM_SUBDOCUMENT");
                        //    }
                        //    CommProjectItem = CreateChildProject(CommProject, "信函", "", "COM_COMTYPE");
                        //    if (CommProjectItem != null)
                        //    {
                        //        Project SubDocProject = CreateChildProject(CommProjectItem, "发文", "", "COM_SUBDOCUMENT");
                        //        SubDocProject = CreateChildProject(CommProjectItem, "收文", "", "COM_SUBDOCUMENT");
                        //    }
                        //}
                        //#endregion

                        //#region 创建“非通信文件”目录
                        //Project NonCommProject = CreateChildProject(m_NewProject, "非通信文件", "", "PRO_NONCOMMUNICATION");
                        //if (NonCommProject != null)
                        //{
                        //    Project NonCommProjectItem = CreateChildProject(NonCommProject, "调试管理", "", "PRO_NOCOMTYPE");
                        //    NonCommProjectItem = CreateChildProject(NonCommProject, "商务管理_国际项目", "", "PRO_NOCOMTYPE");
                        //    NonCommProjectItem = CreateChildProject(NonCommProject, "商务管理_国内项目", "", "PRO_NOCOMTYPE");
                        //    NonCommProjectItem = CreateChildProject(NonCommProject, "设备管理", "", "PRO_NOCOMTYPE");
                        //    NonCommProjectItem = CreateChildProject(NonCommProject, "施工管理", "", "PRO_NOCOMTYPE");
                        //    NonCommProjectItem = CreateChildProject(NonCommProject, "项目达标创优", "", "PRO_NOCOMTYPE");
                        //    NonCommProjectItem = CreateChildProject(NonCommProject, "项目服务", "", "PRO_NOCOMTYPE");
                        //    NonCommProjectItem = CreateChildProject(NonCommProject, "项目管理", "", "PRO_NOCOMTYPE");
                        //    NonCommProjectItem = CreateChildProject(NonCommProject, "项目验收", "", "PRO_NOCOMTYPE");
                        //    NonCommProjectItem = CreateChildProject(NonCommProject, "运维管理", "", "PRO_NOCOMTYPE");
                        //}
                        //#endregion


                        #region 根据立项单模板，生成立项单文档

                        //获取立项单文档所在的目录
                        Project m_Project = m_NewProject;

                        List<TempDefn> docTempDefnByCode = m_Project.dBSource.GetTempDefnByCode("EXCHANGEDOC");
                        TempDefn docTempDefn = (docTempDefnByCode != null && docTempDefnByCode.Count > 0) ? docTempDefnByCode[0] : null;
                        if (docTempDefn == null)
                        {
                            reJo.msg = "没有与其相关的模板管理，创建无法正常完成";
                            return reJo;
                        }

                        IEnumerable<string> source = from docx in m_Project.DocList select docx.Code;
                        string filename = projectDesc + "项目立项单";
                        if (source.Contains<string>(filename))
                        {
                            for (int i = 1; i < 0x3e8; i++)
                            {
                                filename = projectDesc + "项目立项单" + i.ToString();
                                if (!source.Contains<string>(filename))
                                {
                                    reJo.msg = "新建项目立项单出错！";
                                    return reJo;
                                }
                            }
                        }

                        //文档名称
                        Doc item = m_Project.NewDoc(filename + ".doc", filename, "", docTempDefn);
                        if (item == null)
                        {
                            reJo.msg = "新建项目立项单出错！";
                            return reJo;
                        }

                        string strDocList = "";//获取附件

                        //录入数据进入表单
                        Hashtable htUserKeyWord = new Hashtable();

                        //if (strPriImport == "一般资料")
                        //{
                        //    htUserKeyWord.Add("DATA_NOR", "☑");//一般
                        //    htUserKeyWord.Add("DTAT_IMP", "□");//重要
                        //}
                        //else
                        //{
                        //    htUserKeyWord.Add("DATA_NOR", "□");//一般
                        //    htUserKeyWord.Add("DTAT_IMP", "☑");//重要
                        //}

                        #region 添加勾选项
                        #region 勾选选项注释
                        //PROTYPE_POWER     //发电工程
                        //PROTYPE_ENVI      //环保工程
                        //PROTYPE_PUBLIC    //市政工程
                        //PROTYPE_HEAT      //热力工程
                        //PROTYPE_ESTATE    //地产项目
                        //PROTYPE_OTHER     //其他

                        //PROTYPE_POWER_COAL    //燃煤发电
                        //PROTYPE_POWER_BIOM    //生物质发电
                        //PROTYPE_POWER_PLANT   //燃机发电
                        //PROTYPE_POWER_DOME    //生活垃圾发电
                        //PROTYPE_POWER_OPTO    //光热发电
                        //PROTYPE_POWER_PV      //光伏发电
                        //PROTYPE_POWER_WIND    //风力发电
                        //PROTYPE_POWER_NUCL    //核电工程 
                        //PROTYPE_POWER_GEOT    //地热发电
                        //PROTYPE_POWER_COGE    //余热发电
                        //PROTYPE_POWER_OILGAS  //油·气发电
                        //PROTYPE_POWER_GASC    //燃气联合循环发电

                        //PROTYPE_ENVI_DESU     //烟气脱硫工程
                        //PROTYPE_ENVI_DENI     //烟气脱硝工程
                        //PROTYPE_ENVI_DUST     //烟气除尘工程
                        //PROTYPE_ENVI_KITC     //厨余垃圾处理工程
                        //PROTYPE_ENVI_VENO     //静脉产业园工程
                        //PROTYPE_ENVI_WATER    //水处理工程
                        //PROTYPE_ENVI_SLUD     //污泥处理工程
                        //PROTYPE_ENVI_FORE     //农林废弃物综合利用
                        //PROTYPE_ENVI_OTHER    //其他

                        //PROTYPE_PUBLIC_MUNI   //市政设施工程
                        //PROTYPE_PUBLIC_INDU   //工业设施工程
                        //PROTYPE_PUBLIC_CIVIL  //民用设施工程
                        //PROTYPE_PUBLIC_PUBL   //公共设施项目工程
                        //PROTYPE_PUBLIC_OTHER  //其他

                        //PROTYPE_HEAT_URBAN    //城市供热工程
                        //PROTYPE_HEAT_INDU     //工业供热工程
                        //PROTYPE_HEAT_OTHER    //其他

                        //PROTYPE_ESTATE_RESI   //住宅项目
                        //PROTYPE_ESTATE_BUSI   //商业项目
                        //PROTYPE_ESTATE_COMP   //综合体项目
                        //PROTYPE_ESTATE_OTHER  //其他

                        ////////////////项目实施性质////////////////
                        //PRONATURE_NEW         //新建
                        //PRONATURE_EXP         //扩建
                        //PRONATURE_RENO        //改造
                        //PRONATURE_OTHER       //其他

                        ///////////////项目来源/////////////////
                        //PROSOURCE_INTE    //华西工程国际销售公司
                        //PROSOURCE_DOME    //华西工程国内销售公司
                        //PROSOURCE_PPP     //华西工程PPP销售公司
                        //PROSOURCE_DESI    //华西工程设计研究院
                        //PROSOURCE_OPER    //华西工程运维公司
                        //PROSOURCE_SERV    //华西工程项目服务部
                        //PROSOURCE_STOCK   //股份公司
                        //PROSOURCE_INDU    //华西工业
                        //PROSOURCE_OTHER   //其他

                        ///////////////合约类型///////////////////
                        //CONTRACTTYPE_E    
                        //CONTRACTTYPE_P
                        //CONTRACTTYPE_C    
                        //CONTRACTTYPE_PC    
                        //CONTRACTTYPE_EP
                        //CONTRACTTYPE_EPC
                        //CONTRACTTYPE_EPCC
                        //CONTRACTTYPE_PPP
                        //CONTRACTTYPE_BOT
                        //CONTRACTTYPE_BOO
                        //CONTRACTTYPE_OTHER   //其他 

                        #endregion

                        #region 初始化勾选选项注释

                        //////////////////项目类型//////////////////////
                        string PROTYPE_POWER = "□";     //发电工程
                        string PROTYPE_ENVI = "□";      //环保工程
                        string PROTYPE_PUBLIC = "□";    //市政工程
                        string PROTYPE_HEAT = "□";      //热力工程
                        string PROTYPE_ESTATE = "□";    //地产项目
                        string PROTYPE_OTHER = "□";    //其他

                        string PROTYPE_POWER_COAL = "□";   //燃煤发电
                        string PROTYPE_POWER_BIOM = "□";   //生物质发电
                        string PROTYPE_POWER_PLANT = "□";  //燃机发电
                        string PROTYPE_POWER_DOME = "□";     //生活垃圾发电
                        string PROTYPE_POWER_OPTO = "□";    //光热发电
                        string PROTYPE_POWER_PV = "□";     //光伏发电
                        string PROTYPE_POWER_WIND = "□";     //风力发电
                        string PROTYPE_POWER_NUCL = "□";    //核电工程 
                        string PROTYPE_POWER_GEOT = "□";    //地热发电
                        string PROTYPE_POWER_COGE = "□";    //余热发电
                        string PROTYPE_POWER_OILGAS = "□";   //油·气发电
                        string PROTYPE_POWER_GASC = "□";   //燃气联合循环发电

                        string PROTYPE_ENVI_DESU = "□";    //烟气脱硫工程
                        string PROTYPE_ENVI_DENI = "□";    //烟气脱硝工程
                        string PROTYPE_ENVI_DUST = "□";    //烟气除尘工程
                        string PROTYPE_ENVI_KITC = "□";     //厨余垃圾处理工程
                        string PROTYPE_ENVI_VENO = "□";     //静脉产业园工程
                        string PROTYPE_ENVI_WATER = "□";   //水处理工程
                        string PROTYPE_ENVI_SLUD = "□";    //污泥处理工程
                        string PROTYPE_ENVI_FORE = "□";     //农林废弃物综合利用
                        string PROTYPE_ENVI_OTHER = "□";   //其他

                        string PROTYPE_PUBLIC_MUNI = "□";   //市政设施工程
                        string PROTYPE_PUBLIC_INDU = "□";   //工业设施工程
                        string PROTYPE_PUBLIC_CIVIL = "□";   //民用设施工程
                        string PROTYPE_PUBLIC_PUBL = "□";   //公共设施项目工程
                        string PROTYPE_PUBLIC_OTHER = "□";   //其他

                        string PROTYPE_HEAT_URBAN = "□";    //城市供热工程
                        string PROTYPE_HEAT_INDU = "□";     //工业供热工程
                        string PROTYPE_HEAT_OTHER = "□";    //其他

                        string PROTYPE_ESTATE_RESI = "□";  //住宅项目
                        string PROTYPE_ESTATE_BUSI = "□";   //商业项目
                        string PROTYPE_ESTATE_COMP = "□";   //综合体项目
                        string PROTYPE_ESTATE_OTHER = "□";  //其他

                        ////////////////项目实施性质////////////////
                        string PRONATURE_NEW = "□";         //新建
                        string PRONATURE_EXP = "□";         //扩建
                        string PRONATURE_RENO = "□";       //改造
                        string PRONATURE_OTHER = "□";      //其他

                        ///////////////项目来源/////////////////
                        string PROSOURCE_INTE = "□";   //华西工程国际销售公司
                        string PROSOURCE_DOME = "□";    //华西工程国内销售公司
                        string PROSOURCE_PPP = "□";    //华西工程PPP销售公司
                        string PROSOURCE_DESI = "□";    //华西工程设计研究院
                        string PROSOURCE_OPER = "□";   //华西工程运维公司
                        string PROSOURCE_SERV = "□";     //华西工程项目服务部
                        string PROSOURCE_STOCK = "□";   //股份公司
                        string PROSOURCE_INDU = "□";    //华西工业
                        string PROSOURCE_OTHER = "□";   //其他

                        ///////////////合约类型///////////////////
                        string CONTRACTTYPE_E = "□";
                        string CONTRACTTYPE_P = "□";
                        string CONTRACTTYPE_C = "□";
                        string CONTRACTTYPE_PC = "□";
                        string CONTRACTTYPE_EP = "□";
                        string CONTRACTTYPE_EPC = "□";
                        string CONTRACTTYPE_EPCC = "□";
                        string CONTRACTTYPE_PPP = "□";
                        string CONTRACTTYPE_BOT = "□";
                        string CONTRACTTYPE_BOO = "□";
                        string CONTRACTTYPE_OTHER = "□";   //其他 


                        string PROTYPE_ENVI_SUPP = ""; //环保工程 补充说明
                        string PROTYPE_PUBLIC_SUPP = "";//市政工程 补充说明
                        string PROTYPE_HEAT_SUPP = "";//热力工程 补充说明
                        string PROTYPE_ESTATE_SUPP = "";//地产项目 补充说明
                        string PROTYPE_SUPP = "";//项目类型 补充说明

                        string PRONATURE_SUPP = "";//项目实施性质 补充说明
                        string PROSOURCE_SUPP = "";//项目来源 补充说明
                        string CONTRACTTYPE_SUPP = "";//合约类型 补充说明
                        #endregion

                        #region 勾选项目类型
                        if (projectType == "发电工程")
                        {
                            PROTYPE_POWER = "☑";

                            if (projectTypeII == "燃煤发电")
                                PROTYPE_POWER_COAL = "☑";
                            else if (projectTypeII == "生物质发电")
                                PROTYPE_POWER_BIOM = "☑";
                            else if (projectTypeII == "燃机发电")
                                PROTYPE_POWER_PLANT = "☑";
                            else if (projectTypeII == "生活垃圾发电")
                                PROTYPE_POWER_DOME = "☑";
                            else if (projectTypeII == "光热发电")
                                PROTYPE_POWER_OPTO = "☑";
                            else if (projectTypeII == "光伏发电")
                                PROTYPE_POWER_PV = "☑";
                            else if (projectTypeII == "风力发电")
                                PROTYPE_POWER_WIND = "☑";
                            else if (projectTypeII == "核电工程")
                                PROTYPE_POWER_NUCL = "☑";
                            else if (projectTypeII == "地热发电")
                                PROTYPE_POWER_GEOT = "☑";
                            else if (projectTypeII == "余热发电")
                                PROTYPE_POWER_COGE = "☑";
                            else if (projectTypeII == "油·气发电")
                                PROTYPE_POWER_OILGAS = "☑";
                            else if (projectTypeII == "燃气联合循环发电")
                                PROTYPE_POWER_GASC = "☑";
                        }
                        else if (projectType == "环保工程")
                        {
                            PROTYPE_ENVI = "☑";

                            if (projectTypeII == "烟气脱硫工程")
                                PROTYPE_ENVI_DESU = "☑";
                            else if (projectTypeII == "烟气脱硝工程")
                                PROTYPE_ENVI_DENI = "☑";
                            else if (projectTypeII == "烟气除尘工程")
                                PROTYPE_ENVI_DUST = "☑";
                            else if (projectTypeII == "厨余垃圾处理工程")
                                PROTYPE_ENVI_KITC = "☑";
                            else if (projectTypeII == "静脉产业园工程")
                                PROTYPE_ENVI_VENO = "☑";
                            else if (projectTypeII == "水处理工程")
                                PROTYPE_ENVI_WATER = "☑";
                            else if (projectTypeII == "污泥处理工程")
                                PROTYPE_ENVI_SLUD = "☑";
                            else if (projectTypeII == "农林废弃物综合利用")
                                PROTYPE_ENVI_FORE = "☑";
                            else if (projectTypeII == "其他")
                            {
                                PROTYPE_ENVI_OTHER = "☑";
                                PROTYPE_ENVI_SUPP = projectTypeSupp;
                            }

                        }
                        else if (projectType == "市政工程")
                        {
                            PROTYPE_PUBLIC = "☑";

                            if (projectTypeII == "市政设施工程")
                                PROTYPE_PUBLIC_MUNI = "☑";
                            else if (projectTypeII == "工业设施工程")
                                PROTYPE_PUBLIC_INDU = "☑";
                            else if (projectTypeII == "民用设施工程")
                                PROTYPE_PUBLIC_CIVIL = "☑";
                            else if (projectTypeII == "公共设施项目工程")
                                PROTYPE_PUBLIC_PUBL = "☑";
                            else if (projectTypeII == "其他")
                            {
                                PROTYPE_PUBLIC_OTHER = "☑";
                                PROTYPE_PUBLIC_SUPP = projectTypeSupp;
                            }
                        }
                        else if (projectType == "热力工程")
                        {
                            PROTYPE_HEAT = "☑";

                            if (projectTypeII == "城市供热工程")
                                PROTYPE_HEAT_URBAN = "☑";
                            else if (projectTypeII == "工业供热工程")
                                PROTYPE_HEAT_INDU = "☑";
                            else if (projectTypeII == "其他")
                            {
                                PROTYPE_HEAT_OTHER = "☑";
                                PROTYPE_HEAT_SUPP = projectTypeSupp;
                            }

                        }
                        else if (projectType == "地产项目")
                        {
                            PROTYPE_ESTATE = "☑";

                            if (projectTypeII == "住宅项目")
                                PROTYPE_ESTATE_RESI = "☑";
                            else if (projectTypeII == "商业项目")
                                PROTYPE_ESTATE_BUSI = "☑";
                            else if (projectTypeII == "综合体项目")
                                PROTYPE_ESTATE_COMP = "☑";
                            else if (projectTypeII == "其他")
                            {
                                PROTYPE_ESTATE_OTHER = "☑";
                                PROTYPE_ESTATE_SUPP = projectTypeSupp;
                            }

                        }
                        else if (projectType == "其他")
                        {
                            PROTYPE_OTHER = "☑";
                            PROTYPE_SUPP = projectTypeSupp;
                        }

                        #endregion


                        #region 勾选项目实施性质，项目来源和合约类型

                        //勾选项目实施性质
                        if (projectNature == "新建")
                        {
                            PRONATURE_NEW = "☑";
                        }
                        else if (projectNature == "扩建")
                        {
                            PRONATURE_EXP = "☑";
                        }
                        else if (projectNature == "改造")
                        {
                            PRONATURE_RENO = "☑";
                        }
                        else if (projectNature == "其他")
                        {
                            PRONATURE_OTHER = "☑";
                            PRONATURE_SUPP = projectNatureSupp;
                        }

                        //勾选项目来源
                        if (projectSource == "华西工程国际销售公司")
                        {
                            PROSOURCE_INTE = "☑";
                        }
                        else if (projectSource == "华西工程国内销售公司")
                        {
                            PROSOURCE_DOME = "☑";
                        }
                        else if (projectSource == "华西工程PPP销售公司")
                        {
                            PROSOURCE_PPP = "☑";
                        }
                        else if (projectSource == "华西工程设计研究院")
                        {
                            PROSOURCE_DESI = "☑";
                        }
                        else if (projectSource == "华西工程运维公司")
                        {
                            PROSOURCE_OPER = "☑";
                        }
                        else if (projectSource == "华西工程项目服务部")
                        {
                            PROSOURCE_SERV = "☑";
                        }
                        else if (projectSource == "股份公司")
                        {
                            PROSOURCE_STOCK = "☑";
                        }
                        else if (projectSource == "华西工业")
                        {
                            PROSOURCE_INDU = "☑";
                        }
                        else if (projectSource == "其他")
                        {
                            PROSOURCE_OTHER = "☑";
                            PROSOURCE_SUPP = projectSourceSupp;
                        }

                        //勾选合约类型
                        if (contractType == "E")
                        {
                            CONTRACTTYPE_E = "☑";
                        }
                        else if (contractType == "P")
                        {
                            CONTRACTTYPE_P = "☑";
                        }
                        else if (contractType == "C")
                        {
                            CONTRACTTYPE_C = "☑";
                        }
                        else if (contractType == "PC")
                        {
                            CONTRACTTYPE_PC = "☑";
                        }
                        else if (contractType == "EP")
                        {
                            CONTRACTTYPE_EP = "☑";
                        }
                        else if (contractType == "EPC")
                        {
                            CONTRACTTYPE_EPC = "☑";
                        }
                        else if (contractType == "EPCC")
                        {
                            CONTRACTTYPE_EPCC = "☑";
                        }
                        else if (contractType == "PPP")
                        {
                            CONTRACTTYPE_PPP = "☑";
                        }
                        else if (contractType == "BOT")
                        {
                            CONTRACTTYPE_BOT = "☑";
                        }
                        else if (contractType == "BOO")
                        {
                            CONTRACTTYPE_BOO = "☑";
                        }
                        else if (contractType == "其他")
                        {
                            CONTRACTTYPE_OTHER = "☑";
                            CONTRACTTYPE_SUPP = contractTypeSupp;
                        }

                        #endregion


                        #region 添加勾选项到哈希表
                        htUserKeyWord.Add("PROTYPE_POWER", PROTYPE_POWER);//发电工程
                        htUserKeyWord.Add("PROTYPE_ENVI", PROTYPE_ENVI);//环保工程
                        htUserKeyWord.Add("PROTYPE_PUBLIC", PROTYPE_PUBLIC);//环保工程
                        htUserKeyWord.Add("PROTYPE_HEAT", PROTYPE_HEAT);//热力工程
                        htUserKeyWord.Add("PROTYPE_ESTATE", PROTYPE_ESTATE);//地产项目
                        htUserKeyWord.Add("PROTYPE_OTHER", PROTYPE_OTHER);//其他

                        htUserKeyWord.Add("PROTYPE_POWER_COAL", PROTYPE_POWER_COAL);//燃煤发电
                        htUserKeyWord.Add("PROTYPE_POWER_BIOM", PROTYPE_POWER_BIOM);//生物质发电
                        htUserKeyWord.Add("PROTYPE_POWER_PLANT", PROTYPE_POWER_PLANT);//燃机发电
                        htUserKeyWord.Add("PROTYPE_POWER_DOME", PROTYPE_POWER_DOME);//生活垃圾发电
                        htUserKeyWord.Add("PROTYPE_POWER_OPTO", PROTYPE_POWER_OPTO);//光热发电
                        htUserKeyWord.Add("PROTYPE_POWER_PV", PROTYPE_POWER_PV);//光伏发电
                        htUserKeyWord.Add("PROTYPE_POWER_WIND", PROTYPE_POWER_WIND);//风力发电
                        htUserKeyWord.Add("PROTYPE_POWER_NUCL", PROTYPE_POWER_NUCL);//核电工程 
                        htUserKeyWord.Add("PROTYPE_POWER_GEOT", PROTYPE_POWER_GEOT);//地热发电
                        htUserKeyWord.Add("PROTYPE_POWER_COGE", PROTYPE_POWER_COGE);//余热发电
                        htUserKeyWord.Add("PROTYPE_POWER_OILGAS", PROTYPE_POWER_OILGAS);//油·气发电
                        htUserKeyWord.Add("PROTYPE_POWER_GASC", PROTYPE_POWER_GASC);//燃气联合循环发电

                        htUserKeyWord.Add("PROTYPE_ENVI_DESU", PROTYPE_ENVI_DESU);//烟气脱硫工程
                        htUserKeyWord.Add("PROTYPE_ENVI_DENI", PROTYPE_ENVI_DENI);//烟气脱硝工程
                        htUserKeyWord.Add("PROTYPE_ENVI_DUST", PROTYPE_ENVI_DUST);//烟气除尘工程
                        htUserKeyWord.Add("PROTYPE_ENVI_KITC", PROTYPE_ENVI_KITC);//厨余垃圾处理工程
                        htUserKeyWord.Add("PROTYPE_ENVI_VENO", PROTYPE_ENVI_VENO);//静脉产业园工程
                        htUserKeyWord.Add("PROTYPE_ENVI_WATER", PROTYPE_ENVI_WATER);//水处理工程
                        htUserKeyWord.Add("PROTYPE_ENVI_SLUD", PROTYPE_ENVI_SLUD);//污泥处理工程
                        htUserKeyWord.Add("PROTYPE_ENVI_FORE", PROTYPE_ENVI_FORE);//农林废弃物综合利用
                        htUserKeyWord.Add("PROTYPE_ENVI_OTHER", PROTYPE_ENVI_OTHER);//其他

                        htUserKeyWord.Add("PROTYPE_PUBLIC_MUNI", PROTYPE_PUBLIC_MUNI);//市政设施工程
                        htUserKeyWord.Add("PROTYPE_PUBLIC_INDU", PROTYPE_PUBLIC_INDU);//工业设施工程
                        htUserKeyWord.Add("PROTYPE_PUBLIC_CIVIL", PROTYPE_PUBLIC_CIVIL);//民用设施工程
                        htUserKeyWord.Add("PROTYPE_PUBLIC_PUBL", PROTYPE_PUBLIC_PUBL);//公共设施项目工程
                        htUserKeyWord.Add("PROTYPE_PUBLIC_OTHER", PROTYPE_PUBLIC_OTHER);//其他

                        htUserKeyWord.Add("PROTYPE_HEAT_URBAN", PROTYPE_HEAT_URBAN);//城市供热工程
                        htUserKeyWord.Add("PROTYPE_HEAT_INDU", PROTYPE_HEAT_INDU);//工业供热工程
                        htUserKeyWord.Add("PROTYPE_HEAT_OTHER", PROTYPE_HEAT_OTHER);//其他

                        htUserKeyWord.Add("PROTYPE_ESTATE_RESI", PROTYPE_ESTATE_RESI);//住宅项目
                        htUserKeyWord.Add("PROTYPE_ESTATE_BUSI", PROTYPE_ESTATE_BUSI);//商业项目
                        htUserKeyWord.Add("PROTYPE_ESTATE_COMP", PROTYPE_ESTATE_COMP);//综合体项目
                        htUserKeyWord.Add("PROTYPE_ESTATE_OTHER", PROTYPE_ESTATE_OTHER);//其他

                        ////////////////项目实施性质////////////////
                        htUserKeyWord.Add("PRONATURE_NEW", PRONATURE_NEW);//新建
                        htUserKeyWord.Add("PRONATURE_EXP", PRONATURE_EXP);//扩建
                        htUserKeyWord.Add("PRONATURE_RENO", PRONATURE_RENO);//改造
                        htUserKeyWord.Add("PRONATURE_OTHER", PRONATURE_OTHER);//其他

                        ///////////////项目来源/////////////////
                        htUserKeyWord.Add("PROSOURCE_INTE", PROSOURCE_INTE);//华西工程国际销售公司
                        htUserKeyWord.Add("PROSOURCE_DOME", PROSOURCE_DOME);//华西工程国内销售公司
                        htUserKeyWord.Add("PROSOURCE_PPP", PROSOURCE_PPP);//华西工程PPP销售公司
                        htUserKeyWord.Add("PROSOURCE_DESI", PROSOURCE_DESI);//华西工程设计研究院
                        htUserKeyWord.Add("PROSOURCE_OPER", PROSOURCE_OPER);//华西工程运维公司
                        htUserKeyWord.Add("PROSOURCE_SERV", PROSOURCE_SERV);//华西工程项目服务部
                        htUserKeyWord.Add("PROSOURCE_STOCK", PROSOURCE_STOCK);//股份公司
                        htUserKeyWord.Add("PROSOURCE_INDU", PROSOURCE_INDU);//华西工业
                        htUserKeyWord.Add("PROSOURCE_OTHER", PROSOURCE_OTHER);//其他


                        ///////////////合约类型///////////////////
                        htUserKeyWord.Add("CONTRACTTYPE_E", CONTRACTTYPE_E);//E
                        htUserKeyWord.Add("CONTRACTTYPE_P", CONTRACTTYPE_P);//P
                        htUserKeyWord.Add("CONTRACTTYPE_C", CONTRACTTYPE_C);//C
                        htUserKeyWord.Add("CONTRACTTYPE_PC", CONTRACTTYPE_PC);//PC
                        htUserKeyWord.Add("CONTRACTTYPE_EP", CONTRACTTYPE_EP);//EP
                        htUserKeyWord.Add("CONTRACTTYPE_EPC", CONTRACTTYPE_EPC);//EPC
                        htUserKeyWord.Add("CONTRACTTYPE_EPCC", CONTRACTTYPE_EPCC);//EPCC
                        htUserKeyWord.Add("CONTRACTTYPE_PPP", CONTRACTTYPE_PPP);//PPP
                        htUserKeyWord.Add("CONTRACTTYPE_BOT", CONTRACTTYPE_BOT);//BOT
                        htUserKeyWord.Add("CONTRACTTYPE_BOO", CONTRACTTYPE_BOO);//BOO
                        htUserKeyWord.Add("CONTRACTTYPE_OTHER", CONTRACTTYPE_OTHER);//其他 


                        htUserKeyWord.Add("PROTYPE_ENVI_SUPP", PROTYPE_ENVI_SUPP);//环保工程 补充说明 
                        htUserKeyWord.Add("PROTYPE_PUBLIC_SUPP", PROTYPE_PUBLIC_SUPP);//市政工程 补充说明
                        htUserKeyWord.Add("PROTYPE_HEAT_SUPP", PROTYPE_HEAT_SUPP);//热力工程 补充说明
                        htUserKeyWord.Add("PROTYPE_ESTATE_SUPP", PROTYPE_ESTATE_SUPP);//地产项目 补充说明
                        htUserKeyWord.Add("PROTYPE_SUPP", PROTYPE_SUPP);//项目类型 补充说明

                        htUserKeyWord.Add("PRONATURE_SUPP", PRONATURE_SUPP);//项目实施性质 补充说明
                        htUserKeyWord.Add("PROSOURCE_SUPP", PROSOURCE_SUPP);//项目来源 补充说明
                        htUserKeyWord.Add("CONTRACTTYPE_SUPP", CONTRACTTYPE_SUPP);//合约类型 补充说明

                        #endregion
                        #endregion



                        htUserKeyWord.Add("PROCODE", projectCode);//项目代码
                        htUserKeyWord.Add("PRONAME", projectDescCN);//项目名称（中文）
                        htUserKeyWord.Add("PROENGLISH", projectDescEN);//项目名称（英文）
                        htUserKeyWord.Add("ABBREVIATION", projectShortDescCN);//项目简称（中文）
                        htUserKeyWord.Add("ABBREVIATIONENG", projectShortDescEN);//项目简称（英文）
                        htUserKeyWord.Add("COMMUNICATIONCODE", communicationCode);//通信代码
                        htUserKeyWord.Add("OWNER", unintName);//业主
                        htUserKeyWord.Add("PROADDRESS", projectAddr);//项目实施地址
                        htUserKeyWord.Add("ESTIMATEMONEY", projectAmount); //预估金额
                        htUserKeyWord.Add("PLANTIME", projectDuration);//计划工期
                        htUserKeyWord.Add("PLANSTARTDATE", planStartDate);//计划开始时间
                        htUserKeyWord.Add("PLANENDDATE", planEndDate);//计划结束时间
                        htUserKeyWord.Add("REMARKS", projectOverview);//项目概况
                        htUserKeyWord.Add("EMAIL", email);
                        //htUserKeyWord.Add("FORM", strFormQuantity);


                        //this.Cursor = Cursors.WaitCursor;
                        string workingPath = m_Project.dBSource.LoginUser.WorkingPath;
                        //AttrData ad = m_Project.GetAttrDataByKeyWord("ISSAVE");
                        //if (ad != null)
                        //{
                        //    ad.SetCodeDesc("");
                        //}
                        //m_Project.AttrDataList.SaveData();


                        try
                        {
                            //上传下载文档
                            string exchangfilename = "项目立项单";

                            //获取网站路径
                            string sPath = System.Web.HttpContext.Current.Server.MapPath("/ISO/HXEPC/");

                            //获取模板文件路径
                            string modelFileName = sPath + exchangfilename + ".doc";

                            //获取即将生成的联系单文件路径
                            string locFileName = item.FullPathFile;

                            //FTPFactory factory = m_Project.Storage.FTP ?? new FTPFactory(m_Project.Storage);
                            //string locFileName = m_Project.dBSource.LoginUser.WorkingPath + item.Code + ".doc";
                            //factory.download(@"\ISO\" + exchangfilename + ".doc", locFileName, false);
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
                                    office.WriteDataToDocument(item, locFileName, htUserKeyWord, htUserKeyWord);
                                }
                                catch { }
                                finally
                                {

                                    //解锁
                                    muxConsole.ReleaseMutex();
                                }
                            }


                            int length = (int)info.Length;
                            item.O_size = new int?(length);
                            item.Modify();
                            //base.DialogResult = DialogResult.OK;
                            //this.Cursor = Cursors.Default;
                            //base.Close();
                            //CommonFunction.InsertDocListAndOpenDoc(this.m_DocList, item);

                            if (string.IsNullOrEmpty(strDocList))
                            {
                                strDocList = item.KeyWord;
                            }
                            else
                            {
                                strDocList = item.KeyWord + "," + strDocList;
                            }

                            ////增加领导组
                            //AVEVA.CDMS.Server.Group gp = dbsource.GetGroupByName("AdminGroup");
                            //if (gp != null)
                            //{
                            //    m_NewProject.groupList[0].AddGroup(gp);
                            //    m_NewProject.groupList[0].Modify();
                            //}

                            //这里刷新数据源，否则创建流程的时候获取不了专业字符串
                            DBSourceController.RefreshDBSource(sid);

                            reJo.success = true;
                            reJo.data = new JArray(new JObject(new JProperty("ProjectKeyword", item.Project.KeyWord),
                                new JProperty("DocKeyword", item.KeyWord), new JProperty("DocList", strDocList)));
                            return reJo;
                        }
                        catch { }
                        #endregion


                        reJo.data = new JArray(new JObject(new JProperty("projectKeyword", m_NewProject.KeyWord)));
                        reJo.success = true;
                        return reJo;

                        //AVEVA.CDMS.WebApi.DBSourceController.RefreshDBSource(sid);

                    }
                }
            }
            catch (Exception e)
            {
                reJo.msg = e.Message;
                CommonController.WebWriteLog(reJo.msg);
            }

            return reJo;
        }
        internal static Project CreateChildProject(Project parentProject,string code,string desc,string tempDefn) {
            AVEVA.CDMS.Server.Project project = null;
            
            List<TempDefn> tempDefnByKeyWord = parentProject.dBSource.GetTempDefnByKeyWord(tempDefn);
            TempDefn defn2 = ((tempDefnByKeyWord != null) && (tempDefnByKeyWord.Count > 0)) ? tempDefnByKeyWord[0] : null;
            string str = code;
            project = parentProject.GetProjectByName(str);
            if (project == null)
            {
                //project = m_NewProject.NewProject(str.Substring(0, str.IndexOf("__")), str.Substring(str.IndexOf("__") + 2), m_NewProject.Storage, defn2);
                project = parentProject.NewProject(str, desc, parentProject.Storage, defn2);

            }
            if (project == null)
            {
                
                return null;
            }
            return project;
        }
        //项目立项备份
        public static ExReJObject CreateRootProjectX_bak(DBSource dbsource, JArray jaAttr)
        {
            ExReJObject reJo = new ExReJObject();

            try
            {

                #region 获取项目参数项目

                //JArray jaAttr = (JArray)JsonConvert.DeserializeObject(projectAttrJson);

                //获取项目参数项目
                string projectCode = "", projectDesc = "", unintName = "", projectNo = "",
                    unitMan = "", projectAmount = "", unitManPhone = "", buildArea = "",
                    writeMan = "", contractTerms = "", designPhase = "", projectType = "",
                    quality = "", releaseDate = "", planStartDate = "", planEndDate = "",
                    projectTypeR = "", projectSize = "";

                //合同号
                string projectNomber = "";
                //备注
                string Remarks = "";

                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    //获取项目代码
                    if (strName == "projectCode") projectCode = strValue;

                    //获取项目名称
                    else if (strName == "projectDesc") projectDesc = strValue;

                    //获取建设单位
                    else if (strName == "unintName") unintName = strValue;

                    //获取合同号
                    else if (strName == "projectNo") projectNo = strValue;

                    //获取甲方联系人
                    else if (strName == "unitMan") unitMan = strValue;

                    //获取合同额
                    else if (strName == "projectAmount") projectAmount = strValue;

                    //获取甲方联系人电话
                    else if (strName == "unitManPhone") unitManPhone = strValue;

                    //获取建筑面积
                    else if (strName == "buildArea") buildArea = strValue;

                    //获取立项人
                    else if (strName == "writeMan") writeMan = strValue;

                    //获取合同主要条款
                    else if (strName == "contractTerms") contractTerms = strValue;

                    //获取设计阶段
                    else if (strName == "designPhase") designPhase = strValue;

                    //获取项目类型
                    else if (strName == "projectType") projectType = strValue;

                    //获取质量目标
                    else if (strName == "quality") quality = strValue;

                    //获取任务下达时间
                    else if (strName == "releaseDate") releaseDate = strValue;

                    //获取计划开始时间
                    else if (strName == "planStartDate") planStartDate = strValue;

                    //获取计划结束时间
                    else if (strName == "planEndDate") planEndDate = strValue;

                    //获取项目类型
                    else if (strName == "projectTypeR") projectTypeR = strValue;

                    //获取项目大小
                    else if (strName == "projectSize") projectSize = strValue;

                }

                if (string.IsNullOrEmpty(projectCode))
                {
                    reJo.msg = "请填写项目编号！";
                    return reJo;
                }
                else if (string.IsNullOrEmpty(projectDesc))
                {
                    reJo.msg = "请填写项目名称！";
                    return reJo;
                }
                else if (string.IsNullOrEmpty(unintName))
                {
                    reJo.msg = "请填写建设单位！";
                    return reJo;
                }
                else if (string.IsNullOrEmpty(designPhase))
                {
                    reJo.msg = "请选择设计阶段！";
                    return reJo;
                }

                #endregion


                //  根据名称查找项目模板(根目录)对象
                List<TempDefn> tempDefnByCode = dbsource.GetTempDefnByCode("GEDI_INTERFACEPROJECT");
                TempDefn mTempDefn = (tempDefnByCode != null) ? tempDefnByCode[0] : null;
                if (mTempDefn == null)
                {
                    reJo.msg = "欲创建的项目关联的模板不存在！不能完成创建";
                    return reJo;
                }
                else
                {

                    //获取DBSource的虚拟Local目录
                    Project m_Project = dbsource.NewProject(enProjectType.Local);

                    //创建项目
                    Project m_NewProject = m_Project.NewProject(projectCode, projectDesc, null, mTempDefn);


                    if (m_NewProject == null)
                    {
                        reJo.msg = "新建项目失败";
                        return reJo;

                    }
                    else
                    {

                        //foreach (KeyValuePair<string, string> kvp in context)
                        //{
                        //    SetAttrData(pro, kvp.Key.ToUpper(), kvp.Value, "");

                        //}
                        //pro.GetAttrDataByKeyWord("unintName").SetCodeDesc(unintName);

                        foreach (JObject joAttr in jaAttr)
                        {
                            string strName = joAttr["name"].ToString();
                            string strValue = joAttr["value"].ToString();

                            AttrData data;
                            if ((data = m_NewProject.GetAttrDataByKeyWord(strName)) != null)
                            {
                                data.SetCodeDesc(strValue);
                            }

                        }

                        ////存进数据库
                        m_NewProject.AttrDataList.SaveData();

                        AVEVA.CDMS.Server.Project project = null;

                        List<TempDefn> tempDefnByKeyWord = m_NewProject.dBSource.GetTempDefnByKeyWord("GEDIHD_DESIGNPHASE");
                        TempDefn defn2 = ((tempDefnByKeyWord != null) && (tempDefnByKeyWord.Count > 0)) ? tempDefnByKeyWord[0] : null;
                        string str = designPhase;
                        project = m_NewProject.GetProjectByName(str.Substring(0, str.IndexOf("__")));
                        if (project == null)
                        {
                            project = m_NewProject.NewProject(str.Substring(0, str.IndexOf("__")), str.Substring(str.IndexOf("__") + 2), m_NewProject.Storage, defn2);
                        }
                        if (project == null)
                        {

                            reJo.msg = "创建/获取设计阶段失败，请联系管理员！";
                            return reJo;
                        }
                        else
                        {

                            //DBSourceController.RefreshDBSource(sid);
                            reJo.data = new JArray(new JObject(new JProperty("projectKeyword", project.KeyWord)));
                            reJo.success = true;
                            return reJo;

                        }
                        //AVEVA.CDMS.WebApi.DBSourceController.RefreshDBSource(sid);
                    }

                }
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

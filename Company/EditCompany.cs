using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using AVEVA.CDMS.Server;
using AVEVA.CDMS.Common;
using AVEVA.CDMS.WebApi;

namespace AVEVA.CDMS.HXEPC_Plugins
{
    public class EditCompanyInfo
    {
        /// <summary>
        /// 新建厂家资料目录时，获取默认值
        /// </summary>
        /// <param name="sid"></param>
        /// <returns></returns>
        public static JObject GetEditCompanyDefault(string sid, string ProjectKeyword)
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

                Project m_prj = dbsource.GetProjectByKeyWord(ProjectKeyword);
                if (m_prj == null)
                {
                    reJo.msg = "参数错误，目录不存在！";
                    return reJo.Value;
                }

                Project prjProject = CommonFunction.getParentProjectByTempDefn(m_prj, "HXNY_DOCUMENTSYSTEM");

                if (prjProject == null)
                {
                    reJo.msg = "获取项目目录失败！";
                    return reJo.Value;
                }

                //获取项目代码
                string strProjCode = prjProject.Code;//.GetAttrDataByKeyWord("COMPANY").ToString;
                string strProjDesc = prjProject.Description;

                JArray jaCompany = new JArray();
                JObject joCompany = new JObject();

                List<DictData> dictDataList = dbsource.GetDictDataList("Unit");
                //[o_Code]:公司编码,[o_Desc]：公司描述,[o_sValue1]：项目代码

                foreach (DictData data6 in dictDataList)
                {
                    if (!string.IsNullOrEmpty(data6.O_sValue1) && data6.O_sValue1 == strProjCode)
                    {
                        joCompany = new JObject(
                            new JProperty("companyId", data6.O_ID.ToString()),
                            new JProperty("companyCode", data6.O_Code),
                            new JProperty("companyDesc", data6.O_Desc),
                            new JProperty("secretarilman", data6.O_sValue3)
                            );
                        jaCompany.Add(joCompany);
                    }
                }


                reJo.data = new JArray(
                    new JObject(new JProperty("projectCode", strProjCode),
                    new JProperty("projectDesc", strProjDesc),
                    new JProperty("CompanyList", jaCompany)));
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
        /// 新建厂家资料目录
        /// </summary>
        public static JObject CreateCompany(string sid, string ProjectKeyword, string projectAttrJson)
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

                Project m_prj = dbsource.GetProjectByKeyWord(ProjectKeyword);
                if (m_prj == null)
                {
                    reJo.msg = "参数错误，目录不存在！";
                    return reJo.Value;
                }


                #region 获取传递过来的属性参数
                //获取传递过来的属性参数
                JArray jaAttr = (JArray)JsonConvert.DeserializeObject(projectAttrJson);

                string strCompanyCode = "", strCompanyDesc = "",
                    strSecretarilman = "", strCompanyChinese = "",
                    strAddress = "", strProvince = "",
                    strPostCode = "", strEMail = "",
                    strReceiver = "", strFaxNo = "", strPhone = "";

                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    switch (strName)
                    {
                        case "companyCode":
                            strCompanyCode = strValue;
                            break;
                        case "companyDesc":
                            strCompanyDesc = strValue;
                            break;
                        case "secretarilman":
                            strSecretarilman = strValue;
                            break;
                    }
                }

                if (string.IsNullOrEmpty(strCompanyCode))
                {
                    reJo.msg = "请输入项目编号！";
                    return reJo.Value;
                }
                if (string.IsNullOrEmpty(strCompanyDesc))
                {
                    reJo.msg = "请输入项目名称！";
                    return reJo.Value;
                }

                #endregion

                Project prjProject = CommonFunction.getParentProjectByTempDefn(m_prj, "HXNY_DOCUMENTSYSTEM");

                if (prjProject == null)
                {
                    reJo.msg = "获取项目目录失败！";
                    return reJo.Value;
                }

                //获取项目代码
                string strProjCode = prjProject.Code;//.GetAttrDataByKeyWord("COMPANY").ToString;

                JArray jaData = new JArray();
                JObject joCompany = new JObject();

                List<DictData> dictDataList = dbsource.GetDictDataList("Unit");
                //[o_Code]:公司编码,[o_Desc]：公司描述,[o_sValue1]：项目代码

                foreach (DictData data6 in dictDataList)
                {
                    if (!string.IsNullOrEmpty(data6.O_sValue1) && data6.O_sValue1 == strProjCode && data6.O_Code == strCompanyCode)
                    {
                        reJo.msg = "已经存在相同的参建单位，请返回重试！";
                        return reJo.Value;
                    }
                }

                //自动设置文控，如果没有与单位代码一样的用户，就自动添加用户
                SetUnitSecUser(sid, dbsource, prjProject, strCompanyCode, strCompanyDesc, ref strSecretarilman);

                //#region 自动设置文控，如果没有与单位代码一样的用户，就自动添加用户

                //if (string.IsNullOrEmpty(strSecretarilman))
                //{
                //    User secUser = dbsource.GetUserByCode(strCompanyCode);
                //    if (secUser == null)
                //    {
                //        ////UserController. CreateUser(sid, strCompanyCode, strCompanyDesc + "文控", "", "0",
                //        ////            "0", "", strCompanyCode, strCompanyCode);
                //        ////secUser = dbsource.GetUserByCode(strCompanyCode);
                //        secUser = dbsource.NewUser(
                //                            enUserFlage.OnLine,
                //                            enUserType.Default,
                //                                "",
                //                               strCompanyCode,
                //                                strCompanyDesc + "文控",
                //                                strCompanyCode,
                //                                "",
                //                                null
                //                                );

                //        if (secUser != null)
                //        {

                //            User m_user = secUser;
                //            m_user.O_suser1 = m_user.dBSource.GUID;

                //            m_user.Modify();

                //            // 强制刷新共享数据源
                //            //
                //            //DBSourceController.RefreshShareDBManager();
                //            DBSourceController.RefreshDBSource(sid);
                //            strSecretarilman = secUser.ToString;
                //        }
                //    }
                //    else
                //    {
                //        strSecretarilman = secUser.ToString;
                //    }
                //}
                //#endregion

                //dbsource.NewDictData
                #region 添加到数据字典
                //添加到数据字典

                //DictData dictdata = new DictData();
                //dictdata.StatusNew = true;
                //dictdata.O_skey = "Unit";
                //dictdata.O_datatype = (int)enDictDataType.TableHead;
                //dictdata.O_Code = strCompanyCode;
                //dictdata.O_Desc = strCompanyDesc;
                //dictdata.O_sValue1 = strProjCode;

                ////设置属性的值
                //SetDictDataPropertyValue(dictdata, 0);

                //dictdata.Modify();
                //strProjCode = "GEDI";

                string format = "insert CDMS_DictData (" +
                    "o_parentno,o_datatype,o_ikey,o_skey,o_Code,o_Desc,o_sValue1,o_sValue2,o_sValue3,o_sValue4,o_sValue5,o_iValue1 ,o_iValue2)" +
                    " values ({0},{1},{2},'{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}',{11},{12}" + ")";
                //0,2,0,'Unit','"+strCompanyCode+"','"+strCompanyDesc+"','"+strProjCode+ "','','','','',0,0
                format = string.Format(format, new object[] {
                    0,2,0,"Unit",strCompanyCode,strCompanyDesc,strProjCode,"",strSecretarilman,"","",0,0
                     });
                dbsource.DBExecuteSQL(format);

                DBSourceController.refreshDBSource(sid);

                //自动创建通信类下的单位目录
                CreateUnitProject(dbsource, prjProject, strCompanyCode, strCompanyDesc, strSecretarilman);

                ////通过以上的dictdata新建一个DictData
                //DictData newDictData = CreateNewDictData(dbsource, dictdata, enDictDataType.TableHead);

                //if (!newDictData.Write())
                //{
                //    //MessageBox.Show("创建失败!", "提示", MessageBoxButtons.OK);
                //    reJo.msg = "创建失败!";
                //    return reJo.Value;
                //} 
                #endregion

                ////获取所有厂家信息
                //TempDefn mTempDefn = GetTempDefn(dbsource, "COM_UNIT");
                //if (mTempDefn == null)
                //{
                //    reJo.msg = "获取参建单位模板失败，请联系管理员！";
                //    return reJo.Value;
                //}

                //Project project = m_prj.NewProject(strCompanyCode, strCompanyDesc, m_prj.Storage, mTempDefn);
                //if (project == null)
                //{
                //    reJo.msg = "新建版本目录失败，请联系管理员！";
                //    return reJo.Value;
                //}

                ////增加附加属性
                //try
                //{
                //    //project.GetAttrDataByKeyWord("FC_COMPANYCODE").SetCodeDesc(strCompanyCode);       //厂家编码
                //    //project.GetAttrDataByKeyWord("FC_COMPANYCHINESE").SetCodeDesc(strCompanyChinese);    //厂家名称
                //    //project.GetAttrDataByKeyWord("FC_ADDRESS").SetCodeDesc(strAddress);           //厂家地址
                //    //project.GetAttrDataByKeyWord("FC_PROVINCE").SetCodeDesc(strProvince);          //厂家省份
                //    //project.GetAttrDataByKeyWord("FC_POSTCODE").SetCodeDesc(strPostCode);          //厂家邮政
                //    //project.GetAttrDataByKeyWord("FC_EMAIL").SetCodeDesc(strEMail);             //厂家邮箱
                //    //project.GetAttrDataByKeyWord("FC_RECEIVER").SetCodeDesc(strReceiver);          //厂家收件人
                //    //project.GetAttrDataByKeyWord("FC_FAXNO").SetCodeDesc(strFaxNo);             //厂家传真号
                //    //project.GetAttrDataByKeyWord("FC_PHONE").SetCodeDesc(strPhone);             //收件人电话
                //    //project.AttrDataList.SaveData();
                //}
                //catch (Exception ex)
                //{
                //    //MessageBox.Show("获取厂家模板失败，请联系管理员！");
                //    reJo.msg = "获取厂家模板失败，请联系管理员！";
                //    return reJo.Value;
                //}

                //reJo.data = new JArray(new JObject(new JProperty("ProjectKeyword", project.KeyWord)));

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
        /// 新建厂家资料目录
        /// </summary>
        public static JObject EditCompany(string sid, string ProjectKeyword, string projectAttrJson)
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

                Project m_prj = dbsource.GetProjectByKeyWord(ProjectKeyword);
                if (m_prj == null)
                {
                    reJo.msg = "参数错误，目录不存在！";
                    return reJo.Value;
                }


                #region 获取传递过来的属性参数
                //获取传递过来的属性参数
                JArray jaAttr = (JArray)JsonConvert.DeserializeObject(projectAttrJson);

                string strCompanyId = "", strCompanyCode = "", strCompanyDesc = "",
                    strSecretarilman = "", strCompanyChinese = "",
                    strAddress = "", strProvince = "",
                    strPostCode = "", strEMail = "",
                    strReceiver = "", strFaxNo = "", strPhone = "";

                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    switch (strName)
                    {
                        case "companyId":
                            strCompanyId = strValue;
                            break;
                        case "companyCode":
                            strCompanyCode = strValue;
                            break;
                        case "companyDesc":
                            strCompanyDesc = strValue;
                            break;
                        case "secretarilman":
                            strSecretarilman = strValue;
                            break;
                    }
                }

                if (string.IsNullOrEmpty(strCompanyCode))
                {
                    reJo.msg = "请输入项目编号！";
                    return reJo.Value;
                }
                if (string.IsNullOrEmpty(strCompanyDesc))
                {
                    reJo.msg = "请输入项目名称！";
                    return reJo.Value;
                }
                #endregion

                Project prjProject = CommonFunction.getParentProjectByTempDefn(m_prj, "HXNY_DOCUMENTSYSTEM");

                if (prjProject == null)
                {
                    reJo.msg = "获取项目目录失败！";
                    return reJo.Value;
                }

                //User secretarilman = dbsource.GetUserByKeyWord(strSecretarilman);
                //if (secretarilman == null)
                //{
                //    reJo.msg = "参数错误！文控角色所选择的用户不存在！";
                //    return reJo.Value;
                //}

                int companyId = Convert.ToInt32(strCompanyId);
                //获取项目代码
                string strProjCode = prjProject.Code;//.GetAttrDataByKeyWord("COMPANY").ToString;


                JArray jaData = new JArray();
                JObject joCompany = new JObject();

                List<DictData> dictDataList = dbsource.GetDictDataList("Unit");
                //[o_Code]:公司编码,[o_Desc]：公司描述,[o_sValue1]：项目代码

                foreach (DictData data6 in dictDataList)
                {
                    if (!string.IsNullOrEmpty(data6.O_sValue1) && data6.O_sValue1 == strProjCode
                        && data6.O_Code == strCompanyCode && data6.O_ID != companyId)
                    {
                        reJo.msg = "已经存在相同的参建单位，请返回重试！";
                        return reJo.Value;
                    }
                }
                //dbsource.NewDictData


                #region 添加到数据字典
                //添加到数据字典

                //DictData dictdata = new DictData();
                //dictdata.StatusNew = true;
                //dictdata.O_skey = "Unit";
                //dictdata.O_datatype = (int)enDictDataType.TableHead;
                //dictdata.O_Code = strCompanyCode;
                //dictdata.O_Desc = strCompanyDesc;
                //dictdata.O_sValue1 = strProjCode;

                ////设置属性的值
                //SetDictDataPropertyValue(dictdata, 0);

                //dictdata.Modify();
                DictData dictData = null;

                foreach (DictData data6 in dictDataList)
                {
                    if (data6.O_ID == companyId)
                    {
                        dictData = data6;

                    }
                }

                if (dictData == null)
                {
                    reJo.msg = "参建单位ID不存在，请返回重试！";
                    return reJo.Value;

                }

                //自动设置文控，如果没有与单位代码一样的用户，就自动添加用户
                SetUnitSecUser(sid, dbsource, prjProject, strCompanyCode, strCompanyDesc,ref strSecretarilman);
                //#region 自动设置文控，如果没有与单位代码一样的用户，就自动添加用户
                //User secUser = null;
                //if (string.IsNullOrEmpty(strSecretarilman))
                //{
                //    secUser = dbsource.GetUserByCode(strCompanyCode);
                //    if (secUser == null)
                //    {
                //        ////UserController. CreateUser(sid, strCompanyCode, strCompanyDesc + "文控", "", "0",
                //        ////            "0", "", strCompanyCode, strCompanyCode);
                //        ////secUser = dbsource.GetUserByCode(strCompanyCode);
                //        secUser = dbsource.NewUser(
                //                            enUserFlage.OnLine,
                //                            enUserType.Default,
                //                                "",
                //                               strCompanyCode,
                //                                strCompanyDesc + "文控",
                //                                strCompanyCode,
                //                                "",
                //                                null
                //                                );

                //        if (secUser != null)
                //        {

                //            User m_user = secUser;
                //            m_user.O_suser1 = m_user.dBSource.GUID;

                //            m_user.Modify();

  

                //            // 强制刷新共享数据源
                //            //
                //            //DBSourceController.RefreshShareDBManager();
                //            DBSourceController.RefreshDBSource(sid);
                //            strSecretarilman = secUser.ToString;
                //        }
                //    }
                //    else
                //    {
                //        strSecretarilman = secUser.ToString;
                //    }

                   
                //}

                //if (secUser != null)
                //{
                //    //把用户添加到项目管理类里面的项目单位用户组
                //    Group group = dbsource.GetGroupByName(prjProject.Code + "_ALLUnit");


                //    if (group != null)
                //    {
                //        group.AddUser(secUser);
                //        group.Modify();
                //    }
                //}
                //    #endregion

                    dictData.O_Code = strCompanyCode;
                dictData.O_Desc = strCompanyDesc;
                dictData.O_sValue1 = strProjCode;
                dictData.O_sValue3 = strSecretarilman;// secretarilman.ToString;//
                dictData.Modify();
               
                DBSourceController.refreshDBSource(sid);

                ////通过以上的dictdata新建一个DictData
                //DictData newDictData = CreateNewDictData(dbsource, dictdata, enDictDataType.TableHead);

                //if (!newDictData.Write())
                //{
                //    //MessageBox.Show("创建失败!", "提示", MessageBoxButtons.OK);
                //    reJo.msg = "创建失败!";
                //    return reJo.Value;
                //} 
                #endregion

                //自动创建通信类下的单位目录
                CreateUnitProject(dbsource,prjProject,strCompanyCode,strCompanyDesc,strSecretarilman);

                ////自动创建通信类下的单位目录
                //#region 自动创建通信类下的单位目录
                //try
                //{
                //    TempDefn mTempDefn = Company.GetTempDefn(dbsource, "COM_UNIT");
                //    if (mTempDefn != null)
                //    {
                //        Project cdProject = CommonFunction.GetProjectByDesc(prjProject, "存档管理");
                //        if (cdProject != null)
                //        {
                //            Project txProject = CommonFunction.GetProjectByDesc(cdProject, "通信类");
                //            if (txProject != null)
                //            {
                //                Project project = txProject.NewProject(strCompanyCode, strCompanyDesc, m_prj.Storage, mTempDefn);

                //                if (project != null)
                //                {
                //                    //增加附加属性
                //                    try
                //                    {
                //                        project.GetAttrDataByKeyWord("UN_SECRETAARECTOR").SetCodeDesc(strSecretarilman);             //文控
                //                        project.AttrDataList.SaveData();
                //                    }
                //                    catch (Exception ex)
                //                    {
                //                        //MessageBox.Show("获取厂家模板失败，请联系管理员！");
                //                        reJo.msg = "获取厂家模板失败，请联系管理员！";
                //                        return reJo.Value;
                //                    }

                //                    TempDefn sfwTempDefn = Company.GetTempDefn(dbsource, "STO_SUBDOCUMENT");
                //                    if (sfwTempDefn != null)
                //                    {
                //                        Project swProject = project.NewProject("收文", "", m_prj.Storage, sfwTempDefn);
                //                        Project fwProject = project.NewProject("发文", "", m_prj.Storage, sfwTempDefn);

                //                        TempDefn typeTempDefn = Company.GetTempDefn(dbsource, "STO_COMTYPE");
                //                        if (typeTempDefn != null)
                //                        {
                //                            if (swProject != null)
                //                            {
                //                                swProject.NewProject("红头文", "", m_prj.Storage, typeTempDefn);
                //                                swProject.NewProject("会议纪要", "", m_prj.Storage, typeTempDefn);
                //                                swProject.NewProject("文件传递单", "", m_prj.Storage, typeTempDefn);
                //                                swProject.NewProject("信函", "", m_prj.Storage, typeTempDefn);
                //                            }

                //                            if (fwProject != null)
                //                            {
                //                                fwProject.NewProject("红头文", "", m_prj.Storage, typeTempDefn);
                //                                fwProject.NewProject("会议纪要", "", m_prj.Storage, typeTempDefn);
                //                                fwProject.NewProject("文件传递单", "", m_prj.Storage, typeTempDefn);
                //                                fwProject.NewProject("信函", "", m_prj.Storage, typeTempDefn);
                //                            }
                //                        }
                //                    }
                //                }
                //            }
                //        }
                //    }
                //}
                //catch
                //{

                //} 
                //#endregion

                //prjProject.GetProjectByName();


                //    reJo.msg = "获取参建单位模板失败，请联系管理员！";
                //    return reJo.Value;
                //}

                //Project project = m_prj.NewProject(strCompanyCode, strCompanyDesc, m_prj.Storage, mTempDefn);
                //if (project == null)
                //{
                //    reJo.msg = "新建版本目录失败，请联系管理员！";
                //    return reJo.Value;
                //}

                //增加附加属性
                try
                {
                    //project.GetAttrDataByKeyWord("FC_COMPANYCODE").SetCodeDesc(strCompanyCode);       //厂家编码
                    //project.GetAttrDataByKeyWord("FC_COMPANYCHINESE").SetCodeDesc(strCompanyChinese);    //厂家名称
                    //project.GetAttrDataByKeyWord("FC_ADDRESS").SetCodeDesc(strAddress);           //厂家地址
                    //project.GetAttrDataByKeyWord("FC_PROVINCE").SetCodeDesc(strProvince);          //厂家省份
                    //project.GetAttrDataByKeyWord("FC_POSTCODE").SetCodeDesc(strPostCode);          //厂家邮政
                    //project.GetAttrDataByKeyWord("FC_EMAIL").SetCodeDesc(strEMail);             //厂家邮箱
                    //project.GetAttrDataByKeyWord("FC_RECEIVER").SetCodeDesc(strReceiver);          //厂家收件人
                    //project.GetAttrDataByKeyWord("FC_FAXNO").SetCodeDesc(strFaxNo);             //厂家传真号
                    //project.GetAttrDataByKeyWord("FC_PHONE").SetCodeDesc(strPhone);             //收件人电话
                    //project.AttrDataList.SaveData();
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("获取厂家模板失败，请联系管理员！");
                    reJo.msg = "获取厂家模板失败，请联系管理员！";
                    return reJo.Value;
                }

                //reJo.data = new JArray(new JObject(new JProperty("ProjectKeyword", project.KeyWord)));

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

        private static bool SetUnitSecUser(string sid,DBSource dbsource,Project prjProject,string strCompanyCode,string strCompanyDesc,ref string strSecretarilman)
        {
            try
            {
                #region 自动设置文控，如果没有与单位代码一样的用户，就自动添加用户
                User secUser = null;
                if (string.IsNullOrEmpty(strSecretarilman))
                {
                    secUser = dbsource.GetUserByCode(strCompanyCode);
                    if (secUser == null)
                    {
                        ////UserController. CreateUser(sid, strCompanyCode, strCompanyDesc + "文控", "", "0",
                        ////            "0", "", strCompanyCode, strCompanyCode);
                        ////secUser = dbsource.GetUserByCode(strCompanyCode);
                        secUser = dbsource.NewUser(
                                            enUserFlage.OnLine,
                                            enUserType.Default,
                                                "",
                                               strCompanyCode,
                                                strCompanyDesc + "文控",
                                                strCompanyCode,
                                                "",
                                                null
                                                );

                        if (secUser != null)
                        {

                            User m_user = secUser;
                            m_user.O_suser1 = m_user.dBSource.GUID;

                            m_user.Modify();



                            // 强制刷新共享数据源
                            //
                            //DBSourceController.RefreshShareDBManager();
                            DBSourceController.RefreshDBSource(sid);
                            strSecretarilman = secUser.ToString;
                        }
                    }
                    else
                    {
                        strSecretarilman = secUser.ToString;
                    }


                }
                else {
                    secUser = dbsource.GetUserByCode(strCompanyCode);
                }

                if (secUser != null)
                {
                    //把用户添加到项目管理类里面的项目单位用户组
                    Group group = dbsource.GetGroupByName(prjProject.Code + "_ALLUnit");

                    if (group != null)
                    {
                        if (!group.UserList.Contains(secUser))
                        {
                            group.AddUser(secUser);
                            group.Modify();
                        }
                    }
                }
                #endregion
            }
            catch { }
            return true;
    }
        private static bool CreateUnitProject(DBSource dbsource, Project prjProject,string strCompanyCode,string strCompanyDesc, string strSecretarilman) {
            #region 自动创建通信类下的单位目录
            try
            {
                Project m_prj = prjProject;
                TempDefn mTempDefn = Company.GetTempDefn(dbsource, "COM_UNIT");
                if (mTempDefn != null)
                {
                    Project cdProject = CommonFunction.GetProjectByDesc(prjProject, "存档管理");
                    if (cdProject != null)
                    {
                        Project txProject = CommonFunction.GetProjectByDesc(cdProject, "通信类");
                        if (txProject != null)
                        {
                            Project project = txProject.NewProject(strCompanyCode, strCompanyDesc, m_prj.Storage, mTempDefn);

                            if (project != null)
                            {
                                //增加附加属性
                                try
                                {
                                    project.GetAttrDataByKeyWord("UN_SECRETAARECTOR").SetCodeDesc(strSecretarilman);             //文控
                                    project.AttrDataList.SaveData();
                                }
                                catch (Exception ex)
                                {

                                }

                                TempDefn sfwTempDefn = Company.GetTempDefn(dbsource, "STO_SUBDOCUMENT");
                                if (sfwTempDefn != null)
                                {
                                    Project swProject = project.NewProject("收文", "", m_prj.Storage, sfwTempDefn);
                                    Project fwProject = project.NewProject("发文", "", m_prj.Storage, sfwTempDefn);

                                    TempDefn typeTempDefn = Company.GetTempDefn(dbsource, "STO_COMTYPE");
                                    if (typeTempDefn != null)
                                    {
                                        if (swProject != null)
                                        {
                                            swProject.NewProject("红头文", "", m_prj.Storage, typeTempDefn);
                                            swProject.NewProject("会议纪要", "", m_prj.Storage, typeTempDefn);
                                            swProject.NewProject("文件传递单", "", m_prj.Storage, typeTempDefn);
                                            swProject.NewProject("信函", "", m_prj.Storage, typeTempDefn);
                                        }

                                        if (fwProject != null)
                                        {
                                            fwProject.NewProject("红头文", "", m_prj.Storage, typeTempDefn);
                                            fwProject.NewProject("会议纪要", "", m_prj.Storage, typeTempDefn);
                                            fwProject.NewProject("文件传递单", "", m_prj.Storage, typeTempDefn);
                                            fwProject.NewProject("信函", "", m_prj.Storage, typeTempDefn);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch
            {

            }
            #endregion
            return  true;
        }

    }
}

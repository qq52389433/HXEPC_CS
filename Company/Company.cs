using System;
using System.Collections.Generic;
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
    public class Company
    {
        /// <summary>
        /// 新建厂家资料目录时，获取默认值
        /// </summary>
        /// <param name="sid"></param>
        /// <returns></returns>
        public static JObject GetEditCompanyDefault(string sid, string ProjectKeyword)
        {
            return EditCompanyInfo.GetEditCompanyDefault(sid, ProjectKeyword);
            
        }



        /// <summary>
        /// 新建厂家资料目录
        /// </summary>
        public static JObject EditCompany(string sid, string ProjectKeyword, string projectAttrJson) {
            return EditCompanyInfo.EditCompany(sid,  ProjectKeyword, projectAttrJson);
           }

        /// <summary>
        /// 新建厂家资料目录
        /// </summary>
        public static JObject CreateCompany(string sid, string ProjectKeyword, string projectAttrJson)
        {
            return EditCompanyInfo.CreateCompany(sid, ProjectKeyword, projectAttrJson);
   
        }

        /// <summary>
        /// 新建厂家资料目录时，获取默认值
        /// </summary>
        /// <param name="sid"></param>
        /// <returns></returns>
        public static JObject GetEditDepartmentDefault(string sid, string ProjectKeyword)
        {
            return EditDepartmentInfo.GetEditDepartmentDefault(sid, ProjectKeyword);

        }



        /// <summary>
        /// 新建厂家资料目录
        /// </summary>
        public static JObject EditDepartment(string sid, string ProjectKeyword, string projectAttrJson)
        {
            return EditDepartmentInfo.EditDepartment(sid, ProjectKeyword, projectAttrJson);
        }

        /// <summary>
        /// 新建厂家资料目录
        /// </summary>
        public static JObject CreateDepartment(string sid, string ProjectKeyword, string projectAttrJson)
        {
            return EditDepartmentInfo.CreateDepartment(sid, ProjectKeyword, projectAttrJson);

        }

        public static JObject GetSelectUnitList(string sid, string ProjectKeyword, string Filter)
        {
            return SelectUnit.GetSelectUnitList(sid, ProjectKeyword, Filter);
        }

        public static JObject GetSelectProfessionList(string sid, string ProjectKeyword, string Filter)
        {
            return SelectProfession.GetSelectProfessionList(sid, ProjectKeyword, Filter);
        }

        public static JObject GetSelectReceiveTypeList(string sid, string ProjectKeyword, string page,string Filter)
        {
            return SelectReceiveType.GetSelectReceiveTypeList(sid, ProjectKeyword,page, Filter);
        }

        public static JObject GetSelectWorkSubList(string sid, string ProjectKeyword, string page, string Filter)
        {
            return SelectWorkSub.GetSelectWorkSubList(sid, ProjectKeyword, page, Filter);
        }

        public static JObject GetSelectDepartmentList(string sid, string ProjectKeyword, string page, string Filter)
        {
            return SelectDepartment.GetSelectDepartmentList(sid, ProjectKeyword, page, Filter);
        }

        public static JObject GetSelectCrewList(string sid, string ProjectKeyword, string page, string Filter)
        {
            return SelectCrew.GetSelectCrewList(sid, ProjectKeyword, page, Filter);

        }
        public static JObject GetSelectFactoryList(string sid, string ProjectKeyword, string page, string Filter)
        {
            return SelectFactory.GetSelectFactoryList(sid, ProjectKeyword, page, Filter);

        }

        public static JObject GetSelectSystemList(string sid, string ProjectKeyword, string page, string Filter)
        {
            return SelectSystem.GetSelectSystemList(sid, ProjectKeyword, page, Filter);

        }

        public static JObject GetSelectProjectList(string sid, string Filter)
        {
            
            return SelectProject.GetSelectProjectList(sid, Filter);
        }


        /// <summary>
        /// 新建厂家资料目录
        /// </summary>
        public static JObject CreateCompanyProject(string sid, string ProjectKeyword, string projectAttrJson)
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

                //获取所有厂家信息
                TempDefn mTempDefn = GetTempDefn(dbsource, "COM_UNIT");
                if (mTempDefn == null)
                {
                    reJo.msg = "获取参建单位模板失败，请联系管理员！";
                    return reJo.Value;
                }

                #region 获取传递过来的属性参数
                //获取传递过来的属性参数
                JArray jaAttr = (JArray)JsonConvert.DeserializeObject(projectAttrJson);

                string strCompanyCode = "", strCompanyDesc = "", companyType = "";

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
                        case "companyType":
                            companyType = strValue;
                            break;
                    }
                }
                #endregion

                if (string.IsNullOrEmpty(strCompanyCode)) {
                    reJo.msg = "新建参建单位目录失败，请输入单位编码！";
                    return reJo.Value;
                }

                Project project = m_prj.NewProject(strCompanyCode, strCompanyDesc, m_prj.Storage, mTempDefn);
                if (project == null)
                {
                    reJo.msg = "新建参建单位目录失败，请联系管理员！";
                    return reJo.Value;
                }

                #region 添加文控角色到新建立的参建单位目录
                //增加附加属性
                try
                {
                    string secretarilMan = "";
                    if (companyType == "项目部门")
                    {
                        //获取所有项目部门（不区分项目）
                        List<DictData> departDdList = dbsource.GetDictDataList("Communication");

                        foreach (DictData data6 in departDdList)
                        {
                            if (data6.O_sValue1.Trim() != strCompanyCode)
                            {
                                continue;
                            }
                            if (!string.IsNullOrEmpty(data6.O_sValue4.Trim()))
                            {
                                secretarilMan = data6.O_sValue4.Trim();
                            }
                        }
                    }
                    else if (companyType == "参建单位")
                    {
                        Project rootProj = CommonFunction.getParentProjectByTempDefn(project, "HXNY_DOCUMENTSYSTEM");
                        if (rootProj != null)
                        {
                            string rootProjCode = rootProj.Code;

                            List<DictData> departDdList = dbsource.GetDictDataList("Unit");

                            foreach (DictData data6 in departDdList)
                            {
                                if (string.IsNullOrEmpty(data6.O_sValue1.Trim()))
                                {
                                    continue;
                                }
                                if (data6.O_sValue1.Trim() != rootProjCode)
                                {
                                    continue;
                                }
                                if (data6.O_Code.Trim() != strCompanyCode)
                                {
                                    continue;
                                }
                                if (!string.IsNullOrEmpty(data6.O_sValue3.Trim()))
                                {
                                    secretarilMan = data6.O_sValue3.Trim();
                                }
                            }
                        }
                    }

                    project.GetAttrDataByKeyWord("UN_SECRETAARECTOR").SetCodeDesc(secretarilMan);    //文控
                    project.AttrDataList.SaveData();
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("获取厂家模板失败，请联系管理员！");
                    reJo.msg = "获取厂家模板失败，请联系管理员！";
                    return reJo.Value;
                } 
                #endregion

                reJo.data = new JArray(new JObject(new JProperty("ProjectKeyword", project.KeyWord)));

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
        /// //获取所有厂家信息
        /// </summary>
        /// <param name="td"></param>
        internal static TempDefn GetTempDefn(DBSource m_dbs,string keyword)
        {
            List<TempDefn> mTempDefnList = m_dbs.GetTempDefnByCode(keyword);
            if (mTempDefnList != null && mTempDefnList.Count > 0)
            {
                return mTempDefnList[0];
            }

            return null;
        }

   


    }


}

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
    public class EditDepartmentInfo
    {
        /// <summary>
        /// 新建厂家资料目录时，获取默认值
        /// </summary>
        /// <param name="sid"></param>
        /// <returns></returns>
        public static JObject GetEditDepartmentDefault(string sid, string ProjectKeyword)
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

                //Project prjProject = CommonFunction.getParentProjectByTempDefn(m_prj, "HXNY_DOCUMENTSYSTEM");

                //if (prjProject == null)
                //{
                //    reJo.msg = "获取项目目录失败！";
                //    return reJo.Value;
                //}

                ////获取项目代码
                //string strProjCode = prjProject.Code;//.GetAttrDataByKeyWord("COMPANY").ToString;
                //string strProjDesc = prjProject.Description;

                JArray jaDepartment = new JArray();
                JObject joDepartment = new JObject();

                List<DictData> dictDataList = dbsource.GetDictDataList("Communication");
                //[o_Code]:公司编码,[o_Desc]：公司描述,[o_sValue1]：项目代码

                foreach (DictData data6 in dictDataList)
                {
                    //if (!string.IsNullOrEmpty(data6.O_sValue1) && data6.O_sValue1 == strProjCode)
                    {
                        joDepartment = new JObject(
                            new JProperty("departmentId", data6.O_ID.ToString()),
                            new JProperty("departmentCode", data6.O_sValue1),
                            new JProperty("departmentDesc", data6.O_Desc),
                            new JProperty("secretarilman", data6.O_sValue4)
                            );
                        jaDepartment.Add(joDepartment);
                    }
                }


                reJo.data = new JArray(
                    new JObject(new JProperty("projectCode", ""),
                    new JProperty("projectDesc", ""),
                    new JProperty("DepartmentList", jaDepartment)));
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
        public static JObject CreateDepartment(string sid, string ProjectKeyword, string projectAttrJson)
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

                string strDepartmentCode = "", strDepartmentDesc = "",
                    strSecretarilman = "", strDepartmentChinese = "",
                    strAddress = "", strProvince = "",
                    strPostCode = "", strEMail = "",
                    strReceiver = "", strFaxNo = "", strPhone = "";

                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    switch (strName)
                    {
                        case "departmentCode":
                            strDepartmentCode = strValue;
                            break;
                        case "departmentDesc":
                            strDepartmentDesc = strValue;
                            break;
                        case "secretarilman":
                            strSecretarilman = strValue;
                            break;
                    }
                }

                if (string.IsNullOrEmpty(strDepartmentCode))
                {
                    reJo.msg = "请输入部门编号！";
                    return reJo.Value;
                }
                if (string.IsNullOrEmpty(strDepartmentDesc))
                {
                    reJo.msg = "请输入部门名称！";
                    return reJo.Value;
                }

                #endregion

                //Project prjProject = CommonFunction.getParentProjectByTempDefn(m_prj, "HXNY_DOCUMENTSYSTEM");

                //if (prjProject == null)
                //{
                //    reJo.msg = "获取项目目录失败！";
                //    return reJo.Value;
                //}

                ////获取项目代码
                //string strProjCode = prjProject.Code;//.GetAttrDataByKeyWord("COMPANY").ToString;

                JArray jaData = new JArray();
                JObject joDepartment = new JObject();

                List<DictData> dictDataList = dbsource.GetDictDataList("Communication");
                //[o_Code]:公司编码,[o_Desc]：公司描述,[o_sValue1]：项目代码

                foreach (DictData data6 in dictDataList)
                {
                   // if (!string.IsNullOrEmpty(data6.O_sValue1) && data6.O_sValue1 == strProjCode &&
                     if (data6.O_sValue1 == strDepartmentCode)
                    {
                        reJo.msg = "已经存在相同的项目部门，请返回重试！";
                        return reJo.Value;
                    }
                }
                //dbsource.NewDictData
                #region 添加到数据字典
                //添加到数据字典

                      ////设置属性的值


                string format = "insert CDMS_DictData (" +
                    "o_parentno,o_datatype,o_ikey,o_skey,o_Code,o_Desc,o_sValue1,o_sValue2,o_sValue3,o_sValue4,o_sValue5,o_iValue1 ,o_iValue2)" +
                    " values ({0},{1},{2},'{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}',{11},{12}" + ")";
                //0,2,0,'Unit','"+strDepartmentCode+"','"+strDepartmentDesc+"','"+strProjCode+ "','','','','',0,0
                format = string.Format(format, new object[] {
                    0,2,0,"Communication","",strDepartmentDesc,strDepartmentCode,"","",strSecretarilman,"",0,0
                     });
                dbsource.DBExecuteSQL(format);

                DBSourceController.refreshDBSource(sid);

                #endregion

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
        public static JObject EditDepartment(string sid, string ProjectKeyword, string projectAttrJson)
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

                string strDepartmentId = "", strDepartmentCode = "", strDepartmentDesc = "",
                    strSecretarilman = "";

                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    switch (strName)
                    {
                        case "departmentId":
                            strDepartmentId = strValue;
                            break;
                        case "departmentCode":
                            strDepartmentCode = strValue;
                            break;
                        case "departmentDesc":
                            strDepartmentDesc = strValue;
                            break;
                        case "secretarilman":
                            strSecretarilman = strValue;
                            break;
                    }
                }

                if (string.IsNullOrEmpty(strDepartmentCode))
                {
                    reJo.msg = "请输入部门编号！";
                    return reJo.Value;
                }
                if (string.IsNullOrEmpty(strDepartmentDesc))
                {
                    reJo.msg = "请输入部门名称！";
                    return reJo.Value;
                }
                #endregion

                //Project prjProject = CommonFunction.getParentProjectByTempDefn(m_prj, "HXNY_DOCUMENTSYSTEM");

                //if (prjProject == null)
                //{
                //    reJo.msg = "获取项目目录失败！";
                //    return reJo.Value;
                //}



                //获取项目代码
                //string strProjCode = prjProject.Code;//.GetAttrDataByKeyWord("COMPANY").ToString;

                int departmentId = Convert.ToInt32(strDepartmentId);

                JArray jaData = new JArray();
                JObject joDepartment = new JObject();

                List<DictData> dictDataList = dbsource.GetDictDataList("Communication");
                //[o_Code]:公司编码,[o_Desc]：公司描述,[o_sValue1]：项目代码

                foreach (DictData data6 in dictDataList)
                {
                    //if (!string.IsNullOrEmpty(data6.O_sValue1) && data6.O_sValue1 == strProjCode
                    if( data6.O_sValue1 == strDepartmentCode && data6.O_ID != departmentId)
                    {
                        reJo.msg = "已经存在相同的项目部门，请返回重试！";
                        return reJo.Value;
                    }
                }
                //dbsource.NewDictData
                #region 添加到数据字典
                //添加到数据字典

                DictData dictData = null;

                foreach (DictData data6 in dictDataList)
                {
                    if (data6.O_ID == departmentId)
                    {
                        dictData = data6;

                    }
                }

                if (dictData == null)
                {
                    reJo.msg = "项目部门ID不存在，请返回重试！";
                    return reJo.Value;

                }

                dictData.O_sValue1 = strDepartmentCode;
                dictData.O_Desc = strDepartmentDesc;
                //dictData.O_sValue1 = strProjCode;
                dictData.O_sValue4 = strSecretarilman;// secretarilman.ToString;//
                dictData.Modify();

                DBSourceController.refreshDBSource(sid);

                #endregion


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

    }
}

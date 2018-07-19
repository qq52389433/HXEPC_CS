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
    public class EditCrew
    {
        /// <summary>
        /// 新建厂家资料目录时，获取默认值
        /// </summary>
        /// <param name="sid"></param>
        /// <returns></returns>
        public static JObject GetEditCrewDefault(string sid, string ProjectKeyword)
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

                JArray jaCrew = new JArray();
                JObject joCrew = new JObject();

                List<DictData> dictDataList = dbsource.GetDictDataList("Crew");
                //[o_Code]:公司编码,[o_Desc]：公司描述,[o_sValue1]：项目代码

                foreach (DictData data6 in dictDataList)
                {
                    if (!string.IsNullOrEmpty(data6.O_sValue2) && data6.O_sValue2 == strProjCode)
                    {
                        joCrew = new JObject(
                            new JProperty("crewId", data6.O_ID.ToString()),
                            new JProperty("crewCode", data6.O_Code),
                            new JProperty("crewDesc", data6.O_Desc),
                            new JProperty("crewEngDesc", data6.O_sValue1)
                            );
                        jaCrew.Add(joCrew);
                    }
                }


                reJo.data = new JArray(
                    new JObject(new JProperty("projectCode", strProjCode),
                    new JProperty("projectDesc", strProjDesc),
                    new JProperty("CrewList", jaCrew)));
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
        public static JObject CreateCrew(string sid, string ProjectKeyword, string crewAttrJson)
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
                JArray jaAttr = (JArray)JsonConvert.DeserializeObject(crewAttrJson);

                string strCrewCode = "", strCrewDesc = "",
                    strCrewEngDesc = "", strCrewChinese = "",
                    strAddress = "", strProvince = "",
                    strPostCode = "", strEMail = "",
                    strReceiver = "", strFaxNo = "", strPhone = "";

                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    switch (strName)
                    {
                        case "crewCode":
                            strCrewCode = strValue;
                            break;
                        case "crewDesc":
                            strCrewDesc = strValue;
                            break;
                        case "crewEngDesc":
                            strCrewEngDesc = strValue;
                            break;
                    }
                }

                if (string.IsNullOrEmpty(strCrewCode))
                {
                    reJo.msg = "请输入项目编号！";
                    return reJo.Value;
                }
                if (string.IsNullOrEmpty(strCrewDesc))
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
                JObject joCrew = new JObject();

                List<DictData> dictDataList = dbsource.GetDictDataList("Crew");
                //[o_Code]:公司编码,[o_Desc]：公司描述,[o_sValue1]：项目代码

                foreach (DictData data6 in dictDataList)
                {
                    if (!string.IsNullOrEmpty(data6.O_sValue2) && data6.O_sValue2 == strProjCode && data6.O_Code == strCrewCode)
                    {
                        reJo.msg = "已经存在相同的参建单位，请返回重试！";
                        return reJo.Value;
                    }
                }
                //dbsource.NewDictData
                #region 添加到数据字典
                //添加到数据字典

                string format = "insert CDMS_DictData (" +
                    "o_parentno,o_datatype,o_ikey,o_skey,o_Code,o_Desc,o_sValue1,o_sValue2,o_sValue3,o_sValue4,o_sValue5,o_iValue1 ,o_iValue2)" +
                    " values ({0},{1},{2},'{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}',{11},{12}" + ")";
                //0,2,0,'Unit','"+strCrewCode+"','"+strCrewDesc+"','"+strProjCode+ "','','','','',0,0
                format = string.Format(format, new object[] {
                    0,2,0,"Crew",strCrewCode,strCrewDesc,strCrewEngDesc,strProjCode,"","","",0,0
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
        public static JObject EDITCrew(string sid, string ProjectKeyword, string crewAttrJson)
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
                JArray jaAttr = (JArray)JsonConvert.DeserializeObject(crewAttrJson);

                string strCrewId = "", strCrewCode = "", strCrewDesc = "",
                    strCrewEngDesc = "", strCrewChinese = "",
                    strAddress = "", strProvince = "",
                    strPostCode = "", strEMail = "",
                    strReceiver = "", strFaxNo = "", strPhone = "";

                foreach (JObject joAttr in jaAttr)
                {
                    string strName = joAttr["name"].ToString();
                    string strValue = joAttr["value"].ToString();

                    switch (strName)
                    {
                        case "crewId":
                            strCrewId = strValue;
                            break;
                        case "crewCode":
                            strCrewCode = strValue;
                            break;
                        case "crewDesc":
                            strCrewDesc = strValue;
                            break;
                        case "crewEngDesc":
                            strCrewEngDesc = strValue;
                            break;
                    }
                }

                if (string.IsNullOrEmpty(strCrewCode))
                {
                    reJo.msg = "请输入项目编号！";
                    return reJo.Value;
                }
                if (string.IsNullOrEmpty(strCrewDesc))
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


                int crewId = Convert.ToInt32(strCrewId);
                //获取项目代码
                string strProjCode = prjProject.Code;//.GetAttrDataByKeyWord("COMPANY").ToString;


                JArray jaData = new JArray();
                JObject joCrew = new JObject();

                List<DictData> dictDataList = dbsource.GetDictDataList("Crew");
                //[o_Code]:公司编码,[o_Desc]：公司描述,[o_sValue1]：项目代码

                foreach (DictData data6 in dictDataList)
                {
                    if (!string.IsNullOrEmpty(data6.O_sValue2) && data6.O_sValue2 == strProjCode
                        && data6.O_Code == strCrewCode && data6.O_ID != crewId)
                    {
                        reJo.msg = "已经存在相同的参建单位，请返回重试！";
                        return reJo.Value;
                    }
                }

                #region 添加到数据字典

                DictData dictData = null;

                foreach (DictData data6 in dictDataList)
                {
                    if (data6.O_ID == crewId)
                    {
                        dictData = data6;

                    }
                }

                if (dictData == null)
                {
                    reJo.msg = "参建单位ID不存在，请返回重试！";
                    return reJo.Value;

                }

                dictData.O_Code = strCrewCode;
                dictData.O_Desc = strCrewDesc;
                dictData.O_sValue2 = strProjCode;
                dictData.O_sValue1 = strCrewEngDesc;
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

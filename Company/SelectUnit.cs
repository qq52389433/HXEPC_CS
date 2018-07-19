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
    public class SelectUnit
    {
        public static JObject GetSelectUnitList(string sid, string ProjectKeyword, string Filter)
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

                Filter = Filter.Trim().ToLower();

                string curUnitCode = "";

                curUnitCode = prjProject.Code;//.GetValueByKeyWord("PRO_COMPANY");
                if (string.IsNullOrEmpty(curUnitCode))
                {
                    reJo.msg = "获取项目来源失败！";
                    return reJo.Value;
                }

                JArray jaData = new JArray();


                #region 获取项目的通信代码
                AttrData data;
                if ((data = prjProject.GetAttrDataByKeyWord("RPO_ONSHORE")) != null)
                {
                    string strData = data.ToString;
                    if (!string.IsNullOrEmpty(strData))
                    {
                        JObject joData = new JObject(
                            new JProperty("unitType", "通信代码"),
                            new JProperty("unitId", data.DefnID),
                            new JProperty("unitCode", data.ToString),
                            new JProperty("unitDesc", "OnShore")
                            );
                        jaData.Add(joData);
                    }
                }
                if ((data = prjProject.GetAttrDataByKeyWord("RPO_OFFSHORE")) != null)
                {
                    string strData = data.ToString;
                    if (!string.IsNullOrEmpty(strData))
                    {
                        JObject joData = new JObject(
                            new JProperty("unitType", "通信代码"),
                            new JProperty("unitId", data.DefnID),
                            new JProperty("unitCode", data.ToString),
                            new JProperty("unitDesc", "OffShore")
                            );
                        jaData.Add(joData);
                    }
                } 
                #endregion
                #region 获取所有参建单位
                //获取所有参建单位
                List<DictData> tmpDataList = dbsource.GetDictDataList("Unit");

                List<DictData> dictDataList = new List<DictData>();

                //过滤不是当前项目的参建单位，提高排序速度
                foreach (DictData data6 in tmpDataList)
                {
                    //判断是否符合过滤条件
                    if (data6.O_sValue1 == curUnitCode)
                    {
                        dictDataList.Add(data6);
                    }

                }
                //按代码排序
                dictDataList.Sort(delegate (DictData x, DictData y)
                {
                    return x.O_Code.CompareTo(y.O_Code);
                });

                foreach (DictData data6 in dictDataList)
                {
                    //判断是否符合过滤条件
                    if (!string.IsNullOrEmpty(Filter) &&
                        data6.O_Code.ToLower().IndexOf(Filter) < 0 && data6.O_Desc.ToLower().IndexOf(Filter) < 0)
                    {
                        continue;
                    }

                    if (data6.O_sValue1 == curUnitCode)
                    {
                        JObject joData = new JObject(
                            new JProperty("unitType", "参建单位"),
                            new JProperty("unitId", data6.O_ID.ToString()),
                            new JProperty("unitCode", data6.O_Code),
                            new JProperty("unitDesc", data6.O_Desc)
                            );
                        jaData.Add(joData);
                    }
                }
                #endregion



                //获取所有项目部门（区分项目）
                //List<DictData> departDdList = dbsource.GetDictDataList("DeparDate");

                //foreach (DictData data6 in departDdList)
                //{
                //    //if (data6.O_sValue1 == curUnitCode)
                //    {
                //        JObject joData = new JObject(
                //            new JProperty("unitType", "项目部门"),
                //            new JProperty("unitId", data6.O_ID.ToString()),
                //            new JProperty("unitCode", data6.O_Code),
                //            new JProperty("unitDesc", data6.O_Desc)
                //            );
                //          jaData.Add(joData);
                //    }
                //}

                #region 获取所有项目部门（不区分项目）
                //获取所有项目部门（不区分项目）
                List<DictData> departDdList = dbsource.GetDictDataList("Communication");

                //按代码排序
                departDdList.Sort(delegate (DictData x, DictData y)
                {
                    return x.O_sValue1.CompareTo(y.O_sValue1);
                });

                foreach (DictData data6 in departDdList)
                {
                    //判断是否符合过滤条件
                    if (!string.IsNullOrEmpty(Filter) &&
                        data6.O_sValue1.ToLower().IndexOf(Filter) < 0 && data6.O_Desc.ToLower().IndexOf(Filter) < 0)
                    {
                        continue;
                    }

                    //if (data6.O_sValue1 == curUnitCode)
                    if (!string.IsNullOrEmpty(data6.O_sValue1.Trim()))
                    {
                        JObject joData = new JObject(
                            new JProperty("unitType", "项目部门"),
                            new JProperty("unitId", data6.O_ID.ToString()),
                            new JProperty("unitCode", data6.O_sValue1),
                            new JProperty("unitDesc", data6.O_Desc)
                            );
                        jaData.Add(joData);
                    }
                } 
                #endregion

                reJo.data = jaData;
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

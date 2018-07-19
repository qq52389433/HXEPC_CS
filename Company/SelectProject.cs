using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using AVEVA.CDMS.Server;
using AVEVA.CDMS.Common;
using AVEVA.CDMS.WebApi;

namespace AVEVA.CDMS.HXEPC_Plugins
{
    public class SelectProject
    {
        public static JObject GetSelectProjectList(string sid, string Filter)
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

                Project m_RootProject = dbsource.RootLocalProjectList.Find(itemProj => itemProj.TempDefn.KeyWord == "PRODOCUMENTADMIN");
                if (m_RootProject == null)
                {
                    reJo.msg = "[项目管理类文件目录]不存在,或者未添加目录模板！不能获取项目列表";
                    return reJo.Value;
                }

                JArray jaData = new JArray();
                foreach (Project proj in m_RootProject.AllProjectList) {

                    //判断是否符合过滤条件
                    if (!string.IsNullOrEmpty(Filter) &&
                        proj.Code.ToLower().IndexOf(Filter) < 0 && proj.Description.ToLower().IndexOf(Filter) < 0)
                    {
                        continue;
                    }

                    if (proj.TempDefn.KeyWord == "HXNY_DOCUMENTSYSTEM")
                    {
                        JObject joData = new JObject(
                                new JProperty("projectType", "项目"),
                                new JProperty("projectId", proj.KeyWord),
                                new JProperty("projectCode", proj.Code),
                                new JProperty("projectDesc", proj.Description)
                                );
                        jaData.Add(joData);
                    }
                }

                

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

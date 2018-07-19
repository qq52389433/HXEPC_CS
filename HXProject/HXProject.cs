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
    public class HXProject
    {

        public static JObject GetCreateProjectListingDefault(string sid) {
            return CreateProject.GetCreateProjectListingDefault(sid);
        }

        public static JObject GetProjectTypeII(string sid, string ProjectType) {
            return CreateProject.GetProjectTypeII(sid, ProjectType);
        }

        public static JObject CreateRootProject(string sid, string projectAttrJson) {
            return CreateProject.CreateRootProject(sid, projectAttrJson);
        }

        public static JObject CreateProjectListing(string sid, string ProjectKeyword, string projectAttrJson) {
            return CreateProject.CreateProjectListing(sid, ProjectKeyword, projectAttrJson);
        }

        public static JObject GetEditCrewDefault(string sid, string ProjectKeyword) {
            return EditCrew.GetEditCrewDefault(sid, ProjectKeyword);
        }

        public static JObject EDITCrew(string sid, string ProjectKeyword, string crewAttrJson)
        {
            return EditCrew.EDITCrew(sid, ProjectKeyword, crewAttrJson);
        }

        public static JObject CreateCrew(string sid, string ProjectKeyword, string crewAttrJson)
        {
            return EditCrew.CreateCrew(sid, ProjectKeyword, crewAttrJson);
        }

        public static JObject GetEditSystemDefault(string sid, string ProjectKeyword)
        {
            return EditSystem.GetEditSystemDefault(sid, ProjectKeyword);
        }

        public static JObject EDITSystem(string sid, string ProjectKeyword, string systemAttrJson)
        {
            return EditSystem.EDITSystem(sid, ProjectKeyword, systemAttrJson);
        }

        public static JObject CreateSystem(string sid, string ProjectKeyword, string systemAttrJson)
        {
            return EditSystem.CreateSystem(sid, ProjectKeyword, systemAttrJson);
        }

        public static JObject GetEditFactoryDefault(string sid, string ProjectKeyword)
        {
            return EditFactory.GetEditFactoryDefault(sid, ProjectKeyword);
        }

        public static JObject EDITFactory(string sid, string ProjectKeyword, string factoryAttrJson)
        {
            return EditFactory.EDITFactory(sid, ProjectKeyword, factoryAttrJson);
        }

        public static JObject CreateFactory(string sid, string ProjectKeyword, string factoryAttrJson)
        {
            return EditFactory.CreateFactory(sid, ProjectKeyword, factoryAttrJson);
        }
    }
}

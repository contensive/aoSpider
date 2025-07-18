using Contensive.BaseClasses;

namespace Contensive.Addons.Spider {
    public class SpiderCleanUp : AddonBaseClass {
        public override object Execute(CPBaseClass CP) {


            CP.Db.ExecuteNonQuery("delete from ccspiderdocs from ccspiderdocs  left join cclinkaliases on ccSpiderDocs.pageid = cclinkaliases.pageid  where (ccSpiderDocs.pageid <> 0) and ccLinkAliases.id is null");
            return default;

        }

    }
}
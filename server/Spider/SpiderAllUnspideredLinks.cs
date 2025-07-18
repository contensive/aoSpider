using Contensive.BaseClasses;

namespace Contensive.Addons.Spider {

    public class SpiderAllUnspideredLinks : AddonBaseClass {

        public override object Execute(CPBaseClass CP) {
            CP.Doc.SetProperty("Spider Where Clause", "spidered=0  or spidered is null");
            CP.Doc.SetProperty("Spider Count", 9999);
            CP.Addon.Execute("{A5B29F03-4FEE-432F-8F34-704B7FB03560}");
            return default;
        }
    }
}
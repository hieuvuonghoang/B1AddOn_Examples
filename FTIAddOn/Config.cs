
namespace FTIAddOn
{
    public class Config
    {
        private SAPbouiCOM.Application SBO_Application;
        private SAPbobsCOM.Company oCompany;

        public Config(SAPbouiCOM.Application SBO_Application, SAPbobsCOM.Company oCompany)
        {
            this.SBO_Application = SBO_Application;
            this.oCompany = oCompany;
        }

        public void CreateMenu()
        {
        }

    }

}

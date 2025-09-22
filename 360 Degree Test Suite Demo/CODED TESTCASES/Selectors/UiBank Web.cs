

using UiPath.CodedWorkflows.DescriptorIntegration;
using UiPath.UIAutomationNext;
using UiPath.UIAutomationNext.API.Models;
using UiPath.UIAutomationNext.Enums;
using Target = UiPath.UIAutomationNext.API.Models.Target;

namespace _360DegreeTestSuiteDemo.CODEDTESTCASES.Selectors
{
    public class UiBank_Web : CodedWorkflow
    {
        
        private static NBrowserType myBrowser = NBrowserType.Edge;

        public NBrowserType browser
        {
            get
            {
                return myBrowser; // Getter returns the value of the private member variable
            }
            set
            {
                myBrowser = value; // Setter sets the value of the private member variable
            }
        }

        //Screens

        public  TargetAppModel homeScreen = AppTarget.FromUrl("https://uibank.uipath.com/welcome", "<html  app='msedge.exe' title='UiBank' />", myBrowser);
        public TargetAppModel selectionScreen = AppTarget.FromUrl("https://uibank.uipath.com/loans", "<html app='msedge.exe' title='UiBank' />", myBrowser);
        public  TargetAppModel formScreen = AppTarget.FromUrl("https://uibank.uipath.com/loans/apply", "<html app='msedge.exe' title='UiBank' />", myBrowser);
        public  TargetAppModel quoteScreen = AppTarget.FromUrl("https://uibank.uipath.com/lonas/result", "<html app='msedge.exe' title='UiBank' />", myBrowser);


        //Elements

        //HomeScreen
        public  TargetAnchorableModel products = Target.FromSelector("<webctrl aaname='Products' />");
        public  TargetAnchorableModel loans = Target.FromSelector("<webctrl aaname='Loans' />");

        //SelectionScreen
        public  TargetAnchorableModel applyForLoan = Target.FromSelector("<webctrl id='applyButton'/>");

        //FormSCreen
        public  TargetAnchorableModel email = Target.FromSelector("<webctrl id='email'/>");
        public  TargetAnchorableModel loanAmount = Target.FromSelector("<webctrl id='amount'/>");
        public  TargetAnchorableModel loanTerm = Target.FromSelector("<webctrl id='term'/>");
        public  TargetAnchorableModel income = Target.FromSelector("<webctrl id='income'/>");
        public  TargetAnchorableModel age = Target.FromSelector("<webctrl id='age'/>");
        public  TargetAnchorableModel submit = Target.FromSelector("<webctrl id='submitButton'/>");

        //QuoteScreen
        public  TargetAnchorableModel loanRate = Target.FromSelector("<webctrl id='rateValue'/>");



    }
}
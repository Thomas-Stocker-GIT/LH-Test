using _360DegreeTestSuiteDemo.ObjectRepository;
using System;
using System.Collections.Generic;
using System.Data;
using UiPath.CodedWorkflows;
using UiPath.Core;
using UiPath.Core.Activities.Storage;
using UiPath.Excel;
using UiPath.Excel.Activities;
using UiPath.Excel.Activities.API;
using UiPath.Excel.Activities.API.Models;
using UiPath.Orchestrator.Client.Models;
using UiPath.Testing;
using UiPath.Testing.Activities.TestData;
using UiPath.Testing.Activities.TestDataQueues.Enums;
using UiPath.Testing.Enums;
using UiPath.UIAutomationNext.API.Contracts;
using UiPath.UIAutomationNext.API.Models;
using UiPath.UIAutomationNext.Enums;

namespace _360DegreeTestSuiteDemo.CODEDTESTCASES
{
    
    using app = _360DegreeTestSuiteDemo.ObjectRepository.Descriptors.UiBank_Web;
    
    
    public class Coded_ApplyForLoan : CodedWorkflow
    {       
        
        [TestCase]
        public void Execute()
        {
            var homeScreen = uiAutomation.Open("Home");
            homeScreen.Click("Products");
            homeScreen.Click("Loans");
            
            var selectionScreen = uiAutomation.Attach(app.Selection);
            selectionScreen.Click(app.Selection.Apply_For_A_Loan);
            
            var formScreen = uiAutomation.Attach(app.Form);
            formScreen.TypeInto(app.Form.Email_Address, "john@doe.com");
            formScreen.TypeInto(app.Form.Loan_Amount_Requested, "10000");
            formScreen.SelectItem(app.Form.Loan_Term, "3");
            formScreen.TypeInto(app.Form.Current_Yearly_Income, "20000");
            formScreen.TypeInto(app.Form.Age, "40");
            formScreen.Click(app.Form.Submit_Loan_Application);
            
            var quoteScreen = uiAutomation.Attach("Quote");
            var loanRate = quoteScreen.GetText(app.Quote.Loan_Rate);
            
            quoteScreen.TakeScreenshot(app.Quote.Loan_Rate,"loanrate.png");
            testing.VerifyAreEqual(loanRate, "9");
            quoteScreen.Dispose();
            
            
            
            
        }
    }
}
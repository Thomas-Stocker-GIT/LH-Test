using _360DegreeTestSuiteDemo.CODEDTESTCASES.Selectors;
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
    public class CodedTestCase_ApplyForLoan_NativeSelectors : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            UiBank_Web selectors = new UiBank_Web();
            
            var homeScreen = uiAutomation.Open(selectors.homeScreen);
            homeScreen.Click(selectors.products);
            homeScreen.Click(selectors.loans);
            
            var selectionScreen = uiAutomation.Attach(selectors.selectionScreen);
            selectionScreen.Click(selectors.applyForLoan);
            
            var formScreen = uiAutomation.Attach(selectors.formScreen);
            formScreen.TypeInto(selectors.email, "john@doe.com");
            formScreen.TypeInto(selectors.loanAmount, "10000");
            formScreen.SelectItem(selectors.loanTerm, "3");
            formScreen.TypeInto(selectors.income, "20000");
            formScreen.TypeInto(selectors.age, "40");
            formScreen.Click(selectors.submit);
            
            var quoteScreen = uiAutomation.Attach(selectors.quoteScreen);
            var loanRate = quoteScreen.GetText(selectors.loanRate);
            testing.VerifyAreEqual(loanRate, "9");
            quoteScreen.Dispose();
        }
    }
}
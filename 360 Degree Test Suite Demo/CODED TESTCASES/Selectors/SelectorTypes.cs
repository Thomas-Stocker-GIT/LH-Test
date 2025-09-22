using _360DegreeTestSuiteDemo.ObjectRepository;
using System;
using System.Collections.Generic;
using System.Data;
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
using UiPath.UIAutomationNext.Models.CV;

namespace _360DegreeTestSuiteDemo.CODEDTESTCASES.Selectors
{
    public class SelectorTypes
    {
        
        // Strict Selector
        public static readonly TargetAnchorableModel applyForLoanButtonStrict =
            Target.FromSelector("<webctrl id='applyButton'/>");
        
        // Fuzzy Selector
         public static readonly TargetAnchorableModel applyForLoanButtonFuzzy =
            Target.FromFuzzySelector("<webctrl id='applyButton' tag='BUTTON' type='submit' class='btn btn-white btn-xl' aaname='Apply For A Loan' />");
        
        // Image Selector
         public static readonly TargetAnchorableModel applyForLoanButtonImage = 
            Target.FromImage(@"C:\Users\thomas.stocker\OneDrive - UiPath\Documents\UiPath\360 Degree Test Suite Demo\.screenshots\ApplyForLoan.png", 0.8);
        
        // Computer Vision Selector
        public static readonly TargetAnchorableModel applyForLoanButtonCV =
            Target.FromComputerVision(@"C:\Users\thomas.stocker\OneDrive - 
            UiPath\Documents\UiPath\360 Degree Test Suite Demo\.screenshots\9b32e2f5656b5d44939d02aac9d15f61.png",UIVisionCategoryType.Button, "Apply For A Loan");
        
    }
}
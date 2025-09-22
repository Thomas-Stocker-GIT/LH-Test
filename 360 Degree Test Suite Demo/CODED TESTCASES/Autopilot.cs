using System;
using System.Collections.Generic;
using System.Data;
using _360DegreeTestSuiteDemo.ObjectRepository;
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
    public class Autopilot : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
        }

        // Create a method to compare strings for similarity using Levenstein arlgorithm named 'CompareStrings' with the input parameters: "string1, string"
        
        
        
        
        
       // Improve the performance of the following method:
        
        public static int SequentialSearch(List<int> list, int target)
        {
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i] == target)
                {
                    Console.WriteLine("Found the target.");
                    return i;
                }
            }

            Console.WriteLine("Target not found.");
            return -1;
        }

        

    // Click on Products
    // Click on Loans
    // Click on Apply For A Loan
    // Fill in Form with email: john@doe.com, loan amount requested: 10000, loan term: 3, yearly income: 20000, age: 40
    // Click on Submit Loan
    }
}
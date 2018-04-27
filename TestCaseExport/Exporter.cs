using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using Microsoft.TeamFoundation.TestManagement.Client;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Text;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace TestCaseExport
{
    /// <summary>
    /// Exports a passed set of test cases to the supplied file.
    /// </summary>
    public class Exporter
    {
        public void Export(string filename, ITestPlan testPlan)
        {
            using (var pkg = new ExcelPackage())
            {
                var sheet = pkg.Workbook.Worksheets.Add("Test Script");
                sheet.Cells[1, 1].Value = "Test Case ID";
                sheet.Cells[1, 2].Value = "User Story ID";
                sheet.Cells[1, 3].Value = "Test Condition";
                sheet.Cells[1, 4].Value = "Step Number";
                sheet.Cells[1, 5].Value = "Action/Description";
                sheet.Cells[1, 6].Value = "Attachment Name";
                sheet.Cells[1, 7].Value = "Expected Result";

                sheet.Column(1).Width = 15;
                sheet.Column(2).Width = 15;
                sheet.Column(3).Width = 15;
                sheet.Column(4).Width = 15;
                sheet.Column(5).Width = 50;
                sheet.Column(6).Width = 25;
                sheet.Column(7).Width = 50;

                int row = 2;

                //export all test cases from root Test Plan
                //foreach (var testCase in testSuite.AllTestCases)
                foreach (var testCase in testPlan.RootSuite.AllTestCases)
                {
                    var replacementSets = GetReplacementSets(testCase);
                    foreach (var replacements in replacementSets)
                    {
                        int teststepcounter = 1;
                        var firstRow = row;
                        foreach (var testAction in testCase.Actions)
                        {
                            AddSteps(sheet, testAction, replacements, ref row, ref teststepcounter);
                        }
                        if (firstRow != row)
                        {
                            var mergedID = sheet.Cells[firstRow, 1, row - 1, 1];
                            mergedID.Merge = true;
                            mergedID.Value = testCase.WorkItem == null ? "" : testCase.WorkItem.Id.ToString();
                            mergedID.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            mergedID.Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                            var mergedText = sheet.Cells[firstRow, 3, row - 1, 3];
                            mergedText.Merge = true;
                            CleanupText(mergedText, testCase.Title, replacements);
                            mergedText.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            mergedText.Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                            var userStoryText = sheet.Cells[firstRow, 2, row - 1, 2];
                            userStoryText.Merge = true;

                            //find all linked stories via "Tests" link type instead of "first match"
                            string linkedStories = testCase.WorkItem.WorkItemLinks.Cast<WorkItemLink>().Where<WorkItemLink>(x => x.LinkTypeEnd.Name == "Tests").ToList().Select<WorkItemLink, string>((a, b) => a.TargetId.ToString()).Aggregate((a, b) => a + "\n" + b);

                            CleanupText(userStoryText, linkedStories, replacements);
                            userStoryText.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            userStoryText.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        }
                    }
                }

                var header = sheet.Cells[1, 1, 1, 8];
                header.Style.Font.Bold = true;
                header.Style.Fill.PatternType = ExcelFillStyle.Solid;
                header.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 226, 238, 18));

                sheet.Cells[1, 1, row, 8].Style.WrapText = true;

                pkg.SaveAs(new FileInfo(filename));
            }
        }

        private List<Dictionary<string, string>> GetReplacementSets(ITestCase testCase)
        {
            var replacementSets = new List<Dictionary<string, string>>();

            try
            {
                foreach (DataRow r in testCase.DefaultTableReadOnly.Rows)
                {
                    var replacement = new Dictionary<string, string>();
                    foreach (DataColumn c in testCase.DefaultTableReadOnly.Columns)
                    {
                        replacement[c.ColumnName] = r[c] as string;
                    }
                    replacementSets.Add(replacement);
                }
                return replacementSets.DefaultIfEmpty(new Dictionary<string, string>()).ToList();
            }
            catch (System.Exception)
            {
                //swallow exception
                return replacementSets.DefaultIfEmpty(new Dictionary<string, string>()).ToList();
            }
        }

        private void AddSteps(ExcelWorksheet xlWorkSheet, ITestAction testAction, Dictionary<string, string> replacements, ref int row, ref int teststepCounter)
        {
            var testStep = testAction as ITestStep;
            var group = testAction as ITestActionGroup;
            var sharedRef = testAction as ISharedStepReference;
            if (null != testStep)
            {
                CleanupText(xlWorkSheet.Cells[row, 4], "test step " + teststepCounter, replacements);
                CleanupText(xlWorkSheet.Cells[row, 5], testStep.Title.ToString(), replacements);

                if (testStep.Attachments.Count > 0)
                {
                    string attachmentNames = testStep.Attachments.Select(o => o.Name).Aggregate((a, b) => a + "," + b);
                    CleanupText(xlWorkSheet.Cells[row, 6], attachmentNames, replacements);
                }

                CleanupText(xlWorkSheet.Cells[row, 7], testStep.ExpectedResult.ToString(), replacements);
                teststepCounter++;
                row++;
            }
            else if (null != group)
            {
                foreach (var action in group.Actions)
                {
                    AddSteps(xlWorkSheet, action, replacements, ref row, ref teststepCounter);
                }
            }
            else if (null != sharedRef)
            {
                var step = sharedRef.FindSharedStep();
                foreach (var action in step.Actions)
                {
                    AddSteps(xlWorkSheet, action, replacements, ref row, ref teststepCounter);
                }
            }
        }

        private void CleanupText(ExcelRangeBase cell, string input, Dictionary<string, string> replacements)
        {
            foreach (var kvp in replacements)
            {
                input = input.Replace("@" + kvp.Key, kvp.Value);
            }

            new HtmlToRichTextHelper().HtmlToRichText(cell, input);
        }
    }
}

using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using FuzzySharp;

namespace CharlieChatBotApplication.Bots
{
    public class EchoBot : ActivityHandler
    {
        private readonly Dictionary<string, string> _qna;
        private readonly List<Holiday> _holidays;
        private readonly List<LeaveDetails> _leaveDetailsList;

        public EchoBot()
        {
            _qna = LoadQnAFromExcel("D:\\VS Projects\\CharlieChatBotApplication\\CharlieChatBotApplication\\bin\\AskCharlie.xlsx");
            _holidays = LoadHolidaysFromExcel("D:\\VS Projects\\CharlieChatBotApplication\\CharlieChatBotApplication\\bin\\AskCharlie2.xlsx");
            _leaveDetailsList = LoadLeaveDetailsFromDataSource(); // Load leave details from a data source
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var userQuestion = turnContext.Activity.Text?.Trim().ToLower();

            // Check if the length of the user's input is sufficient for matching
            if (userQuestion.Length < 8) // Adjust the threshold as needed
            {
                await turnContext.SendActivityAsync(MessageFactory.Text("Please provide a longer question."), cancellationToken);
                return;
            }

            // Check if the user's question is specifically asking for leave balance details
            if (userQuestion.StartsWith("leave balance for "))
            {
                string employeeName = userQuestion.Replace("leave balance for ", "").Trim();
                string leaveBalance = GetLeaveDetails(employeeName);
                if (!string.IsNullOrEmpty(leaveBalance))
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text(leaveBalance), cancellationToken);
                }
                else
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text($"Sorry, I couldn't find the leave balance details for {employeeName}."), cancellationToken);
                }
                return;
            }

            // Check if the user's question is asking for the next holiday
            if (userQuestion.Contains("next holiday"))
            {
                var nextHoliday = GetNextHoliday();
                if (nextHoliday != null)
                {
                    string holidayText = $"Next holiday: {nextHoliday.Name} on {nextHoliday.Date.ToShortDateString()}";
                    await turnContext.SendActivityAsync(MessageFactory.Text(holidayText), cancellationToken);
                }
                else
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text("There are no upcoming holidays."), cancellationToken);
                }
                return;
            }

            // Use Fuzzy Matching to get the closest question
            var bestMatch = _qna.Keys.Select(q => new { Question = q, Score = Fuzz.PartialRatio(userQuestion, q) })
                                     .OrderByDescending(m => m.Score)
                                     .FirstOrDefault();

            if (bestMatch != null && bestMatch.Score > 70) // Threshold for fuzzy matching, can be adjusted
            {
                await turnContext.SendActivityAsync(MessageFactory.Text(_qna[bestMatch.Question]), cancellationToken);
            }
            else
            {
                await turnContext.SendActivityAsync(MessageFactory.Text("Sorry, I don't know the answer to that."), cancellationToken);
            }
        }

        public Dictionary<string, string> LoadQnAFromExcel(string filePath)
        {
            var qnaDictionary = new Dictionary<string, string>();
            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row

                    foreach (var row in rows)
                    {
                        string question = row.Cell(1).GetValue<string>().ToLower().Trim(); // Normalize and trim case
                        string response = row.Cell(2).GetValue<string>().Trim(); // Trim response

                        // Handle synonyms or additional forms of the question
                        qnaDictionary[question] = response;
                        if (row.CellCount() > 2) // Assuming synonyms are in the third column
                        {
                            var synonyms = row.Cell(3).GetValue<string>().Split(',');
                            foreach (var synonym in synonyms)
                            {
                                string synonymNormalized = synonym.Trim().ToLower(); // Normalize synonym
                                if (!qnaDictionary.ContainsKey(synonymNormalized))
                                {
                                    qnaDictionary[synonymNormalized] = response; // Add synonym with the same response as the original question
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading Q&A from Excel: {ex.Message}");
                // Handle the exception according to your application's logic
            }
            return qnaDictionary;
        }

        public List<Holiday> LoadHolidaysFromExcel(string filePath)
        {
            var holidayList = new List<Holiday>();
            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row

                    foreach (var row in rows)
                    {
                        DateTime date = row.Cell(1).GetValue<DateTime>(); // Get holiday date
                        string name = row.Cell(2).GetValue<string>(); // Get holiday name

                        holidayList.Add(new Holiday { Date = date, Name = name });
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading holidays from Excel: {ex.Message}");
                // Handle the exception according to your application's logic
            }
            return holidayList;
        }

        public List<LeaveDetails> LoadLeaveDetailsFromDataSource()
        {
            // Method to load leave details from a data source (implementation not provided)
            // This could involve querying a database, reading from an Excel file, etc.
            // Return the leave details as a list of LeaveDetails objects
            // For demonstration purposes, let's assume you have a hard-coded list of leave details

            var leaveDetailsList = new List<LeaveDetails>
            {
                new LeaveDetails { EmployeeName = "Vinod Sai", LeaveType = "Vacation Leave", LeaveBalance = 15, LeaveTaken = 5, LeaveAccruals = 2, Adjustments = 0 },
                new LeaveDetails { EmployeeName = "Akhil", LeaveType = "Sick Leave", LeaveBalance = 10, LeaveTaken = 2, LeaveAccruals = 1, Adjustments = 0 },
                new LeaveDetails { EmployeeName = "Yogesh", LeaveType = "Personal Leave", LeaveBalance = 8, LeaveTaken = 3, LeaveAccruals = 1, Adjustments = 0 }
                // Add more leave details as needed
            };

            return leaveDetailsList;
        }

        private string GetLeaveDetails(string employeeName)
        {
            var leaveDetails = _leaveDetailsList.FirstOrDefault(ld => ld.EmployeeName.ToLower() == employeeName.ToLower());
            if (leaveDetails != null)
            {
                return $"Leave details for {employeeName}:\n\n" +
                       $"Leave Type: {leaveDetails.LeaveType}\n" +
                       $"Leave Balance: {leaveDetails.LeaveBalance}\n" +
                       $"Leave Taken: {leaveDetails.LeaveTaken}\n" +
                       $"Leave Accruals: {leaveDetails.LeaveAccruals}\n" +
                       $"Adjustments: {leaveDetails.Adjustments}";
            }
            else
            {
                return $"Sorry, I couldn't find the leave details for {employeeName}.";
            }
        }

        private Holiday GetNextHoliday()
        {
            var today = DateTime.Today;

            var nextHoliday = _holidays.Where(h => h.Date >= today)
                                       .OrderBy(h => h.Date)
                                       .FirstOrDefault();

            return nextHoliday;
        }

        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            var welcomeText = "Hello and welcome!";
            foreach (var member in membersAdded)
            {
                if (member.Id != turnContext.Activity.Recipient.Id)
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text(welcomeText), cancellationToken);
                }
            }
        }

        // Class representing a holiday
        public class Holiday
        {
            public DateTime Date { get; set; }
            public string Name { get; set; }
        }

        // Class representing leave details
        public class LeaveDetails
        {
            public string EmployeeName { get; set; }
            public string LeaveType { get; set; }
            public int LeaveBalance { get; set; }
            public int LeaveTaken { get; set; }
            public int LeaveAccruals { get; set; }
            public int Adjustments { get; set; }
        }
    }
}

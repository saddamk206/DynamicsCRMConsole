using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using System.ComponentModel;
using System.Data;
using System.Xml.Linq;
using System.Xml;

namespace MergeAccRecords
{
    class Program
    {
        private static IOrganizationService organizationService = null;
        static void Main(string[] args)
        {

            try
            {
                // e.g. https://yourorg.crm.dynamics.com
                string url = "https://*******.crm.dynamics.com/";
                // e.g. you@yourorg.onmicrosoft.com
                string userName = "S***********@*****.com";
                // e.g. y0urp455w0rd 
                string password = "********";

                string conn = $@"
    Url = {url};
    AuthType = OAuth;
    UserName = {userName};
    Password = {password};
    AppId = 51f81489-12ee-4a9e-aaae-a2591f45987d;
    RedirectUri = app://58145B91-0C36-4500-8554-080854F2AC97;
    LoginPrompt=Auto;
    RequireNewInstance = True";

                using (var service = new CrmServiceClient(conn))
                {
                    organizationService = service;
                    string xml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' no-lock='false' distinct='true'><entity name='opportunity'><attribute name='name'/><order attribute='name' descending='false'/><attribute name='parentaccountid'/><attribute name='ownerid'/><attribute name='modifiedon'/><attribute name='im360_createdon'/><attribute name='opportunityid'/><filter type='and'><condition attribute='ownerid' operator='eq-userid'/><condition attribute='statecode' operator='eq' value='0'/><condition attribute='im360_lineofbusiness2name' operator='like' value='%itad%'/></filter><link-entity name='account' from='accountid' to='parentaccountid' link-type='outer' alias='a_3e31afa8968a4882a6472cee3d5010c5' visible='false'><attribute name='name'/></link-entity></entity></fetch>"; //<fetch version='1.0' output-format='xml-platform' mapping='logical' no-lock='false' distinct='true'><entity name='contact'><attribute name='entityimage_url'/><attribute name='fullname'/><order attribute='fullname' descending='false'/><attribute name='parentcustomerid'/><attribute name='telephone1'/><attribute name='emailaddress1'/><filter type='and'><condition attribute='im360_lineofbusiness' operator='eq' value='100000002'/><condition attribute='statecode' operator='eq' value='0'/></filter><link-entity alias='a_dc9b80f8c78146d89fd6a3b610836975' name='account' from='accountid' to='parentcustomerid' link-type='outer' visible='false'><attribute name='im360_bcn'/></link-entity><attribute name='contactid'/></entity></fetch>";


                    var fetchXmlDoc = XDocument.Parse(xml);

                    //In Unified Interface Fetch Expression may contain filter with isquickfindfields attribute.
                    //This attribute comes from customer search(lookup or quick find).
                    //In this example We use only view filter, if it exists
                    var originalFilter = fetchXmlDoc
                        .Descendants("filter")
                        .Where(f => f.Attribute("isquickfindfields") == null)
                        .FirstOrDefault();

                    //If filter does not exist We may create a new.
                    XElement filterElement;
                    if (originalFilter != null)
                    {
                        filterElement = originalFilter;
                        //Remove filter from main object
                        // filterElement.Remove();

                        List<XNode> xNodes = fetchXmlDoc.DescendantNodes().ToList();
                        foreach (XNode node in xNodes)
                        {
                            XElement element = node as XElement;
                            if (element.Name != "condition") continue;

                            if ((element.FirstAttribute.Value == "im360_lineofbusiness" || element.FirstAttribute.Value == "im360_lineofbusiness2name") && element.LastAttribute.Value.Contains("100000002"))
                            {

                                foreach (XAttribute attribute in element.Attributes())

                                {
                                    if (attribute.Value == "eq")
                                    {

                                    }

                                    if (attribute.Value == "like")
                                    {

                                    }

                                    if (attribute.Value == "link-entity")
                                    {

                                    }

                                }

                            }
                        }
                        //XmlNodeList nodeList = fetchXmlDoc.Nodes();

                        ////Loop through the selected Nodes.
                        //foreach (XmlNode node in filterElement.Nodes)
                        //{
                        //    //Fetch the Node and Attribute values.
                        //    Console.WriteLine("Name: " + node["EmployeeName"].InnerText + " City: " + node.Attributes["City"].Value);
                        //}
                    }
                    


                    //#region
                    //try
                    //{
                    //    //tracingService.Trace("23....");
                    //    var _query = context.InputParameters["Query"];
                    //    if (_query is QueryExpression)
                    //    {
                    //        QueryExpression contactQ = (QueryExpression)context.InputParameters["Query"];
                    //        //tracingService.Trace("26....");

                    //        if (contactQ.EntityName == "contact") // Add your entity logical name
                    //        {
                    //            //tracingService.Trace("27....");
                    //            // //tracingService.Trace(contactQ.);

                    //            if (IsITAD(service) && IsIngram(service))
                    //                return;

                    //            if (!IsITAD(service) && !IsIngram(service))
                    //                return;

                    //            if (IsITAD(service) && !IsIngram(service))
                    //            {
                    //                //tracingService.Trace("32....");

                    //                bool isim360_lineofbusiness = false;
                    //                foreach (var cond in contactQ.Criteria.Conditions)
                    //                {
                    //                    if (cond.AttributeName == "im360_lineofbusiness" && cond.Operator == ConditionOperator.Equal && cond.Values.Contains(100000002))
                    //                    {
                    //                        isim360_lineofbusiness = true;
                    //                        cond.Operator = ConditionOperator.NotEqual;
                    //                    }
                    //                }
                    //                if (!isim360_lineofbusiness)
                    //                {
                    //                    //ConditionExpression condition1 = new ConditionExpression("ownerid", ConditionOperator.Equal, context.UserId);
                    //                    ConditionExpression condition2 = new ConditionExpression("im360_lineofbusiness", ConditionOperator.NotEqual, 100000002);
                    //                    // contactQ.Criteria.AddCondition(condition1);
                    //                    contactQ.Criteria.AddCondition(condition2);


                    //                }

                    //                //tracingService.Trace("40....");
                    //            }

                    //            if (!IsITAD(service) && IsIngram(service))
                    //            {
                    //                //tracingService.Trace("32....");

                    //                bool isim360_lineofbusiness = false;
                    //                foreach (var cond in contactQ.Criteria.Conditions)
                    //                {
                    //                    if (cond.AttributeName == "im360_lineofbusiness" && cond.Operator == ConditionOperator.Equal && cond.Values.Contains(100000002))
                    //                    {
                    //                        isim360_lineofbusiness = true;
                    //                        cond.Operator = ConditionOperator.NotEqual;
                    //                    }
                    //                }
                    //                if (!isim360_lineofbusiness)
                    //                {
                    //                    //ConditionExpression condition1 = new ConditionExpression("ownerid", ConditionOperator.Equal, context.UserId);
                    //                    ConditionExpression condition2 = new ConditionExpression("im360_lineofbusiness", ConditionOperator.NotEqual, 100000002);
                    //                    // contactQ.Criteria.AddCondition(condition1);
                    //                    contactQ.Criteria.AddCondition(condition2);

                    //                }

                    //                //tracingService.Trace("40....");
                    //            }


                    //        }
                    //        //tracingService.Trace("48....");
                    //    }

                    //    if (_query is QueryBase)
                    //    {
                    //        QueryBase contactQueryBase = (QueryBase)context.InputParameters["Query"];
                    //        //tracingService.Trace("51....");
                    //        if (contactQueryBase is FetchExpression)
                    //        {
                    //            //tracingService.Trace("56....");

                    //            FetchExpression fe = contactQueryBase as FetchExpression;
                    //            //tracingService.Trace("58...." + fe.Query.ToString());
                    //            var _fetchquery = fe.Query;
                    //            if (_fetchquery.Contains("filter"))
                    //            {
                    //                if (_fetchquery.Contains("<condition attribute='im360_lineofbusiness'"))
                    //                {


                    //                }
                    //                else
                    //                {

                    //                }

                    //            }


                    //            var fetchExpression = (FetchExpression)context.InputParameters["Query"];

                    //            var fetchXmlDoc = XDocument.Parse(fetchExpression.Query);

                    //            //In Unified Interface Fetch Expression may contain filter with isquickfindfields attribute.
                    //            //This attribute comes from customer search(lookup or quick find).
                    //            //In this example We use only view filter, if it exists
                    //            var originalFilter = fetchXmlDoc
                    //                .Descendants("filter")
                    //                .Where(f => f.Attribute("isquickfindfields") == null)
                    //                .FirstOrDefault();

                    //            //If filter does not exist We may create a new.
                    //            XElement filterElement;
                    //            if (originalFilter != null)
                    //            {
                    //                filterElement = originalFilter;
                    //                //Remove filter from main object
                    //                filterElement.Remove();
                    //            }
                    //            else
                    //            {
                    //                filterElement = new XElement("filter", new XAttribute("type", "and"));
                    //            }

                    //            //All records must contain "Test" word in name field
                    //            var newFilterCondition = new System.Xml.Linq.XElement(
                    //                        "condition",
                    //                        new XAttribute("attribute", "name"),
                    //                        new XAttribute("operator", "like"),
                    //                        new XAttribute("value", "%Test%")
                    //                    );

                    //            //Add new condition to filter
                    //            filterElement.Add(newFilterCondition);
                    //            //tracingService.Trace("Filter result: {0}", filterElement.ToString());

                    //            //Add new filter to main object
                    //            var entityElement = fetchXmlDoc.Descendants("entity").FirstOrDefault();
                    //            entityElement.Add(filterElement);

                    //            var fetchResult = fetchXmlDoc.ToString();
                    //            //tracingService.Trace("Output result: {0}", fetchResult);
                    //            fetchExpression.Query = fetchResult;
                    //        }
                    //    }


                    //}
                    //catch (Exception ex)
                    //{
                    //    //tracingService.Trace("55...." + ex.Message.ToString());

                    //}
                    //#endregion

                    //EntityCollection duplicateList = organizationService.RetrieveMultiple(new FetchExpression(xml));
                    //foreach (var item in duplicateList.Entities)
                    //{
                    //    if (item.RelatedEntities.Count != 0)
                    //    {

                    //    }
                    //}

                    //Console.WriteLine("Press any key to exit.");
                    //Console.ReadLine();
                }

                //string srcConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["CRMConnectionStringSourceDevelopment"].ConnectionString;
                //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                //CrmServiceClient connc = new CrmServiceClient(srcConnectionString);
                //// Cast the proxy client to the IOrganizationService interface.
                //organizationService = (IOrganizationService)connc.OrganizationWebProxyClient != null ? (IOrganizationService)connc.OrganizationWebProxyClient : (IOrganizationService)connc.OrganizationServiceProxy;
                //((OrganizationServiceProxy)(organizationService)).Timeout = new TimeSpan(0, 1200, 0);

                //WhoAmIRequest request = new WhoAmIRequest();
                //WhoAmIResponse response = (WhoAmIResponse)organizationService.Execute(request);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught - " + ex.Message);
            }

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            int RowCount;
            int ColumnCount;
            int TotalRow = 0;
            int TotalColumn = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"C:\1075\contact_dup.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            TotalRow = range.Rows.Count;
            TotalColumn = range.Columns.Count;
            List<contacts> contactlist = new List<contacts>();

            for (RowCount = 1; RowCount <= TotalRow; RowCount++)
            {
                string contactid = string.Empty;
                //  string leadName = string.Empty;
                for (ColumnCount = 1; ColumnCount <= TotalColumn; ColumnCount++)
                {
                    if (ColumnCount == 1)
                    {
                        contactid = (string)(range.Cells[RowCount, ColumnCount] as Excel.Range).Value2;
                    }
                    //else
                    //{
                    //    leadName = (string)(range.Cells[RowCount, ColumnCount] as Excel.Range).Value2;
                    //}
                }
                string xml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='contact'>
                                    <attribute name='fullname' />
                                    <attribute name='contactid' />
                                    <attribute name='im360_lineofbusiness' />
                                    <attribute name='parentcustomerid' />
                                    <attribute name='ownerid' />
                                    <attribute name='createdon' />
                                    <attribute name='emailaddress1' />
                                    <attribute name='im360_createdon' />
                                    <attribute name='createdby' />
                                    <attribute name='statecode' />
                                    <attribute name='im360_createdby' />
                                    <order attribute='fullname' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='contactid' operator='eq' uitype='contact' value='" + contactid + @"' />
                                    </filter>
                                    <link-entity name='account' from='accountid' to='parentcustomerid' visible='false' link-type='outer' alias='accountc'>
                                      <attribute name='im360_lineofbusiness' />
                                    </link-entity>
                                    <link-entity name='contact' from='contactid' to='parentcustomerid' visible='false' link-type='outer' alias='contactc'>
                                      <attribute name='im360_lineofbusiness' />
                                    </link-entity>
                                  </entity>
                                </fetch>";
                EntityCollection duplicateList = organizationService.RetrieveMultiple(new FetchExpression(xml));
                if (duplicateList != null && duplicateList.Entities.Count == 1)
                {
                    var item = duplicateList.Entities[0];
                    contacts cn = new contacts();
                    if (item.Contains("fullname"))
                        cn.fullname = item["fullname"].ToString();

                    if (item.Contains("im360_createdby"))
                        cn.createdby = item["im360_createdby"].ToString();
                    if (item.Contains("emailaddress1"))
                        cn.email = item["emailaddress1"].ToString();

                    if (item.Contains("im360_createdon"))
                        cn.createdon = item["im360_createdon"].ToString();

                    if (item.Contains("createdon"))
                        cn.createdonoob = item["createdon"].ToString();

                    if (item.Contains("createdby"))
                        cn.createdbyoob = ((EntityReference)item["createdby"]).Name.ToString();

                    if (item.Contains("ownerid"))
                        cn.owner = ((EntityReference)item["ownerid"]).Name.ToString();

                    if (item.Contains("parentcustomerid"))
                        cn.company = ((EntityReference)item["parentcustomerid"]).Name;

                    if (item.Contains("im360_lineofbusiness"))
                        cn.lob = item.FormattedValues["im360_lineofbusiness"];


                    if (item.Contains("statecode"))
                        cn.status = item.FormattedValues["statecode"];

                    if (item.Contains("accountc.im360_lineofbusiness"))
                        cn.companylob = item.FormattedValues["accountc.im360_lineofbusiness"];

                    if (item.Contains("contactid"))
                        cn.contactid = item["contactid"].ToString();

                    contactlist.Add(cn);
                    //   myleads += leadName + " - " + leadEmail + " /n " ;
                    // if()
                    //Guid masteraccountId = duplicateList.Entities[0].Id;
                    //foreach (Entity duplicateAccount in duplicateList.Entities)
                    //{
                    //if (masteraccountId != duplicateAccount.Id)
                    //{
                    //    EntityReference target = new EntityReference();

                    //    target.Id = masteraccountId;
                    //    target.LogicalName = "lead";
                    //    Guid subAccountID = duplicateAccount.Id;
                    //    //Create the request. 
                    //    MergeRequest merge = new MergeRequest();

                    //    // SubordinateId is the GUID of the account merging.    
                    //    merge.SubordinateId = subAccountID;
                    //    merge.Target = target;
                    //    merge.PerformParentingChecks = false;
                    //try
                    //{
                    //    MergeResponse mergeRes = (MergeResponse)organizationService.Execute(merge);
                    //    Console.WriteLine("Lead Merged:- " + subAccountID);
                    //}
                    //catch (Exception ex)
                    //{
                    //   // myleads += leadName + " - " + leadEmail + " /n ";
                    //}

                    //   }
                    // }
                }


                Console.Write("\n");
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            //   String csv = String.Join(",", contactlist.Select(x => x.ToString()).ToArray());

            //  File.WriteAllText(@"C:\my code\errors for lead 2.txt", myleads);

            //    SaveToCsv(contactlist, @"C:\my code\dataforcontact.csv");

            // Start Excel and get Application object.
           // ExportToExcel(contactlist);

        }



        public static bool IsITAD(IOrganizationService organizationService)
        {
            string xml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                  <entity name='systemuser'>
                                    <attribute name='fullname' />
                                    <attribute name='systemuserid' />
                                    <order attribute='fullname' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='systemuserid' operator='eq' uiname='Ravina Kadam' uitype='systemuser' value='{8D7A3FEF-D004-EC11-94EF-000D3A5B0F22}' />
                                    </filter>
                                    <link-entity name='teammembership' from='systemuserid' to='systemuserid' visible='false' intersect='true'>
                                      <link-entity name='team' from='teamid' to='teamid' alias='af'>
                                        <filter type='and'>
                                          <condition attribute='name' operator='eq' value='ITAD' />
                                        </filter>
                                      </link-entity>
                                    </link-entity>
                                  </entity>
                                </fetch>";
            EntityCollection systemuserCollection = organizationService.RetrieveMultiple(new FetchExpression(xml));
            if (systemuserCollection.Entities.Count == 0)
                return false;
            else
                return true;
        }

        public static bool IsIngram(IOrganizationService organizationService)
        {
            string xml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                  <entity name='systemuser'>
                                    <attribute name='fullname' />
                                    <attribute name='systemuserid' />
                                    <order attribute='fullname' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='systemuserid' operator='eq' uiname='Ravina Kadam' uitype='systemuser' value='{8D7A3FEF-D004-EC11-94EF-000D3A5B0F22}' />
                                    </filter>
                                    <link-entity name='teammembership' from='systemuserid' to='systemuserid' visible='false' intersect='true'>
                                      <link-entity name='team' from='teamid' to='teamid' alias='af'>
                                        <filter type='and'>
                                          <condition attribute='name' operator='eq' value='CEVA' />
                                        </filter>
                                      </link-entity>
                                    </link-entity>
                                  </entity>
                                </fetch>";
            EntityCollection systemuserCollection = organizationService.RetrieveMultiple(new FetchExpression(xml));
            if (systemuserCollection.Entities.Count == 0)
                return false;
            else
                return true;
        }


        //public static void ExportToExcel(List<contacts> students)

        //{

        //    // Load Excel application

        //    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();



        //    // Create empty workbook

        //    excel.Workbooks.Add();



        //    // Create Worksheet from active sheet

        //    Microsoft.Office.Interop.Excel._Worksheet workSheet = excel.ActiveSheet;

        //    try

        //    {


        //        workSheet.Cells[1, "A"] = "contactid";

        //        workSheet.Cells[1, "B"] = "fullname";

        //        workSheet.Cells[1, "C"] = "Email";


        //        workSheet.Cells[1, "D"] = "lob";

        //        workSheet.Cells[1, "E"] = "company";

        //        workSheet.Cells[1, "F"] = "companylob";

        //        workSheet.Cells[1, "G"] = "createdby";

        //        workSheet.Cells[1, "H"] = "createdbyoob";

        //        workSheet.Cells[1, "I"] = "createdon";

        //        workSheet.Cells[1, "J"] = "createdonoob";
        //        workSheet.Cells[1, "K"] = "owner";
        //        workSheet.Cells[1, "L"] = "Status";




        //        // ————————————————

        //        // Populate sheet with some real data from " Studentts" list

        //        // ————————————————

        //        int row = 2; // start row (in row 1 are header cells)

        //        foreach (contacts student in students)

        //        {

        //            workSheet.Cells[row, "A"] = student.contactid;

        //            workSheet.Cells[row, "B"] = student.fullname;
        //            workSheet.Cells[row, "C"] = student.email;


        //            workSheet.Cells[row, "D"] = student.lob;

        //            workSheet.Cells[row, "E"] = student.company;

        //            workSheet.Cells[row, "F"] = student.companylob;

        //            workSheet.Cells[row, "G"] = student.createdby;

        //            workSheet.Cells[row, "H"] = student.createdbyoob;

        //            workSheet.Cells[row, "I"] = student.createdon;

        //            workSheet.Cells[row, "J"] = student.createdonoob;
        //            workSheet.Cells[row, "K"] = student.owner;
        //            workSheet.Cells[row, "L"] = student.status;




        //            row++;

        //        }



        //        // Apply some predefined styles for data to look nicely 🙂

        //        workSheet.Range["A1"].AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);



        //        // Define filename

        //        string fileName = string.Format(@"C:\my code\ExcelData.xlsx");



        //        // Save this data as a file

        //        workSheet.SaveAs(fileName);



        //        // Display SUCCESS message

        //    }

        //    catch (Exception exception)

        //    {



        //    }

        //    finally

        //    {

        //        // Quit Excel application

        //        excel.Quit();



        //        // Release COM objects (very important!)

        //        if (excel != null)

        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);



        //        if (workSheet != null)

        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);



        //        // Empty variables

        //        excel = null;

        //        workSheet = null;



        //        // Force garbage collector cleaning

        //        GC.Collect();

        //    }

        //}
        ////public static void SaveToCsv<T>(List<T> reportData, string path)
        ////{
        ////    var lines = new List<string>();
        ////    IEnumerable<PropertyDescriptor> props = TypeDescriptor.GetProperties(typeof(T)).OfType<PropertyDescriptor>();
        ////    var header = string.Join(",", props.ToList().Select(x => x.Name));
        ////    lines.Add(header);
        ////    var valueLines = reportData.Select(row => string.Join(",", header.Split(',').Select(a => row.GetType().GetProperty(a).GetValue(row, null))));
        ////    lines.AddRange(valueLines);
        ////    File.WriteAllLines(path, lines.ToArray());
        ////}





    }
}


public class contacts
{

    public string contactid { get; set; }
    public string fullname { get; set; }
    public string email { get; set; }

    public string lob { get; set; }
    public string company { get; set; }



    public string companylob { get; set; }

    public string createdon { get; set; }

    public string createdonoob { get; set; }

    public string createdby { get; set; }
    public string createdbyoob { get; set; }
    public string owner { get; set; }

    public string status { get; set; }



}

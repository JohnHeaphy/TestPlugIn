namespace Console_SetAges
{
    using System;
    using System.ServiceModel.Description;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Xrm.Sdk.Client;
    using Microsoft.Crm.Sdk.Messages;


    class VolunteerConsole
    {
        static void Main(String[] CRMOrgName)
        {

            // Set Today
            DateTime Now = DateTime.Now;
            string Today = Now.ToString("yyyy-MM-dd");

            Console.WriteLine("Hello!");
            Console.WriteLine(Today);

            //Console.ReadLine();

            String CRMOrg = "crm16uat";
            if (CRMOrgName.Length > 0)
            {
                CRMOrg = CRMOrgName[0];
            }

            Uri system_uri = null;

            ClientCredentials clientcred = new ClientCredentials();
            clientcred.UserName.UserName = "TSA\\ConsoleServices";
            clientcred.UserName.Password = "";

            switch (CRMOrg)
            {

                case "CRM16Build":
                    system_uri = new Uri("https://crm16build.stroke.org.uk/XRMServices/2011/Organization.svc");
                    break;

                case "CRM16UAT":
                    system_uri = new Uri("https://crm16uat.stroke.org.uk/XRMServices/2011/Organization.svc");
                    break;

                case "CRM16Training":
                    system_uri = new Uri("https://crm16training.stroke.org.uk/XRMServices/2011/Organization.svc");
                    break;

                case "CRM16":
                    system_uri = new Uri("https://crm16.stroke.org.uk/XRMServices/2011/Organization.svc");
                    break;

                default:
                    system_uri = new Uri("https://crm16build.stroke.org.uk/XRMServices/2011/Organization.svc");
                    CRMOrg = "CRM16";
                    break;

            }

            Console.WriteLine(system_uri);

            OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(system_uri, null, clientcred, null);
            IOrganizationService service = (IOrganizationService)serviceProxy;


            //Active Date

            string ActiveDate = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='ccx_volunteer'>
            <attribute name='ccx_name' />
            <attribute name='ccx_volunteerid' />
            <attribute name='ccx_calculatedtrialenddate' />
            <order attribute='ccx_name' descending='false' />
            <filter type='and'>
            <condition attribute='statecode' operator='eq' value='0' />
            <condition attribute='ccx_activedate' operator='null' />
            <condition attribute='ccx_mandatorytrainingstatus' operator='ne' value='2' />
            <condition attribute='ccx_calculatedtrialenddate' operator='on-or-before' value='";

            ActiveDate += Today;

            ActiveDate += @"' />

            </filter>
            </entity>
            </fetch>";


            FetchExpression fetch_ActiveDate = new FetchExpression(ActiveDate);
            EntityCollection results_ActiveDate = service.RetrieveMultiple(fetch_ActiveDate);

            if (results_ActiveDate.Entities.Count > 0)

                foreach (var ActiveDateRecord in results_ActiveDate.Entities)
                {
                    DateTime TrialEndDate = new DateTime();
                    TrialEndDate = (DateTime)ActiveDateRecord.Attributes["ccx_calculatedtrialenddate"];


                    // Update Volunteer Record
                    Entity VolunteerUpdate = new Entity("ccx_volunteer");
                    VolunteerUpdate.Id = ActiveDateRecord.Id;
                    VolunteerUpdate.Attributes["ccx_activedate"] = TrialEndDate;
                    service.Update(VolunteerUpdate);
                    // End Update Volunteer

                }

            // Run activity expiry process
            string ActivityExpiry = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='ccx_volunteer_activity'>
            <attribute name='createdon' />
            <attribute name='ccx_volunteerroleid' />
            <attribute name='createdby' />
            <attribute name='ccx_activity_date' />
            <attribute name='ccx_activitycodelookup' />
            <attribute name='ccx_duration' />
            <attribute name='ccx_volunteer_id' />
            <attribute name='ccx_mandatorytrainingstatus' />
            <attribute name='ccx_volunteer_activityid' />
            <order attribute='ccx_activity_date' descending='true' />
            <filter type='and'>
            <condition attribute='statecode' operator='eq' value='0' />
            <condition attribute='ccx_activitystatus' operator='eq' value='3' />
            <condition attribute='ccx_expiryprocessrun' operator='ne' value='1' />
            </filter>
            </entity>
            </fetch>";


            FetchExpression fetch_ActivityExpiry = new FetchExpression(ActivityExpiry);
            EntityCollection results_ActivityExpiry = service.RetrieveMultiple(fetch_ActivityExpiry);

            if (results_ActivityExpiry.Entities.Count > 0)

                foreach (var ActivityAlertExpiryEmailRecord in results_ActivityExpiry.Entities)
                {
                    //// Create an ExecuteWorkflow request.
                    ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
                    {
                        WorkflowId = new Guid("D41AACD1-9736-428E-AD74-3ED0CB84F60E"),
                        EntityId = ActivityAlertExpiryEmailRecord.Id
                    };

                    service.Execute(request);
                }

            //Activity Alert expiry Email
            string ActivityAlertExpiryEmail = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='ccx_volunteer_activity'>
            <attribute name='ccx_volunteer_activityid' />
            <attribute name='ccx_name' />
            <order attribute='ccx_name' descending='false' />
            <filter type='and'>
            <condition attribute='statecode' operator='eq' value='0' />
            <condition attribute='ccx_alertemailsent' operator='eq' value='0' />

            <condition attribute='ccx_alertemaildate' operator='on-or-before' value='";

            ActivityAlertExpiryEmail += Today;
            ActivityAlertExpiryEmail += @"' />

            </filter>
            <link-entity name='ccx_volunteer' from='ccx_volunteerid' to='ccx_volunteer_id' alias='ae'>
            <filter type='and'>
            <condition attribute='ccx_volunteerstatus' operator='in'>
            <value>5</value>
            <value>2</value>
            <value>3</value>
            <value>6</value>
            <value>4</value>
            </condition>
            </filter>
            </link-entity>
            </entity>
            </fetch>";


            FetchExpression fetch_ActivityAlertExpiryEmail = new FetchExpression(ActivityAlertExpiryEmail);
            EntityCollection results_ActivityAlertExpiryEmail = service.RetrieveMultiple(fetch_ActivityAlertExpiryEmail);

            if (results_ActivityAlertExpiryEmail.Entities.Count > 0)

                foreach (var ActivityAlertExpiryEmailRecord in results_ActivityAlertExpiryEmail.Entities)
                {
                    //// Create an ExecuteWorkflow request.
                    ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
                    {
                        WorkflowId = new Guid("63ad8780-d649-4ce8-89c1-03c7b589ce44"),
                        EntityId = ActivityAlertExpiryEmailRecord.Id
                    };

                    service.Execute(request);
                }

            string CheckReminderEmail = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='ccx_volunteer'>
            <attribute name='ccx_name' />
            <attribute name='ccx_volunteerid' />
            <order attribute='ccx_name' descending='false' />
            <filter type='and'>
            <condition attribute='statecode' operator='eq' value='0' />
            <condition attribute='ccx_checkexpirydate' operator='next-x-months' value='3' />
            <condition attribute='ccx_check_3month_reminder_email_sent' operator='eq' value='0' />
            <condition attribute='ccx_volunteerstatus' operator='in'>
            <value>5</value>
            <value>2</value>
            <value>3</value>
            <value>6</value>
            <value>4</value>
            </condition>
            </filter>
            </entity>
            </fetch>";


            FetchExpression fetch_CheckReminderEmail = new FetchExpression(CheckReminderEmail);
            EntityCollection results_CheckReminderEmail = service.RetrieveMultiple(fetch_CheckReminderEmail);

            if (results_CheckReminderEmail.Entities.Count > 0)

                foreach (var CheckReminderEmailRecord in results_CheckReminderEmail.Entities)
                {
                    //// Create an ExecuteWorkflow request.
                    ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
                    {
                        WorkflowId = new Guid("14031812-4FFE-4DD9-8360-6E126B5D11A3"),
                        EntityId = CheckReminderEmailRecord.Id
                    };

                    service.Execute(request);
                }


            string CheckExpiredEmail = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='ccx_volunteer'>
            <attribute name='ccx_name' />
            <attribute name='ccx_volunteerid' />
            <order attribute='ccx_name' descending='false' />
            <filter type='and'>
            <condition attribute='statecode' operator='eq' value='0' />
            <condition attribute='ccx_check_received_date' operator='not-null' />
            <condition attribute='ccx_check_expiry_email_workflow' operator='eq' value='0' />
            <condition attribute='ccx_volunteerstatus' operator='in'>
            <value>5</value>
            <value>2</value>
            <value>3</value>
            <value>6</value>
            <value>4</value>
            </condition>
            <condition attribute='ccx_checkexpirydate' operator='on-or-before' value='";
            CheckExpiredEmail += Today;
            CheckExpiredEmail += @"' />
            </filter>
            </entity>
            </fetch>";


            FetchExpression fetch_CheckExpiredEmail = new FetchExpression(CheckExpiredEmail);
            EntityCollection results_CheckExpiredEmail = service.RetrieveMultiple(fetch_CheckExpiredEmail);

            if (results_CheckExpiredEmail.Entities.Count > 0)

                foreach (var CheckExpiredEmailRecord in results_CheckExpiredEmail.Entities)
                {
                    //// Create an ExecuteWorkflow request.
                    ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
                    {
                        WorkflowId = new Guid("EEA42407-DE60-4336-BDBD-FDF86CDFCDC6"),
                        EntityId = CheckExpiredEmailRecord.Id
                    };

                    service.Execute(request);
                }

            //Volunteer Close Email
            string VolunteerCloseEmail = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='ccx_volunteer'>
            <attribute name='ccx_name' />
            <attribute name='ccx_volunteerstatus' />
            <attribute name='ccx_volunteerid' />
            <order attribute='ccx_name' descending='false' />
            <filter type='and'>
            <condition attribute='statecode' operator='eq' value='0' />
            <condition attribute='ccx_volunteerstatus' operator='in'>
            <value>9</value>
            <value>7</value>
            <value>8</value>
            </condition>
            <condition attribute='ccx_closeemailsent' operator='ne' value='1' />
            </filter>
            </entity>
            </fetch>";


            FetchExpression fetch_VolunteerCloseEmail = new FetchExpression(VolunteerCloseEmail);
            EntityCollection results_VolunteerCloseEmail = service.RetrieveMultiple(fetch_VolunteerCloseEmail);

            if (results_VolunteerCloseEmail.Entities.Count > 0)

                foreach (var VolunteerCloseEmailRecord in results_VolunteerCloseEmail.Entities)
                {
                    //// Create an ExecuteWorkflow request.
                    ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
                    {
                        WorkflowId = new Guid("728816CD-C601-4BCD-ADC7-51148528EA38"),
                        EntityId = VolunteerCloseEmailRecord.Id
                    };

                    service.Execute(request);
                }

            //Volunteer DCA Active
            string VolunteerDCAActive = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='ccx_volunteer'>
            <attribute name='ccx_name' />
            <attribute name='ccx_volunteerid' />
            <order attribute='ccx_name' descending='false' />
            <filter type='and'>
            <condition attribute='statecode' operator='eq' value='0' />
            <condition attribute='ccx_volunteerstatus' operator='in'>
            <value>4</value>
            <value>5</value>
            <value>6</value>
            </condition>
            </filter>
            <link-entity name='contact' from='contactid' to='ccx_contactid' alias='aa'>
            <filter type='and'>
            <condition attribute='ccx_dca_vol_active_volunteer' operator='eq' value='0' />
            </filter>
            </link-entity>
            </entity>
            </fetch>";


            FetchExpression fetch_VolunteerDCAActive = new FetchExpression(VolunteerDCAActive);
            EntityCollection results_VolunteerDCAActive = service.RetrieveMultiple(fetch_VolunteerDCAActive);

            if (results_VolunteerDCAActive.Entities.Count > 0)

                foreach (var VolunteerDCAActiveRecord in results_VolunteerDCAActive.Entities)
                {
                    //// Create an ExecuteWorkflow request.
                    ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
                    {
                        WorkflowId = new Guid("BE7298C5-CF1C-480A-ACB6-E64DE362B52F"),
                        EntityId = VolunteerDCAActiveRecord.Id
                    };

                    service.Execute(request);
                }


            //Volunteer DCA Close
            string VolunteerDCAClose = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='ccx_volunteer'>
            <attribute name='ccx_name' />
            <attribute name='ccx_volunteerid' />
            <order attribute='ccx_name' descending='false' />
            <filter type='and'>
            <condition attribute='statecode' operator='eq' value='0' />
            <condition attribute='ccx_volunteerstatus' operator='in'>
            <value>7</value>
            </condition>
            </filter>
            <link-entity name='contact' from='contactid' to='ccx_contactid' alias='ad'>
            <filter type='and'>
            <condition attribute='ccx_dca_vol_active_volunteer' operator='eq' value='1' />
            </filter>
            </link-entity>
            </entity>
            </fetch>";


            FetchExpression fetch_VolunteerDCAClose = new FetchExpression(VolunteerDCAClose);
            EntityCollection results_VolunteerDCAClose = service.RetrieveMultiple(fetch_VolunteerDCAClose);

            if (results_VolunteerDCAClose.Entities.Count > 0)

                foreach (var VolunteerDCACloseRecord in results_VolunteerDCAClose.Entities)
                {
                    //// Create an ExecuteWorkflow request.
                    ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
                    {
                        WorkflowId = new Guid("BE7298C5-CF1C-480A-ACB6-E64DE362B52F"),
                        EntityId = VolunteerDCACloseRecord.Id
                    };

                    service.Execute(request);
                }


            // Needs Update
            string IDReminder = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='ccx_volunteer'>
            <attribute name='ccx_name' />
            <attribute name='ccx_volunteerid' />
            <order attribute='ccx_name' descending='false' />
            <filter type='and'>
            <condition attribute='statecode' operator='eq' value='0' />
            <condition attribute='ccx_calculatedidcardexpirydate' operator='next-x-months' value='2' />
            <condition attribute='ccx_id_badge_reminder_workflow' operator='eq' value='0' />
            <condition attribute='ccx_volunteerstatus' operator='in'>
            <value>5</value>
            <value>2</value>
            <value>3</value>
            <value>6</value>
            <value>4</value>
            </condition>
            </filter>
            </entity>
            </fetch>";


            FetchExpression fetch_IDReminder = new FetchExpression(IDReminder);
            EntityCollection results_IDReminder = service.RetrieveMultiple(fetch_IDReminder);

            if (results_IDReminder.Entities.Count > 0)

                foreach (var IDReminderRecord in results_IDReminder.Entities)
                {
                    //// Create an ExecuteWorkflow request.
                    ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
                    {
                        WorkflowId = new Guid("7DBA95D1-8985-E211-BA17-EA2A7554FC37"),
                        EntityId = IDReminderRecord.Id
                    };

                    service.Execute(request);

                }

            string IDExpired = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='ccx_volunteer'>
            <attribute name='ccx_name' />
            <attribute name='ccx_volunteerid' />
            <order attribute='ccx_name' descending='false' />
            <filter type='and'>
            <condition attribute='statecode' operator='eq' value='0' />
            <condition attribute='ccx_id_badge_expired_workflow' operator='eq' value='0' />
            <condition attribute='ccx_volunteerstatus' operator='in'>
            <value>5</value>
            <value>2</value>
            <value>3</value>
            <value>6</value>
            <value>4</value>
            </condition>
            <condition attribute='ccx_calculatedidcardexpirydate' operator='on-or-before' value='";

            IDExpired += Today;

            IDExpired += @"' />

            </filter>
            </entity>
            </fetch>";


            FetchExpression fetch_IDExpired = new FetchExpression(IDExpired);
            EntityCollection results_IDExpired = service.RetrieveMultiple(fetch_IDExpired);

            if (results_IDExpired.Entities.Count > 0)

                foreach (var IDExpiredRecord in results_IDExpired.Entities)
                {
                    //// Create an ExecuteWorkflow request.
                    ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
                    {
                        WorkflowId = new Guid("3528C52D-8985-E211-BA17-EA2A7554FC37"),
                        EntityId = IDExpiredRecord.Id
                    };

                    service.Execute(request);
                }

            // Awards 100 200 500 hours

            string Award100 = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='ccx_volunteer'>
            <attribute name='ccx_name' />
            <order attribute='ccx_name' descending='false' />
            <filter type='and'>
            <condition attribute='statecode' operator='eq' value='0' />
            <condition attribute='ccx_volunteerstatus' operator='in'>
            <value>4</value>
            <value>5</value>
            <value>6</value>
            </condition>
            <condition attribute='ccx_totalduration' operator='gt' value='5999' />
            <condition attribute='ccx_award_count' operator='lt' value='100' />
            </filter>
            </entity>
            </fetch>";


            FetchExpression fetch_Award100 = new FetchExpression(Award100);
            EntityCollection results_Award100 = service.RetrieveMultiple(fetch_Award100);

            if (results_Award100.Entities.Count > 0)

                foreach (var Award100Record in results_Award100.Entities)
                {
                    // Update Volunteer Record
                    Entity Award100Update = new Entity("ccx_volunteer");
                    Award100Update.Id = Award100Record.Id;
                    Award100Update.Attributes["ccx_award_count"] = 100;
                    service.Update(Award100Update);
                    // End Update Volunteer

                }

            string Award200 = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='ccx_volunteer'>
            <attribute name='ccx_name' />
            <order attribute='ccx_name' descending='false' />
            <filter type='and'>
            <condition attribute='statecode' operator='eq' value='0' />
            <condition attribute='ccx_volunteerstatus' operator='in'>
            <value>4</value>
            <value>5</value>
            <value>6</value>
            </condition>
            <condition attribute='ccx_totalduration' operator='gt' value='11999' />
            <condition attribute='ccx_award_count' operator='lt' value='200' />
            </filter>
            </entity>
            </fetch>";


            FetchExpression fetch_Award200 = new FetchExpression(Award200);
            EntityCollection results_Award200 = service.RetrieveMultiple(fetch_Award200);

            if (results_Award200.Entities.Count > 0)

                foreach (var Award200Record in results_Award200.Entities)
                {
                    // Update Volunteer Record
                    Entity Award200Update = new Entity("ccx_volunteer");
                    Award200Update.Id = Award200Record.Id;
                    Award200Update.Attributes["ccx_award_count"] = 200;
                    service.Update(Award200Update);
                    // End Update Volunteer

                }




            string Award500 = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='ccx_volunteer'>
            <attribute name='ccx_name' />
            <order attribute='ccx_name' descending='false' />
            <filter type='and'>
            <condition attribute='statecode' operator='eq' value='0' />
            <condition attribute='ccx_volunteerstatus' operator='in'>
            <value>4</value>
            <value>5</value>
            <value>6</value>
            </condition>
            <condition attribute='ccx_totalduration' operator='gt' value='29999' />
            <condition attribute='ccx_award_count' operator='lt' value='500' />
            </filter>
            </entity>
            </fetch>";


            FetchExpression fetch_Award500 = new FetchExpression(Award500);
            EntityCollection results_Award500 = service.RetrieveMultiple(fetch_Award500);

            if (results_Award500.Entities.Count > 0)

                foreach (var Award500Record in results_Award500.Entities)
                {
                    // Update Volunteer Record
                    Entity Award500Update = new Entity("ccx_volunteer");
                    Award500Update.Id = Award500Record.Id;
                    Award500Update.Attributes["ccx_award_count"] = 500;
                    service.Update(Award500Update);
                    // End Update Volunteer

                }




            //// need Update
            string DrivingDatesExpired = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='ccx_volunteer'>
            <attribute name='ccx_name' />
            <attribute name='ccx_volunteerid' />
            <order attribute='ccx_name' descending='false' />
            <filter type='and'>
            <condition attribute='statecode' operator='eq' value='0' />
            <condition attribute='ccx_volunteerstatus' operator='in'>
            <value>5</value>
            <value>2</value>
            <value>3</value>
            <value>6</value>
            <value>4</value>
            </condition>
            <condition attribute='ccx_drivingstatus' operator='eq' value='3' />
            <condition attribute='ccx_driving_status_email_sent' operator='ne' value='1' />
            </filter>
            </entity>
            </fetch>";


            FetchExpression fetch_DrivingDatesExpired = new FetchExpression(DrivingDatesExpired);
            EntityCollection results_DrivingDatesExpired = service.RetrieveMultiple(fetch_DrivingDatesExpired);

            if (results_DrivingDatesExpired.Entities.Count > 0)

                foreach (var DrivingDatesExpiredRecord in results_DrivingDatesExpired.Entities)
                {
                    //// Create an ExecuteWorkflow request.
                    ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
                    {
                        WorkflowId = new Guid("4F0141D8-FF63-4D20-89EB-46482755B189"),
                        EntityId = DrivingDatesExpiredRecord.Id
                    };

                    service.Execute(request);
                }

        }
    }

}

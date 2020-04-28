using System;
using System.Collections.Generic;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Sample.UCI;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Lookup = Microsoft.Dynamics365.UIAutomation.Api.UCI.Lookup;
using LookupItem = Microsoft.Dynamics365.UIAutomation.Api.UCI.LookupItem;
using OptionSet = Microsoft.Dynamics365.UIAutomation.Api.UCI.OptionSet;

namespace Microsoft.Dynamics365.UIAutomation.Sample.MSD365.Person
{
    [TestClass]
    public class CreatePerson1 : TestsBase
    {

        [TestInitialize]

        public override void InitTest() => base.InitTest();

        [TestCleanup]
        public override void FinishTest() => base.FinishTest();

        public override void NavigateToHomePage() => NavigateTo(UCIAppName.CRM, "CRM", "Persons");




        [TestMethod]
        public void CreateNewPer()
        
        {

            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            //using (var xrmBrowser = new Api.Browser(TestSettings.Options))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
                

                xrmApp.Navigation.OpenApp(UCIAppName.CRM);
                xrmApp.ThinkTime(500);
                xrmApp.Navigation.OpenSubArea("CRM", "Persons");

                //xrmBrowser.ThinkTime(1000);
                //xrmBrowser.Grid.SwitchView("Active Contacts");

                xrmApp.ThinkTime(2000);
                xrmApp.CommandBar.ClickCommand("New");

                xrmApp.ThinkTime(5000);

                //var fields = new List<Field>
                //{
                //    new Field() {Id = "firstname", Value = "TestNarAuto1"},
                //    new Field() {Id = "lastname", Value = "sahoo"}
                //};
                //xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "fullname", Fields = fields });



                //xrmApp.Entity.SetValue("firstname", TestSettings.GetRandomString(5, 10));
                //xrmApp.Entity.SetValue("lastname", TestSettings.GetRandomString(5, 10));



                //_xrmApp.Entity.SetValue();

                //_client.
                //.SetValue("Per_FN", "abc", "");

                xrmApp.Entity.SetValue("firstname.fieldControl-text-box-text", "TestNarAuto1");
                xrmApp.Entity.SetValue("Last Name", "sahoo");
                //xrmApp.Entity.SetValue("Language", "English");
                //xrmApp.Entity.SetValue("Preferred Phone", "555-555-5555");
                //xrmApp.Entity.SetValue(new DateTimeControl { Name = "Date of Birth", Value = DateTime.Parse("11/1/1980") });
                //xrmBrowser.Entity.SetValue(new OptionSet { Name = "preferredcontactmethodcode", Value = "Email" });

                xrmApp.CommandBar.ClickCommand("Save");

            }

        }

        [TestMethod]
        public void CreateNewPer2()
        {

            _xrmApp.CommandBar.ClickCommand("New");
            _xrmApp.Entity.SetValue("firstname", "TestNarAuto1");


            _xrmApp.Entity.SetValue("lastname", "sahoo1");

            _xrmApp.Entity.SetValue(new OptionSet { Name = "gendercode", Value = "Male" });

            _xrmApp.Entity.SetValue(new LookupItem { Name = "aab_languageid", Value = "Englisch" });
        }

    }
}

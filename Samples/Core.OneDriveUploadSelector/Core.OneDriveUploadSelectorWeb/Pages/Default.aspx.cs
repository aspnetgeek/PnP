using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.SharePoint.Client;
using System.Xml.Linq;

namespace Core.OneDriveUploadSelectorWeb {
    public partial class Default : System.Web.UI.Page {
        XNamespace ns = "http://schemas.microsoft.com/sharepoint/";

        protected void Page_PreInit(object sender, EventArgs e) {
            Uri redirectUrl;
            switch (SharePointContextProvider.CheckRedirectionStatus(Context, out redirectUrl)) {
                case RedirectionStatus.Ok:
                    return;
                case RedirectionStatus.ShouldRedirect:
                    Response.Redirect(redirectUrl.AbsoluteUri, endResponse: true);
                    break;
                case RedirectionStatus.CanNotRedirect:
                    Response.Write("An error occurred while processing your request.");
                    Response.End();
                    break;
            }
        }

        protected void Page_Load(object sender, EventArgs e) {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);
            LoadChrome();
            DocumentsLink.NavigateUrl = spContext.SPHostUrl + "Shared Documents";
        }

        XElement GetCustomActionXmlNode() {
            var filePath = Server.MapPath("~/Models/RibbonCommands.xml");
            var xdoc = XDocument.Load(filePath);
            var customActionNode = xdoc.Element(ns + "Elements").Element(ns + "CustomAction");
            return customActionNode;
        }

        string GetAttributeValue(XElement node, string attributeName, string defaultValue = "") {
            if (node.Attribute(attributeName) == null)
                return defaultValue;

            return node.Attribute(attributeName).Value;
        }

        protected void InitializeButton_Click(object sender, EventArgs e) {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

            using (var clientContext = spContext.CreateUserClientContextForSPHost()) {
                var customActionNode = GetCustomActionXmlNode();
                var customActionName = customActionNode.Attribute("Id").Value;
                var commandUIExtensionNode = customActionNode.Element(ns + "CommandUIExtension");
                var xmlContent = commandUIExtensionNode.ToString();

                // get attribute values
                var location = customActionNode.Attribute("Location").Value; // allow error to be thrown if null
                var registrationId = GetAttributeValue(customActionNode, "RegistrationId");
                var registrationTypeString = GetAttributeValue(customActionNode, "RegistrationType", UserCustomActionRegistrationType.None.ToString());
                var registrationType = (UserCustomActionRegistrationType)Enum.Parse(typeof(UserCustomActionRegistrationType), registrationTypeString);
                var sequenceString = GetAttributeValue(customActionNode, "Sequence");
                var title = GetAttributeValue(customActionNode, "Title", customActionName);

                int? sequence = null;
                if (!string.IsNullOrEmpty(sequenceString))
                    sequence = Convert.ToInt32(sequenceString);

                // see of the custom action already exists
                clientContext.Load(clientContext.Web, web => web.UserCustomActions);
                clientContext.ExecuteQuery();
                var customAction = clientContext.Web.UserCustomActions.FirstOrDefault(uca => uca.Name == customActionName);

                // if it does not exist, create it
                if (customAction == null) {
                    // create the ribbon
                    customAction = clientContext.Web.UserCustomActions.Add();
                    customAction.Name = customActionName;
                }

                // set custom action properties
                customAction.Location = location;
                customAction.CommandUIExtension = xmlContent; // CommandUIExtension xml
                customAction.RegistrationId = registrationId;
                customAction.RegistrationType = registrationType;
                customAction.Title = title;

                if (sequence.HasValue)
                    customAction.Sequence = sequence.Value;

                customAction.Update();
                clientContext.Load(customAction);
                clientContext.ExecuteQuery();
            }
        }

        protected void RemoveButton_Click(object sender, EventArgs e) {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

            using (var clientContext = spContext.CreateUserClientContextForSPHost()) {
                var customActionNode = GetCustomActionXmlNode();
                var customActionName = customActionNode.Attribute("Id").Value;

                clientContext.Load(clientContext.Web, web => web.UserCustomActions);
                clientContext.ExecuteQuery();

                var customAction = clientContext.Web.UserCustomActions.FirstOrDefault(uca => uca.Name == customActionName);

                if (customAction != null) {
                    customAction.DeleteObject();
                    clientContext.ExecuteQuery();
                }
            }
        }

        void LoadChrome() {

            // define initial script, needed to render the chrome control
            string script = @"
            function chromeLoaded() {
                $('body').show();
            }

            //function callback to render chrome after SP.UI.Controls.js loads
            function renderSPChrome() {
                //Set the chrome options for launching Help, Account, and Contact pages
                var options = {
                    'appTitle': document.title,
                    'onCssLoaded': 'chromeLoaded()'
                };

                //Load the Chrome Control in the divSPChrome element of the page
                var chromeNavigation = new SP.UI.Controls.Navigation('divSPChrome', options);
                chromeNavigation.setVisible(true);
            }";

            //register script in page
            Page.ClientScript.RegisterClientScriptBlock(typeof(Default), "BasePageScript", script, true);
        }
    }
}
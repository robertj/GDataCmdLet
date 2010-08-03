using System;
using System.Diagnostics;
using System.Management.Automation;            
using System.ComponentModel;
using Google.Contacts;
using Google.GData.Client;
using Google.GData.Contacts;
using Google.GData.Extensions;
using System.Collections.Generic;


namespace Microsoft.PowerShell.GData
{

    public class Contact
    {

        #region Remove-GDataContact

        [Cmdlet(VerbsCommon.Remove, "GDataContact")]
        public class RemoveGDataContact : Cmdlet
        {

            [Parameter(
            Mandatory = true,
            HelpMessage = "ContactService, new-GdataContactService"
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Contact Uri"
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { selfUri = value; }
            }
            private string selfUri;

            private Dgc.GoogleContactsService dgcGoogleContactsService = new Dgc.GoogleContactsService();

            protected override void ProcessRecord()
            {
                var _domain = dgcGoogleContactsService.GetDomain(service.ContactsService);
                var _query = new ContactsQuery(ContactsQuery.CreateContactsUri(_domain));
                var _feed = service.ContactsService.Query(_query);
                foreach (var _entry in _feed.Entries)
                {
                    if (_entry.SelfUri.Content == selfUri)
                    {
                        try
                        {
                            service.ContactsService.Delete(_entry);
                            WriteObject(_entry, true);
                        }
                        catch (Exception _Exception)
                        {
                            WriteObject(_Exception);
                        }
                     }
                    else
                    {
                        throw new Exception("SelfUri not found!");
                    }

                }

            }

        }

        #endregion Remove-GDataContact

        #region Get-GDataContact

        [Cmdlet(VerbsCommon.Get, "GDataContact")]
        public class GetGDataContact : Cmdlet
        {
            #region Parameters

            [Parameter(
            Mandatory = true,
            HelpMessage = "ContactService, new-GdataContactService"
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;
 
            [Parameter(
            Mandatory = false,
            HelpMessage = "Contact Uri"
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { selfUri = value; }
            }
            private string selfUri;
            
            [Parameter(
            Mandatory = false,
            HelpMessage = "Contact Name"
            )]
            [ValidateNotNullOrEmpty]
            public string Name
            {
                get { return null; }
                set { name = value; }
            }
            private string name;
            
            #endregion Parameters

            private Dgc.GoogleContactsService dgcGoogleContactsService = new Dgc.GoogleContactsService(); 
            protected override void ProcessRecord()
            {
                var _domain = dgcGoogleContactsService.GetDomain(service.ContactsService);
                try
                {
                    if (selfUri != null)
                    {
                        var _query = new ContactsQuery(ContactsQuery.CreateContactsUri(_domain));
                        var _feed = service.ContactsService.Query(_query);
                        
                        foreach (ContactEntry _entry in _feed.Entries)
                        {
                            if (_entry.SelfUri.Content == selfUri)
                            {
                                var _contactEntry = dgcGoogleContactsService.CreateContactEntry(_entry);
                                WriteObject(_contactEntry);
                            }
                        }
                    }
                    else if (name != null)
                    {

                        var _query = new ContactsQuery(ContactsQuery.CreateContactsUri(_domain));
                        var _feed = service.OauthContactsService.Query(_query);
                        var _contactsEntrys = new GDataTypes.GDataContactEntrys();
                        foreach (ContactEntry _entry in _feed.Entries)
                        {
                            if (_entry.Title.Text != null)
                            {
                                if (System.Text.RegularExpressions.Regex.IsMatch(_entry.Title.Text, name))
                                {
                                    var _contactEntry = dgcGoogleContactsService.CreateContactEntry(_entry);
                                    _contactsEntrys = dgcGoogleContactsService.AppendContactEntrys(_entry, _contactsEntrys);
                                }
                            }

                        }
                        WriteObject(_contactsEntrys, true);
                    }
                    else
                    {  
                        var _query = new ContactsQuery(ContactsQuery.CreateContactsUri(_domain));
                        var _feed = service.ContactsService.Query(_query);
                        var _contactEntrys = dgcGoogleContactsService.CreateContactEntrys(_feed);
                        WriteObject(_contactEntrys, true);
                    }
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception); 

                }
            }
        }

        #endregion Get-GDataContact

        #region Set-GDataContact

        [Cmdlet(VerbsCommon.Set, "GDataContact")]
        public class SetGDataContact : Cmdlet
        {
            
            #region Parameters

            [Parameter(
            Mandatory = true,
            HelpMessage = "ContactService, new-GdataContactService"
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Contact Uri"
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { selfUri = value; }
            }
            private string selfUri;

            [Parameter(
             Mandatory = false,
             HelpMessage = "Contact EmailAdress"
            )]
            [ValidateNotNullOrEmpty]
            public string EmailAddress
            {
                get { return null; }
                set { emailAddress = value; }
            }
            private string emailAddress;

            [Parameter(
               Mandatory = false,
               HelpMessage = "Contact Name"
            )]
            [ValidateNotNullOrEmpty]
            public string Name
            {
                get { return null; }
                set { name = value; }
            }
            private string name;


            [Parameter(
               Mandatory = false,
               HelpMessage = "Contact PhoneNumber"
            )]
            [ValidateNotNullOrEmpty]
            public string PhoneNumber
            {
                get { return null; }
                set { phoneNumber = value; }
            }
            private string phoneNumber;


            [Parameter(
               Mandatory = false,
               HelpMessage = "Contact PostalAddress"
            )]
            [ValidateNotNullOrEmpty]
            public string PostalAddress
            {
                get { return null; }
                set { postalAddress = value; }
            }
            private string postalAddress;


            [Parameter(
               Mandatory = false,
               HelpMessage = "Contact City"
            )]
            [ValidateNotNullOrEmpty]
            public string City
            {
                get { return null; }
                set { city = value; }
            }
            private string city;
            
            #endregion Parameters

            private Dgc.GoogleContactsService dgcGoogleContactsService = new Dgc.GoogleContactsService();
            protected override void ProcessRecord()
            {
                var _domain = dgcGoogleContactsService.GetDomain(service.ContactsService);
                var _query = new ContactsQuery(ContactsQuery.CreateContactsUri(_domain));
                var _feed = service.ContactsService.Query(_query);
                foreach (ContactEntry _entry in _feed.Entries)
                {
                    if (_entry.SelfUri.Content == selfUri)
                    {
                        
                        if (emailAddress != null)
                        {
                            var _primaryEmail = new EMail();
                            _primaryEmail.Address = emailAddress;
                            _primaryEmail.Primary = true;
                            _primaryEmail.Rel = ContactsRelationships.IsWork;
                            _entry.Emails.Add(_primaryEmail);
                        }

                        if (phoneNumber != null)
                        {
                            _entry.Phonenumbers.Clear();
                            var _phoneNumber = new PhoneNumber(phoneNumber);
                            _phoneNumber.Primary = true;
                            _phoneNumber.Rel = ContactsRelationships.IsWork;
                            _entry.Phonenumbers.Add(_phoneNumber);
                        }

                        if (postalAddress != null)
                        {
                            _entry.PostalAddresses.Clear();
                            var _postalAddress = new StructuredPostalAddress();
                            _postalAddress.FormattedAddress = postalAddress;
                            _postalAddress.Primary = true;
                            _postalAddress.Rel = ContactsRelationships.IsWork;
                            if (city != null)
                            {
                                _postalAddress.City = city;
                            }
                            else
                            {

                            }
                            _entry.PostalAddresses.Add(_postalAddress);           
                        }

                        Uri _feedUri = new Uri(ContactsQuery.CreateContactsUri(_domain));

                        try
                        {
                            ContactEntry _updateEntry = (ContactEntry)service.ContactsService.Update(_entry);
                            if (name != null)
                            {
                                var _token = service.ContactsService.QueryClientLoginToken();
                                dgcGoogleContactsService.SetContactTitle(_token, _entry.SelfUri.ToString(), name);
                                var _contactEntry = dgcGoogleContactsService.CreateContactModifidEntry(_updateEntry, name);
                                WriteObject(_contactEntry);
                            }
                            else
                            {
                                var _contactEntry = dgcGoogleContactsService.CreateContactEntry(_updateEntry);
                                WriteObject(_contactEntry);
                            }
                        }
                        catch (Exception _exception)
                        {
                            WriteObject(_exception);
                        }
                    }
                }
            }
        }

        #endregion Set-GDataContact

        #region New-GDataContact

        [Cmdlet(VerbsCommon.New, "GDataContact")]
        public class NewGDataContact : Cmdlet
        {
            #region Parameters


            [Parameter(
            Mandatory = true,
            HelpMessage = "ContactService, new-GdataContactService"
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            
            [Parameter(
               Mandatory = true,
               HelpMessage = "Contact EmailAddress"
            )]
            [ValidateNotNullOrEmpty]
            public string EmailAddress
            {
                get { return null; }
                set { emailAddress = value; }
            }
            private string emailAddress;
            
            [Parameter(
               Mandatory = true,
               HelpMessage = "Contact Name"
            )]
            [ValidateNotNullOrEmpty]
            public string Name
            {
                get { return null; }
                set { name = value; }
            }
            private string name;
            
            [Parameter(
               Mandatory = false,
               HelpMessage = "Contact PhoneNumber"
            )]
            [ValidateNotNullOrEmpty]
            public string PhoneNumber
            {
                get { return null; }
                set { phoneNumber = value; }
            }
            private string phoneNumber;

            
            [Parameter(
               Mandatory = false,
               HelpMessage = "Contact PostalAddress"
            )]
            [ValidateNotNullOrEmpty]
            public string PostalAddress
            {
                get { return null; }
                set { postalAddress = value; }
            }
            private string postalAddress;

            [Parameter(
            Mandatory = false,
            HelpMessage = "Contact City"
            )]
            [ValidateNotNullOrEmpty]
            public string City
            {
                get { return null; }
                set { city = value; }
            }
            private string city;
            
            #endregion Parameters

            private Dgc.GoogleContactsService dgcGoogleContactsService = new Dgc.GoogleContactsService();
            protected override void ProcessRecord()
            {
                var _newEntry = new ContactEntry();
                EMail _primaryEmail = new EMail();
                _primaryEmail.Address = emailAddress;
                _primaryEmail.Primary = true;
                _primaryEmail.Rel = ContactsRelationships.IsWork;
                _newEntry.Emails.Add(_primaryEmail);
            
                if (phoneNumber != null)
                {
                    var _phoneNumber = new PhoneNumber(phoneNumber);
                    _phoneNumber.Primary = true;
                    _phoneNumber.Rel = ContactsRelationships.IsWork;
                    _newEntry.Phonenumbers.Add(_phoneNumber);
                }

                if (postalAddress != null)
                {
                    //var postalAddress = new PostalAddress();
                    var _postalAddress = new StructuredPostalAddress();
                    //postalAddress.Value = _PostalAddress;
                    _postalAddress.FormattedAddress = postalAddress;
                    _postalAddress.Primary = true;
                    _postalAddress.Rel = ContactsRelationships.IsWork;
                    if (city != null)
                    {
                        _postalAddress.City = city;
                    }
                    
                    _newEntry.PostalAddresses.Add(_postalAddress);
                }

                var _domain = dgcGoogleContactsService.GetDomain(service.ContactsService);
                
                Uri _feedUri = new Uri(ContactsQuery.CreateContactsUri(_domain));
                try
                {
                    ContactEntry _entry = (ContactEntry)service.ContactsService.Insert(_feedUri, _newEntry);
                    var _token = service.ContactsService.QueryClientLoginToken();
                    dgcGoogleContactsService.SetContactTitle(_token,_entry.SelfUri.ToString(),name);
                    var _contactEntry = dgcGoogleContactsService.CreateContactModifidEntry(_entry, name);
                    WriteObject(_contactEntry);
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }
            }
        }

        #endregion New-GDataContact

    }
}

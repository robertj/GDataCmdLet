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

        // Todo: add scope option fore Contacs (Personal and Domain Sharied)

        #region New-GDataContactService

        [Cmdlet(VerbsCommon.New, "GDataContactService")]
        public class NewGDataContactService : Cmdlet
        {
            #region Parameters

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string AdminUsername
            {
                get { return null; }
                set { _AdminUser = value; }
            }
            private string _AdminUser;

            [Parameter(
               Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string AdminPassword
            {
                get { return null; }
                set { _AdminPassword = value; }
            }
            private string _AdminPassword;

            #endregion Parameters

            protected override void ProcessRecord()
            {
                try
                {
                    ContactsService _ContactService = new ContactsService("GData");
                    _ContactService.setUserCredentials(_AdminUser, _AdminPassword);

                    WriteObject(_ContactService);
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
                }
            }

        }

        #endregion New-GdataContactService

        #region Remove-GDataContact

        [Cmdlet(VerbsCommon.Remove, "GDataContact")]
        public class RemoveGDataContact : Cmdlet
        {

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public ContactsService ContactService
            {
                get { return null; }
                set { _ContactService = value; }
            }
            private ContactsService _ContactService;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { _SelfUri = value; }
            }
            private string _SelfUri;

            protected override void ProcessRecord()
            {



                var _DgcGoogleContactsService = new Dgc.GoogleContactsService();
                var _Domain = _DgcGoogleContactsService.GetDomain(_ContactService);

                /*
                var _DgcGoogle = new Contact.DgcGoogle();
                var _Domain = _DgcGoogle.Get(_ContactService);
                */

                var _Query = new ContactsQuery(ContactsQuery.CreateContactsUri(_Domain));
                var _Feed = _ContactService.Query(_Query);


                foreach (var _Entry in _Feed.Entries)
                {
                    if (_Entry.SelfUri.Content == _SelfUri)
                    {
                        try
                        {
                            _ContactService.Delete(_Entry);
                            WriteObject(_Entry, true);
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
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public ContactsService ContactService
            {
                get { return null; }
                set { _ContactService = value; }
            }
            private ContactsService _ContactService;
 
            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { _SelfUri = value; }
            }
            private string _SelfUri;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string Name
            {
                get { return null; }
                set { _Name = value; }
            }
            private string _Name;

            #endregion Parameters


            protected override void ProcessRecord()
            {


                var _DgcGoogleContactsService = new Dgc.GoogleContactsService();
                var _Domain = _DgcGoogleContactsService.GetDomain(_ContactService);

                var _Query = new ContactsQuery(ContactsQuery.CreateContactsUri(_Domain));
                try
                {
                    var _Feed = _ContactService.Query(_Query);
                

                    if (_SelfUri != null)
                    {
                        foreach (var _Entry in _Feed.Entries)
                        {
                            if (_Entry.SelfUri.Content == _SelfUri)
                            {
                                WriteObject(_Entry, true);
                            }

                        }
                    }
                    else if (_Name != null)
                    {
                        foreach (ContactEntry _Entry in _Feed.Entries)
                        {
                            if (System.Text.RegularExpressions.Regex.IsMatch(_Entry.Title.Text, _Name))
                            {
                                WriteObject(_Entry, true);
                            }

                        }
                    }
                    else
                    {
                        WriteObject(_Feed.Entries, true);     
                    }
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception); 

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
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public ContactsService ContactService
            {
                get { return null; }
                set { _ContactService = value; }
            }
            private ContactsService _ContactService;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { _SelfUri = value; }
            }
            private string _SelfUri;

            [Parameter(
             Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string EmailAddress
            {
                get { return null; }
                set { _EmailAddress = value; }
            }
            private string _EmailAddress;

            [Parameter(
               Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string Name
            {
                get { return null; }
                set { _Name = value; }
            }
            private string _Name;


            [Parameter(
               Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string PhoneNumber
            {
                get { return null; }
                set { _PhoneNumber = value; }
            }
            private string _PhoneNumber;


            [Parameter(
               Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string PostalAddress
            {
                get { return null; }
                set { _PostalAddress = value; }
            }
            private string _PostalAddress;


            #endregion Parameters


            protected override void ProcessRecord()
            {


                var _DgcGoogleContactsService = new Dgc.GoogleContactsService();
                var _Domain = _DgcGoogleContactsService.GetDomain(_ContactService);

                var _Query = new ContactsQuery(ContactsQuery.CreateContactsUri(_Domain));
               
                var _Feed = _ContactService.Query(_Query);

                foreach (ContactEntry _Entry in _Feed.Entries)
                {
                    if (_Entry.SelfUri.Content == _SelfUri)
                    {
                        if (_Name != null)
                        {
                            _Entry.Title.Text = _Name;
                            _Entry.Content.Content = _Name;

                        }
                        
                        if (_EmailAddress != null)
                        {
                            var primaryEmail = new EMail();
                            primaryEmail.Address = _EmailAddress;
                            primaryEmail.Primary = true;
                            primaryEmail.Rel = ContactsRelationships.IsWork;
                            _Entry.Emails.Add(primaryEmail);
                        }

                        if (_PhoneNumber != null)
                        {
                            var phoneNumber = new PhoneNumber(_PhoneNumber);
                            phoneNumber.Primary = true;
                            phoneNumber.Rel = ContactsRelationships.IsWork;
                            _Entry.Phonenumbers.Add(phoneNumber);
                        }

                        if (_PostalAddress != null)
                        {
                            var postalAddress = new PostalAddress();
                            postalAddress.Value = _PostalAddress;
                            postalAddress.Primary = true;
                            postalAddress.Rel = ContactsRelationships.IsWork;
                            _Entry.PostalAddresses.Add(postalAddress);
                        }



                        Uri _FeedUri = new Uri(ContactsQuery.CreateContactsUri(_Domain));

                        try
                        {
                            ContactEntry _UpdatedEntry = (ContactEntry)_ContactService.Update(_Entry);
                            WriteObject(_UpdatedEntry, true);
                        }
                        catch (Exception _Exception)
                        {
                            WriteObject(_Exception);
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
            Mandatory = true
             )]
            [ValidateNotNullOrEmpty]
            public ContactsService ContactService
            {
                get { return null; }
                set { _ContactService = value; }
            }
            private ContactsService _ContactService;

            
            [Parameter(
               Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string EmailAddress
            {
                get { return null; }
                set { _EmailAddress = value; }
            }
            private string _EmailAddress;

            [Parameter(
               Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string Name
            {
                get { return null; }
                set { _Name = value; }
            }
            private string _Name;


            [Parameter(
               Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string PhoneNumber
            {
                get { return null; }
                set { _PhoneNumber = value; }
            }
            private string _PhoneNumber;

            
            [Parameter(
               Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string PostalAddress
            {
                get { return null; }
                set { _PostalAddress = value; }
            }
            private string _PostalAddress;
            
            #endregion Parameters


            protected override void ProcessRecord()
            {
                
                var _NewEntry = new ContactEntry();
                _NewEntry.Title.Text = _Name;
                EMail primaryEmail = new EMail();
                primaryEmail.Address = _EmailAddress;
                primaryEmail.Primary = true;
                primaryEmail.Rel = ContactsRelationships.IsWork;
                _NewEntry.Emails.Add(primaryEmail);

                if (_PhoneNumber != null)
                {
                    var phoneNumber = new PhoneNumber(_PhoneNumber);
                    phoneNumber.Primary = true;
                    phoneNumber.Rel = ContactsRelationships.IsWork;
                    _NewEntry.Phonenumbers.Add(phoneNumber);
                }

                if (_PostalAddress != null)
                {
                    var postalAddress = new PostalAddress();
                    postalAddress.Value = _PostalAddress;
                    postalAddress.Primary = true;
                    postalAddress.Rel = ContactsRelationships.IsWork;
                    _NewEntry.PostalAddresses.Add(postalAddress);
                }
                _NewEntry.Content.Content = _Name;

                var _DgcGoogleContactsService = new Dgc.GoogleContactsService();
                var _Domain = _DgcGoogleContactsService.GetDomain(_ContactService);

                Uri _FeedUri = new Uri(ContactsQuery.CreateContactsUri(_Domain));


                try
                {
                    ContactEntry _CreatedEntry = (ContactEntry)_ContactService.Insert(_FeedUri, _NewEntry);
                    WriteObject(_CreatedEntry, true);
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
                }
            }




        }

        #endregion New-GDataContact

    }
}

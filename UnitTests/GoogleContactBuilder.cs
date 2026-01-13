using Google.Apis.PeopleService.v1.Data;
using System.Collections.Generic;

namespace GoContactSyncMod.UnitTests
{
    public class GoogleContactBuilder
    {
        public Person Build()
        {
            return new Person();
        }

        public Person Build(string bio)
        {
            var newEntry = new Person();
            var name = ContactPropertiesUtils.GetGooglePrimaryName(newEntry);
            if (name == null)
            {
                name = new Name();
                if (newEntry.Names == null)
                    newEntry.Names = new List<Name>();
                newEntry.Names.Add(name);
            }
            name.UnstructuredName = SyncContactsTests.TEST_CONTACT_NAME;

            var primaryEmail = new EmailAddress
            {
                Value = SyncContactsTests.TEST_CONTACT_EMAIL,
                Metadata = new FieldMetadata { Primary = true },
                Type = ContactSync.WORK
            };
            if (newEntry.EmailAddresses == null)
                newEntry.EmailAddresses = new List<EmailAddress>();
            newEntry.EmailAddresses.Add(primaryEmail);

            var phoneNumber = new PhoneNumber
            {
                Value = "555-555-5551",
                Metadata = new FieldMetadata { Primary = true },
                Type = ContactSync.PHONE_MOBILE
            };
            if (newEntry.PhoneNumbers == null)
                newEntry.PhoneNumbers = new List<PhoneNumber>();
            newEntry.PhoneNumbers.Add(phoneNumber);

            var address = new Address
            {
                StreetAddress = "123 somewhere lane",
                Metadata = new FieldMetadata { Primary = true },
                Type = ContactSync.HOME
            };
            if (newEntry.Addresses == null)
                newEntry.Addresses = new List<Address>();
            newEntry.Addresses.Add(address);

            if (newEntry.Biographies == null)
                newEntry.Biographies = new List<Biography>();
            newEntry.Biographies.Add(new Biography { Value = bio });
            return newEntry;
        }
    }
}

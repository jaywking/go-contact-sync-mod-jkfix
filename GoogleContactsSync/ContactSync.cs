using Google.Apis.PeopleService.v1.Data;
using Serilog;
using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    internal static class ContactSync
    {
        internal static DateTime outlookDateNone = new DateTime(4501, 1, 1);
        internal const string REL_SPOUSE = "spouse";
        internal const string REL_CHILD = "child";
        internal const string REL_MANAGER = "manager";
        internal const string REL_ASSISTANT = "assistant";
        internal const string EVENT_ANNIVERSARY = "anniversary";
        internal const string URL_HOMEPAGE = "homePage";
        internal const string WORK = "work";
        internal const string HOME = "home";
        internal const string OTHER = "other";
        internal const string ANDERE = "Andere";
        internal const string ASSISTANT = "assistant"; //Not existing anymore on Google side as standard
        internal const string CAR = "car";             //Not existing anymore on Google side as standard
        internal const string PHONE_MAIN = "main";
        internal const string PHONE_MOBILE = "mobile";
        internal const string PHONE_CALLBACK = "callback";
        internal const string PHONE_RADIO = "radio";
        internal const string PHONE_TTY = "tty";
        internal const string PHONE_COMPANY = "company";
        internal const string FAX_WORK = "workFax";
        internal const string FAX_OTHER = "otherFax";
        internal const string FAX_WEITER = "Weitere Faxnummer";
        internal const string FAX_HOME = "homeFax";
        internal const string PHONE_PAGER = "pager";
        internal const string DESK = "desk";

        #region addresses
        private static void SetAddress(IList<Address> destination, string address, string street, string city, string PostalCode, string pobox, string region, string country, string type)
        {
            if (!string.IsNullOrEmpty(address))
            {
                var postalAddress = new Address
                {
                    StreetAddress = street,
                    City = city,
                    PostalCode = PostalCode,
                    PoBox = pobox,
                    Region = region,
                    Metadata = new FieldMetadata { Primary = destination.Count == 0 },
                    Type = type

                };


                //By default Outlook is not setting Country in formatted string in case Windows is configured for the same country 
                //(Control Panel\Regional Settings).  So set country in Google only if Outlook address has it
                if (address.EndsWith("\r\n" + country))
                {
                    postalAddress.Country = country;
                }

                destination.Add(postalAddress);
            }
        }

        internal static void SetAddresses(Outlook.ContactItem master, Person slave)
        {
            //validate if maybe at Google contact you have more postal addresses than Outlook can handle
            var IsHomeCount = 0;
            var IsWorkCount = 0;
            var IsOtherCount = 0;

            if (slave.Addresses == null)
                slave.Addresses = new List<Address>();

            foreach (var address in slave.Addresses)
            {
                if (address != null && !string.IsNullOrWhiteSpace(address.Type))
                    switch (address.Type.ToLower().Trim())
                    {
                        case HOME:    //ToDo: Find correct enum
                            IsHomeCount++;
                            break;
                        case WORK:
                            IsWorkCount++;
                            break;
                        case OTHER:
                        case ANDERE://ToDo: Really camel case? Temporary fix, add "andere"
                        case "andere":
                            IsOtherCount++;
                            break;
                        default:
                            Log.Warning($"Google contact \"{slave.ToLogString()}\" has custom address type \"{address.Type}\", this address is not synchronized with Outlook. Please update contact at Google and change type to standard one (home, work, other).");
                            break;
                    }
            }

            if (IsHomeCount > 1)
            {
                Log.Warning($"Google contact \"{slave.ToLogString()}\" has {IsHomeCount} home addresses, Outlook can have only 1 home address. You will lose information about additional home addresses.");
            }
            if (IsWorkCount > 1)
            {
                Log.Warning($"Google contact \"{slave.ToLogString()}\" has {IsWorkCount} business addresses, Outlook can have only 1 business address. You will lose information about additional business addresses.");
                Log.Debug("### Dump ###");
                foreach (var address in slave.Addresses)
                {
                    if (address != null && address.Type != null && address.Type.Trim().Equals(WORK, StringComparison.InvariantCultureIgnoreCase)) //ToDo: Find proper enum
                    {
                        Log.Debug($"Street: {address.StreetAddress}");
                        Log.Debug($"City: {address.City}");
                    }
                }
                Log.Debug("### Dump ###");
            }
            if (IsOtherCount > 1)
            {
                Log.Warning($"Google contact \"{slave.ToLogString()}\" has {IsOtherCount} other addresses, Outlook can have only 1 other address. You will lose information about additional other addresses.");
            }

            //clear only addresses synchronized with Outlook, this will prevent losing Google addresses with custom labels
            for (var i = slave.Addresses.Count - 1; i >= 0; i--)
            {
                if (slave.Addresses[i] != null && !string.IsNullOrWhiteSpace(slave.Addresses[i].Type)
                    && (slave.Addresses[i].Type.Trim().Equals(HOME, StringComparison.InvariantCultureIgnoreCase) || //ToDo: Find proper enum
                        slave.Addresses[i].Type.Trim().Equals(WORK, StringComparison.InvariantCultureIgnoreCase) ||
                        slave.Addresses[i].Type.Trim().Equals(OTHER, StringComparison.InvariantCultureIgnoreCase) ||
                        slave.Addresses[i].Type.Trim().Equals(ANDERE, StringComparison.InvariantCultureIgnoreCase))
                    )
                {
                    slave.Addresses.RemoveAt(i);
                }
            };

            SetAddress(slave.Addresses,
                    master.HomeAddress,
                    master.HomeAddressStreet,
                    master.HomeAddressCity,
                    master.HomeAddressPostalCode,
                    master.HomeAddressPostOfficeBox,
                    master.HomeAddressState,
                    master.HomeAddressCountry,
                    HOME
            );

            SetAddress(slave.Addresses,
                    master.BusinessAddress,
                    master.BusinessAddressStreet,
                    master.BusinessAddressCity,
                    master.BusinessAddressPostalCode,
                    master.BusinessAddressPostOfficeBox,
                    master.BusinessAddressState,
                    master.BusinessAddressCountry,
                    WORK
            );

            SetAddress(slave.Addresses,
                    master.OtherAddress,
                    master.OtherAddressStreet,
                    master.OtherAddressCity,
                    master.OtherAddressPostalCode,
                    master.OtherAddressPostOfficeBox,
                    master.OtherAddressState,
                    master.OtherAddressCountry,
                    OTHER
            );
        }

        private static void SetAddress(Address address, Outlook.ContactItem destination)
        {
            if (address != null && address.Type != null && address.Type.Trim().Equals(HOME,StringComparison.InvariantCultureIgnoreCase))
            {
                destination.HomeAddressStreet = address.StreetAddress;
                destination.HomeAddressCity = address.City;
                destination.HomeAddressPostalCode = address.PostalCode;
                destination.HomeAddressCountry = address.Country;
                destination.HomeAddressState = address.Region;
                destination.HomeAddressPostOfficeBox = address.PoBox;

                //Workaround because of Google bug: If a contact was created on GOOGLE side, it uses the unstructured approach
                //Therefore we need to check, if the structure was filled, if yes it resulted in a formatted string in the Address summary field
                //If not, the formatted string is null => overwrite it with the formmatedAddress from Google
                if (string.IsNullOrEmpty(destination.HomeAddress))
                {
                    destination.HomeAddress = address.FormattedValue;
                }

                //By default Outlook is not setting Country in formatted string in case Windows is configured for the same country 
                //(Control Panel\Regional Settings).  So append country, but to be on safe side only if Google address has it
                if (!string.IsNullOrEmpty(address.Country) &&
                    !string.IsNullOrEmpty(address.FormattedValue) &&
                    address.FormattedValue.EndsWith("\n" + address.Country) &&
                    !string.IsNullOrEmpty(destination.HomeAddress) &&
                    !destination.HomeAddress.EndsWith("\r\n" + address.Country))
                {
                    destination.HomeAddress = destination.HomeAddress + "\r\n" + address.Country;
                }

                if (address.Metadata != null && (address.Metadata.Primary ?? false))
                {
                    destination.SelectedMailingAddress = Outlook.OlMailingAddress.olHome;
                }
            }
            else if (address != null && address.Type != null && address.Type.Trim().Equals(WORK,StringComparison.InvariantCultureIgnoreCase))
            {
                destination.BusinessAddressStreet = address.StreetAddress;
                destination.BusinessAddressCity = address.City;
                destination.BusinessAddressPostalCode = address.PostalCode;
                destination.BusinessAddressCountry = address.Country;
                destination.BusinessAddressState = address.Region;
                destination.BusinessAddressPostOfficeBox = address.PoBox;

                //Workaround because of Google bug: If a contact was created on GOOGLE side, it uses the unstructured approach
                //Therefore we need to check, if the structure was filled, if yes it resulted in a formatted string in the Address summary field
                //If not, the formatted string is null => overwrite it with the formmatedAddress from Google
                if (string.IsNullOrEmpty(destination.BusinessAddress))
                {
                    destination.BusinessAddress = address.FormattedValue;
                }

                //By default Outlook is not setting Country in formatted string in case Windows is configured for the same country 
                //(Control Panel\Regional Settings).  So append country, but to be on safe side only if Google address has it
                if (!string.IsNullOrEmpty(address.Country) &&
                    !string.IsNullOrEmpty(address.FormattedValue) &&
                    address.FormattedValue.EndsWith("\n" + address.Country) &&
                    !string.IsNullOrEmpty(destination.BusinessAddress) &&
                    !destination.BusinessAddress.EndsWith("\r\n" + address.Country))
                {
                    destination.BusinessAddress = destination.BusinessAddress + "\r\n" + address.Country;
                }

                if (address.Metadata != null && (address.Metadata.Primary ?? false))
                {
                    destination.SelectedMailingAddress = Outlook.OlMailingAddress.olBusiness;
                }
            }
            else if (address != null && address.Type != null && (address.Type.Trim().Equals(OTHER,StringComparison.InvariantCultureIgnoreCase) || address.Type.Trim().Equals(ANDERE,StringComparison.InvariantCultureIgnoreCase)))
            {
                destination.OtherAddressStreet = address.StreetAddress;
                destination.OtherAddressCity = address.City;
                destination.OtherAddressPostalCode = address.PostalCode;
                destination.OtherAddressCountry = address.Country;
                destination.OtherAddressState = address.Region;
                destination.OtherAddressPostOfficeBox = address.PoBox;

                //Workaround because of Google bug: If a contact was created on GOOGLE side, it uses the unstructured approach
                //Therefore we need to check, if the structure was filled, if yes it resulted in a formatted string in the Address summary field
                //If not, the formatted string is null => overwrite it with the formmatedAddress from Google
                if (string.IsNullOrEmpty(destination.OtherAddress))
                {
                    destination.OtherAddress = address.FormattedValue;
                }

                //By default Outlook is not setting Country in formatted string in case Windows is configured for the same country 
                //(Control Panel\Regional Settings).  So append country, but to be on safe side only if Google address has it
                if (!string.IsNullOrEmpty(address.Country) &&
                    !string.IsNullOrEmpty(address.FormattedValue) &&
                    address.FormattedValue.EndsWith("\n" + address.Country) &&
                    !string.IsNullOrEmpty(destination.OtherAddress) &&
                    !destination.OtherAddress.EndsWith("\r\n" + address.Country))
                {
                    destination.OtherAddress = destination.OtherAddress + "\r\n" + address.Country;
                }

                if (address.Metadata != null && (address.Metadata.Primary ?? false))
                {
                    destination.SelectedMailingAddress = Outlook.OlMailingAddress.olOther;
                }
            }
        }

        private static void SetAddresses(Person master, Outlook.ContactItem slave)
        {
            //validate if maybe at Google contact you have more postal addresses than Outlook can handle
            var IsHomeCount = 0;
            var IsWorkCount = 0;
            var IsOtherCount = 0;

            if (master.Addresses == null)
                master.Addresses = new List<Address>();

            foreach (var address in master.Addresses)
            {
                if (address != null && !string.IsNullOrWhiteSpace(address.Type))
                    switch (address.Type.ToLower().Trim())
                    {
                        case HOME:  //ToDo: Get Proper enum
                            IsHomeCount++;
                            break;
                        case WORK:
                            IsWorkCount++;
                            break;
                        case OTHER:
                        case ANDERE: //ToDo: Really uppercase? Temporary fix: add also "andere":
                        case "andere":
                            IsOtherCount++;
                            break;
                        default:
                            Log.Warning($"Google contact \"{master.ToLogString()}\" has custom address number type \"{address.Type}\", this address is not synchronized with Outlook. Please update contact at Google and change label to standard one (home, work, other).");
                            break;
                    }
                
            }

            if (IsHomeCount > 1)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {IsHomeCount} home addresses, Outlook can have only 1 home address. Please update contact at Google, otherwise you will lose information about additional home addresses.");
            }
            if (IsWorkCount > 1)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {IsWorkCount} business addresses, Outlook can have only 1 business address. Please update contact at Google, otherwise you will lose information about additional business addresses.");
                Log.Debug("### Dump ###");
                foreach (var address in master.Addresses)
                {
                    if (address.Type == WORK)
                    {
                        Log.Debug($"Street: {address.StreetAddress}");
                        Log.Debug($"City: {address.City}");
                    }
                }
                Log.Debug("### Dump ###");
            }
            if (IsOtherCount > 1)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {IsOtherCount} other addresses, Outlook can have only 1 other address. Please update contact at Google, otherwise you will lose information about additional other addresses.");
            }

            slave.HomeAddress = string.Empty;
            slave.HomeAddressStreet = string.Empty;
            slave.HomeAddressCity = string.Empty;
            slave.HomeAddressPostalCode = string.Empty;
            slave.HomeAddressCountry = string.Empty;
            slave.HomeAddressState = string.Empty;
            slave.HomeAddressPostOfficeBox = string.Empty;

            slave.BusinessAddress = string.Empty;
            slave.BusinessAddressStreet = string.Empty;
            slave.BusinessAddressCity = string.Empty;
            slave.BusinessAddressPostalCode = string.Empty;
            slave.BusinessAddressCountry = string.Empty;
            slave.BusinessAddressState = string.Empty;
            slave.BusinessAddressPostOfficeBox = string.Empty;

            slave.OtherAddress = string.Empty;
            slave.OtherAddressStreet = string.Empty;
            slave.OtherAddressCity = string.Empty;
            slave.OtherAddressPostalCode = string.Empty;
            slave.OtherAddressCountry = string.Empty;
            slave.OtherAddressState = string.Empty;
            slave.OtherAddressPostOfficeBox = string.Empty;

            slave.SelectedMailingAddress = Outlook.OlMailingAddress.olNone;
            foreach (var address in master.Addresses)
            {
                SetAddress(address, slave);
            }
        }
        #endregion

        #region phones
        private static void SetPhoneNumber(IList<PhoneNumber> phones, string number, string type)
        {
            if (!string.IsNullOrWhiteSpace(number))
            {
                var phoneNumber = new PhoneNumber()
                {
                    Metadata = new FieldMetadata() { Primary = phones.Count == 0 },
                    Value = number,
                    Type = type
                };
                phones.Add(phoneNumber);
            }
        }

        internal static void SetPhoneNumbers(Outlook.ContactItem master, Person slave)
        {
            //validate if maybe at Google contact you have more phones than Outlook can handle
            var IsMainCount = 0;
            var IsHomeCount = 0;
            var IsWorkCount = 0;
            var IsMobileCount = 0;
            var IsWorkFaxCount = 0;
            var IsOtherFaxCount = 0;
            var IsHomeFaxCount = 0;
            var IsPagerCount = 0;
            var IsOtherCount = 0;
            var IsCarCount = 0;
            var IsAssistantCount = 0;
            var IsCallbackCount = 0;
            var IsRadioCount = 0;
            var IsTtyCount = 0;
            var IsCompanyCount = 0;

            if (slave.PhoneNumbers == null)
                slave.PhoneNumbers = new List<PhoneNumber>();

            foreach (var phone in slave.PhoneNumbers)
            {
                if (phone != null && !string.IsNullOrWhiteSpace(phone.Type))
                    switch (phone.Type.ToLower().Trim())
                    {
                        case PHONE_MAIN:
                            IsMainCount++;
                            break;
                        case HOME:     //ToDo:Find proper enum
                            IsHomeCount++;
                            break;
                        case WORK:
                            IsWorkCount++;
                            break;
                        case PHONE_MOBILE:
                            IsMobileCount++;
                            break;
                        case FAX_WORK://ToDo: Really camel case? temporary fix below
                        case "workfax":
                            IsWorkFaxCount++;
                            break;
                        case FAX_OTHER://ToDo: Really camel case? temporary fix below
                        case "otherfax":
                        case FAX_WEITER://ToDo: Really camel case? temporary fix below
                        case "weitere faxnummer":
                            IsOtherFaxCount++;
                            break;
                        case FAX_HOME://ToDo: Really camel case? temporary fix below
                        case "homefax":
                            IsHomeFaxCount++;
                            break;
                        case PHONE_PAGER:
                            IsPagerCount++;
                            break;
                        case OTHER:
                        case ANDERE: //ToDo: Really camel case? temporary fix below
                        case "andere":
                            IsOtherCount++;
                            break;
                        case CAR:
                            IsCarCount++;
                            break;
                        case ASSISTANT:
                            IsAssistantCount++;
                            break;
                        case PHONE_CALLBACK:
                            IsCallbackCount++;
                            break;
                        case PHONE_RADIO:
                            IsRadioCount++;
                            break;
                        case PHONE_TTY:
                            IsTtyCount++;
                            break;
                        case PHONE_COMPANY:
                            IsCompanyCount++;
                            break;
                        default:
                            Log.Debug($"Google contact \"{master.ToLogString()}\" has custom phone number labeled \"{phone.Type}\", this phone number is not synchronized with Outlook. Please update contact at Google and change label to standard one (home, work, other, ...).");
                            break;
                    }
            }

            if (IsMainCount > 1)
            {
                Log.Debug($"Google contact \"{slave.ToLogString()}\" has {IsMainCount} main phone numbers, Outlook can have only 1 main phone number. You will lose information about additional main phone numbers.");
            }
            if (IsHomeCount > 2)
            {
                Log.Debug($"Google contact \"{slave.ToLogString()}\" has {IsHomeCount} home phone numbers, Outlook can have only 2 home phone numbers. You will lose information about additional home phone numbers.");
            }
            if (IsWorkCount > 2)
            {
                Log.Debug($"Google contact \"{slave.ToLogString()}\" has {IsWorkCount} work phone numbers, Outlook can have only 2 work phone numbers. You will lose information about additional work phone numbers.");
            }
            if (IsMobileCount > 1)
            {
                Log.Debug($"Google contact \"{slave.ToLogString()}\" has {IsMobileCount} mobile phone numbers, Outlook can have only 1 mobile phone number. You will lose information about additional mobile phone numbers.");
            }
            if (IsWorkFaxCount > 1)
            {
                Log.Debug($"Google contact \"{slave.ToLogString()}\" has {IsWorkFaxCount} work fax phone numbers, Outlook can have only 1 work fax phone number. You will lose information about additional work fax phone numbers.");
            }
            if (IsOtherFaxCount > 1)
            {
                Log.Debug($"Google contact \"{slave.ToLogString()}\" has {IsOtherFaxCount} other fax phone numbers, Outlook can have only 1 other fax phone number. You will lose information about additional other fax phone numbers.");
            }
            if (IsHomeFaxCount > 1)
            {
                Log.Debug($"Google contact \"{slave.ToLogString()}\" has {IsHomeFaxCount} home fax phone numbers, Outlook can have only 1 home fax phone number. You will lose information about additional home fax phone numbers.");
            }
            if (IsPagerCount > 1)
            {
                Log.Debug($"Google contact \"{slave.ToLogString()}\" has {IsPagerCount} pager phone numbers, Outlook can have only 1 pager phone number. You will lose information about additional pager phone numbers.");
            }
            if (IsOtherCount > 1)
            {
                Log.Debug($"Google contact \"{slave.ToLogString()}\" has {IsOtherCount} other phone numbers, Outlook can have only 1 other phone number. You will lose information about additional other phone numbers.");
            }
            if (IsCarCount > 1)
            {
                Log.Debug($"Google contact \"{slave.ToLogString()}\" has {IsCarCount} car phone numbers, Outlook can have only 1 car phone number. You will lose information about additional car phone numbers.");
            }
            if (IsAssistantCount > 1)
            {
                Log.Debug($"Google contact \"{slave.ToLogString()}\" has {IsAssistantCount} assistant phone numbers, Outlook can have only 1 assistant phone number. You will lose information about additional assistant phone numbers.");
            }
            if (IsCallbackCount > 1)
            {
                Log.Debug($"Google contact \"{slave.ToLogString()}\" has {IsCallbackCount} callback phone numbers, Outlook can have only 1 callback phone number. You will lose information about additional callback phone numbers.");
            }
            if (IsRadioCount > 1)
            {
                Log.Debug($"Google contact \"{slave.ToLogString()}\" has {IsRadioCount} radio phone numbers, Outlook can have only 1 radio phone number. You will lose information about additional radio phone numbers.");
            }
            if (IsTtyCount > 1)
            {
                Log.Debug($"Google contact \"{slave.ToLogString()}\" has {IsTtyCount} assistant TTY numbers, Outlook can have only 1 TTY phone number. You will lose information about additional TTY phone numbers.");
            }
            if (IsCompanyCount > 1)
            {
                Log.Debug($"Google contact \"{slave.ToLogString()}\" has {IsCompanyCount} company phone numbers, Outlook can have only 1 company phone number. You will lose information about additional company phone numbers.");
            }

            //clear only phone numbers synchronized with Outlook, this will prevent losing Google phones with custom labels
            for (var i = slave.PhoneNumbers.Count - 1; i >= 0; i--)
            {
                if (slave.PhoneNumbers[i] != null && slave.PhoneNumbers[i].Type != null
                    && (slave.PhoneNumbers[i].Type.Trim().Equals(PHONE_MAIN, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(PHONE_MOBILE, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(HOME, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(WORK, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(FAX_HOME, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(FAX_WORK, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(FAX_OTHER, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(FAX_WEITER, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(OTHER, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(ANDERE, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(PHONE_PAGER, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(CAR, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(ASSISTANT, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(PHONE_CALLBACK, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(PHONE_RADIO, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(PHONE_TTY, StringComparison.InvariantCultureIgnoreCase) ||
                    slave.PhoneNumbers[i].Type.Trim().Equals(PHONE_COMPANY, StringComparison.InvariantCultureIgnoreCase))
                    )
                {
                    slave.PhoneNumbers.RemoveAt(i);
                }
            };

            SetPhoneNumber(slave.PhoneNumbers, master.PrimaryTelephoneNumber, PHONE_MAIN);
            SetPhoneNumber(slave.PhoneNumbers, master.MobileTelephoneNumber, PHONE_MOBILE);
            SetPhoneNumber(slave.PhoneNumbers, master.HomeTelephoneNumber, HOME);
            SetPhoneNumber(slave.PhoneNumbers, master.Home2TelephoneNumber, HOME);
            SetPhoneNumber(slave.PhoneNumbers, master.BusinessTelephoneNumber, WORK);
            SetPhoneNumber(slave.PhoneNumbers, master.Business2TelephoneNumber, WORK);
            SetPhoneNumber(slave.PhoneNumbers, master.HomeFaxNumber, FAX_HOME);
            SetPhoneNumber(slave.PhoneNumbers, master.BusinessFaxNumber, FAX_WORK);
            SetPhoneNumber(slave.PhoneNumbers, master.OtherFaxNumber, FAX_OTHER);
            SetPhoneNumber(slave.PhoneNumbers, master.OtherTelephoneNumber, OTHER);
            SetPhoneNumber(slave.PhoneNumbers, master.PagerNumber, PHONE_PAGER);
            SetPhoneNumber(slave.PhoneNumbers, master.CarTelephoneNumber, CAR);
            SetPhoneNumber(slave.PhoneNumbers, master.AssistantTelephoneNumber, ASSISTANT);
            SetPhoneNumber(slave.PhoneNumbers, master.CallbackTelephoneNumber, PHONE_CALLBACK);
            SetPhoneNumber(slave.PhoneNumbers, master.RadioTelephoneNumber, PHONE_RADIO);
            SetPhoneNumber(slave.PhoneNumbers, master.TTYTDDTelephoneNumber, PHONE_TTY);
            SetPhoneNumber(slave.PhoneNumbers, master.CompanyMainTelephoneNumber, PHONE_COMPANY);
        }

        private static void SetPhoneNumber(PhoneNumber phone, Outlook.ContactItem destination)
        {
            if (phone != null && !string.IsNullOrWhiteSpace(phone.Type))
            {
                switch (phone.Type.ToLower().Trim())
                {
                    case PHONE_MAIN:
                        destination.PrimaryTelephoneNumber = phone.Value;
                        break;
                    case HOME:
                        if (destination.HomeTelephoneNumber == null)
                        {
                            destination.HomeTelephoneNumber = phone.Value;
                        }
                        else
                        {
                            destination.Home2TelephoneNumber = phone.Value;
                        }
                        break;
                    case WORK:
                        if (destination.BusinessTelephoneNumber == null)
                        {
                            destination.BusinessTelephoneNumber = phone.Value;
                        }
                        else
                        {
                            destination.Business2TelephoneNumber = phone.Value;
                        }
                        break;
                    case PHONE_MOBILE:
                        destination.MobileTelephoneNumber = phone.Value;
                        break;
                    case FAX_WORK:
                    case "workfax":
                        destination.BusinessFaxNumber = phone.Value;
                        break;
                    case FAX_OTHER:
                    case "otherfax":
                    case FAX_WEITER:
                    case "weitere faxnummer":
                        destination.OtherFaxNumber = phone.Value;
                        break;
                    case FAX_HOME:
                    case "homefax":
                        destination.HomeFaxNumber = phone.Value;
                        break;
                    case PHONE_PAGER:
                        destination.PagerNumber = phone.Value;
                        break;
                    case OTHER:
                    case ANDERE:
                    case "andere":
                        destination.OtherTelephoneNumber = phone.Value;
                        break;
                    case CAR:
                        destination.CarTelephoneNumber = phone.Value;
                        break;
                    case ASSISTANT:
                        destination.AssistantTelephoneNumber = phone.Value;
                        break;
                    case PHONE_CALLBACK:
                        destination.CallbackTelephoneNumber = phone.Value;
                        break;
                    case PHONE_RADIO:
                        destination.RadioTelephoneNumber = phone.Value;
                        break;
                    case PHONE_TTY:
                        destination.TTYTDDTelephoneNumber = phone.Value;
                        break;
                    case PHONE_COMPANY:
                        destination.CompanyMainTelephoneNumber = phone.Value;
                        break;
                }
            }
        }

        private static void SetPhoneNumbers(Person master, Outlook.ContactItem slave)
        {
            slave.PrimaryTelephoneNumber = string.Empty;
            slave.HomeTelephoneNumber = string.Empty;
            slave.Home2TelephoneNumber = string.Empty;
            slave.BusinessTelephoneNumber = string.Empty;
            slave.Business2TelephoneNumber = string.Empty;
            slave.MobileTelephoneNumber = string.Empty;
            slave.BusinessFaxNumber = string.Empty;
            slave.HomeFaxNumber = string.Empty;
            slave.PagerNumber = string.Empty;
            slave.OtherTelephoneNumber = string.Empty;
            slave.CarTelephoneNumber = string.Empty;
            slave.AssistantTelephoneNumber = string.Empty;

            var IsMainCount = 0;
            var IsHomeCount = 0;
            var IsWorkCount = 0;
            var IsMobileCount = 0;
            var IsWorkFaxCount = 0;
            var IsOtherFaxCount = 0;
            var IsHomeFaxCount = 0;
            var IsPagerCount = 0;
            var IsOtherCount = 0;
            var IsCarCount = 0;
            var IsAssistantCount = 0;
            var IsCallbackCount = 0;
            var IsRadioCount = 0;
            var IsTtyCount = 0;
            var IsCompanyCount = 0;

            var emptyCount = 0;

            if (master.PhoneNumbers == null)
                master.PhoneNumbers = new List<PhoneNumber>();

            foreach (var phone in master.PhoneNumbers)
            {
                if (phone == null || string.IsNullOrWhiteSpace(phone.Type))
                {
                    emptyCount++;
                    continue;
                }

                switch (phone.Type.ToLower().Trim())
                {
                    case PHONE_MAIN:
                        IsMainCount++;
                        break;
                    case HOME:
                        IsHomeCount++;
                        break;
                    case WORK:
                        IsWorkCount++;
                        break;
                    case PHONE_MOBILE:
                        IsMobileCount++;
                        break;
                    case FAX_WORK:
                    case "workfax":
                        IsWorkFaxCount++;
                        break;
                    case FAX_OTHER:
                    case "otherfax":
                    case FAX_WEITER:
                    case "weitere faxnummer":
                        IsOtherFaxCount++;
                        break;
                    case FAX_HOME:
                    case "homefax":
                        IsHomeFaxCount++;
                        break;
                    case PHONE_PAGER:
                        IsPagerCount++;
                        break;
                    case OTHER:
                    case ANDERE:
                    case "andere": //ToDo
                        IsOtherCount++;
                        break;
                    case CAR:
                        IsCarCount++;
                        break;
                    case ASSISTANT:
                        IsAssistantCount++;
                        break;
                    case PHONE_CALLBACK:
                        IsCallbackCount++;
                        break;
                    case PHONE_RADIO:
                        IsRadioCount++;
                        break;
                    case PHONE_TTY:
                        IsTtyCount++;
                        break;
                    case PHONE_COMPANY:
                        IsCompanyCount++;
                        break;
                    default:
                        Log.Warning($"Google contact \"{master.ToLogString()}\" has custom phone number labeled \"{phone.Type}\", this phone number is not synchronized with Outlook. Please update contact at Google and change label to standard one (home, work, other, ...).");
                        break;
                }
            }

            var count = emptyCount + IsHomeCount + IsWorkCount + IsMobileCount + IsOtherCount + IsCarCount;
            if (emptyCount > 0 && count > 7)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {emptyCount} phone numbers without label and {count - emptyCount} phone numbers with labels home/work/other/car, Outlook can have only 7 phone numbers of these types. Please update contact at Google, otherwise you will lose information about additional main phone numbers.");
            }
            if (IsMainCount > 1)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {IsMainCount} main phone numbers, Outlook can have only 1 main phone number. Please update contact at Google, otherwise you will lose information about additional main phone numbers.");
            }
            if (IsHomeCount > 2)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {IsHomeCount} home phone numbers, Outlook can have only 2 home phone numbers. Please update contact at Google, otherwise you will lose information about additional home phone numbers.");
            }
            if (IsWorkCount > 2)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {IsWorkCount} work phone numbers, Outlook can have only 2 work phone numbers. Please update contact at Google, otherwise you will lose information about additional work phone numbers.");
            }
            if (IsMobileCount > 1)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {IsMobileCount} mobile phone numbers, Outlook can have only 1 mobile phone number. Please update contact at Google, otherwise you will lose information about additional mobile phone numbers.");
            }
            if (IsWorkFaxCount > 1)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {IsWorkFaxCount} work fax phone numbers, Outlook can have only 1 work fax phone number. Please update contact at Google, otherwise you will lose information about additional work fax phone numbers.");
            }
            if (IsOtherFaxCount > 1)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {IsOtherFaxCount} other fax phone numbers, Outlook can have only 1 othe fax phone number. Please update contact at Google, otherwise you will lose information about additional other fax phone numbers.");
            }
            if (IsHomeFaxCount > 1)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {IsHomeFaxCount} home fax phone numbers, Outlook can have only 1 home fax phone number. Please update contact at Google, otherwise you will lose information about additional home fax phone numbers.");
            }
            if (IsPagerCount > 1)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {IsPagerCount} pager phone numbers, Outlook can have only 1 pager phone number. Please update contact at Google, otherwise you will lose information about additional pager phone numbers.");
            }
            if (IsOtherCount > 1)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {IsOtherCount} other phone numbers, Outlook can have only 1 other phone number. Please update contact at Google, otherwise you will lose information about additional other phone numbers.");
            }
            if (IsCarCount > 1)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {IsCarCount} car phone numbers, Outlook can have only 1 car phone number. Please update contact at Google, otherwise you will lose information about additional car phone numbers.");
            }
            if (IsAssistantCount > 1)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {IsAssistantCount} assistant phone numbers, Outlook can have only 1 assistant phone number. Please update contact at Google, otherwise you will lose information about additional assistant phone numbers.");
            }
            if (IsCallbackCount > 1)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {IsCallbackCount} callback phone numbers, Outlook can have only 1 callback phone number. Please update contact at Google, otherwise you will lose information about additional callback phone numbers.");
            }
            if (IsRadioCount > 1)
            {
                Log.Debug($"Google contact \"{master.ToLogString()}\" has {IsRadioCount} radio phone numbers, Outlook can have only 1 radio phone number. Please update contact at Google, otherwise you will lose information about additional radio phone numbers.");
            }
            if (IsTtyCount > 1)
            {
                Log.Debug($"Google contact \"{master.ToLogString()}\" has {IsTtyCount} assistant TTY numbers, Outlook can have only 1 TTY phone number. Please update contact at Google, otherwise you will lose information about additional TTY phone numbers.");
            }
            if (IsCompanyCount > 1)
            {
                Log.Debug($"Google contact \"{master.ToLogString()}\" has {IsCompanyCount} company phone numbers, Outlook can have only 1 company phone number. Please update contact at Google, otherwise you will lose information about additional company phone numbers.");
            }

            foreach (var phone in master.PhoneNumbers)
            {
                if (string.IsNullOrWhiteSpace(phone.Type))
                {
                    if (IsHomeCount < 2)
                    {
                        phone.Type = HOME;
                        IsHomeCount++;
                    }
                    else if (IsWorkCount < 2)
                    {
                        phone.Type = WORK;
                        IsWorkCount++;
                    }
                    else if (IsMobileCount < 1)
                    {
                        phone.Type = PHONE_MOBILE;
                        IsMobileCount++;
                    }
                    else if (IsOtherCount < 1)
                    {
                        phone.Type = OTHER;
                        IsOtherCount++;
                    }
                    else if (IsCarCount < 1)
                    {
                        phone.Type = CAR;
                        IsCarCount++;
                    }
                }
                SetPhoneNumber(phone, slave);
            }
        }
        #endregion

        #region emails

        internal static void SetEmails(Person master, Outlook.ContactItem slave)
        {
            if (master.EmailAddresses == null)
                master.EmailAddresses = new List<EmailAddress>();
            if (master.EmailAddresses.Count > 3)
            {
                Log.Warning($"Google contact \"{master.ToLogString()}\" has {master.EmailAddresses.Count} emails, Outlook can have only 3 emails. Please update contact at Google, otherwise you will lose information about additional emails.");
            }
            if (master.EmailAddresses.Count > 0)
            {
                //only sync, if Email changed
                if (slave.Email1Address != master.EmailAddresses[0].Value)
                {
                    slave.Email1Address = master.EmailAddresses[0].Value;
                }

                //do not synchronize Gmail outlook label with Outlook email display name as
                //in Google label typically has values like "Home", "Work" or "Other", 
                //but in Outlook email display name is typically set to Full Name 
                /*
                if (!string.IsNullOrEmpty(master.EmailAddresses[0].Label) && slave.Email1DisplayName != master.EmailAddresses[0].Label)
                {//Don't set it to null, because some Handys leave it empty and then Outlook automatically sets (overwrites it)
                    slave.Email1DisplayName = master.EmailAddresses[0].Label; //Unfortunatelly this doesn't work when the email is changes also, because Outlook automatically sets it to default value when the email is changed ==> Call this function again after the first save of Outlook
                }
                */
            }
            else
            {
                slave.Email1Address = string.Empty;
                slave.Email1DisplayName = string.Empty;
            }

            if (master.EmailAddresses.Count > 1)
            {
                //only sync, if Email changed
                if (slave.Email2Address != master.EmailAddresses[1].Value)
                {
                    slave.Email2Address = master.EmailAddresses[1].Value;
                }

                //do not synchronize Gmail outlook label with Outlook email display name as
                //in Google label typically has values like "Home", "Work" or "Other", 
                //but in Outlook email display name is typically set to Full Name 
                /*
                if (!string.IsNullOrEmpty(master.EmailAddresses[1].Label) && slave.Email2DisplayName != master.EmailAddresses[1].Label)
                {//Don't set it to null, because some Handys leave it empty and then Outlook automatically sets (overwrites it)
                    slave.Email2DisplayName = master.EmailAddresses[1].Label; //Unfortunatelly this doesn't work when the email is changes also, because Outlook automatically sets it to default value when the email is changed ==> Call this function again after the first save of Outlook
                }
                */
            }
            else
            {
                slave.Email2Address = string.Empty;
                slave.Email2DisplayName = string.Empty;
            }

            if (master.EmailAddresses.Count > 2)
            {
                //only sync, if Email changed
                if (slave.Email3Address != master.EmailAddresses[2].Value)
                {
                    slave.Email3Address = master.EmailAddresses[2].Value;
                }

                //do not synchronize Gmail outlook label with Outlook email display name as
                //in Google label typically has values like "Home", "Work" or "Other", 
                //but in Outlook email display name is typically set to Full Name 
                /*
                if (!string.IsNullOrEmpty(master.EmailAddresses[2].Label) && slave.Email3DisplayName != master.EmailAddresses[2].Label)
                {//Don't set it to null, because some Handys leave it empty and then Outlook automatically sets (overwrites it)
                    slave.Email3DisplayName = master.EmailAddresses[2].Label; //Unfortunatelly this doesn't work when the email is changes also, because Outlook automatically sets it to default value when the email is changed ==> Call this function again after the first save of Outlook
                }
                */
            }
            else
            {
                slave.Email3Address = string.Empty;
                slave.Email3DisplayName = string.Empty;
            }
        }

        internal static void SetEmails(Outlook.ContactItem master, Person slave)
        {
            var email1 = ContactPropertiesUtils.GetOutlookEmailAddress1(master);
            var email2 = ContactPropertiesUtils.GetOutlookEmailAddress2(master);
            var email3 = ContactPropertiesUtils.GetOutlookEmailAddress3(master);

            var e1 = !string.IsNullOrWhiteSpace(email1);
            var e2 = !string.IsNullOrWhiteSpace(email2);
            var e3 = !string.IsNullOrWhiteSpace(email3);

            if (!e1 && e2)
            {
                email1 = email2;
                e1 = true;
                email2 = string.Empty;
                e2 = false;
            }

            if (!e1 && e3)
            {
                email1 = email3;
                e1 = true;
                email3 = string.Empty;
                e3 = false;
            }

            if (!e2 && e3)
            {
                email2 = email3;
                e2 = true;
                email3 = string.Empty;
                e3 = false;
            }

            if (e1)
            {
                AddEmail(slave, 0, email1, WORK);
            }
            else
            {
                RemoveEmail(slave, 0);
            }

            if (e2)
            {
                AddEmail(slave, 1, email2, HOME);
            }
            else
            {
                RemoveEmail(slave, 1);
            }

            if (e3)
            {
                AddEmail(slave, 2, email3, OTHER);
            }
            else
            {
                RemoveEmail(slave, 2);
            }
        }

        private static void RemoveEmail(Person slave, int index)
        {
            if (slave.EmailAddresses == null)
                slave.EmailAddresses = new List<EmailAddress>();
            if (slave.EmailAddresses.Count > 3)
            {
                Log.Warning($"Google contact \"{slave.ToLogString()}\" has {slave.EmailAddresses.Count} emails, Outlook can have only 3 emails. Additional Google emails will be deleted.");
            }
            //clear all emails above index
            for (var i = slave.EmailAddresses.Count - 1; i >= index; i--)
            {
                slave.EmailAddresses.RemoveAt(i);
            };
        }

        private static void AddEmail(Person slave, int index, string email, string type)
        {
            if (slave.EmailAddresses == null)
                slave.EmailAddresses = new List<EmailAddress>();
            if (slave.EmailAddresses.Count > index)
            {
                if (slave.EmailAddresses[index].Value != email)
                {
                    slave.EmailAddresses[index].Value = email;
                }
            }
            else
            {
                var e = new EmailAddress()
                {
                    Metadata = new FieldMetadata { Primary = slave.EmailAddresses.Count == 0 },
                    Value = email,
                    Type = type
                };
                slave.EmailAddresses.Add(e);
            }
        }

        #endregion

        public static void SetImClients(Outlook.ContactItem source, Person destination)
        {
            if (destination.ImClients == null)
                destination.ImClients = new List<ImClient>();
            else
                destination.ImClients.Clear();

            if (!string.IsNullOrEmpty(source.IMAddress))
            {
                //IMAddress are expected to be in form of ([Protocol]: [Address]; [Protocol]: [Address])
                var ImClientsRaw = source.IMAddress.Split(';');
                foreach (var imRaw in ImClientsRaw)
                {
                    var imDetails = imRaw.Trim().Split(':');
                    var im = new ImClient();
                    if (imDetails.Length == 1)
                    {
                        im.Username = imDetails[0].Trim();
                    }
                    else
                    {
                        im.Protocol = imDetails[0].Trim();
                        im.Username = imDetails[1].Trim();
                    }

                    //Only add the im Address if not empty (to avoid Google exception "address" empty)
                    if (!string.IsNullOrEmpty(im.Username))
                    {
                        if (im.Metadata == null)
                            im.Metadata = new FieldMetadata { Primary = destination.ImClients.Count == 0 };
                        im.Type = HOME;
                        destination.ImClients.Add(im);
                    }
                }
            }
        }

        public static void SetCompanies(Outlook.ContactItem source, Person destination)
        {
            if (destination.Organizations == null)
                destination.Organizations = new List<Organization>();
            else
                destination.Organizations.Clear();

            if (!string.IsNullOrEmpty(source.Companies))
            {
                //todo (obelix30)   test this....
                //Companies are expected to be in form of "[Company]; [Company]".
                var companiesRaw = source.Companies.Split(';');
                foreach (var companyRaw in companiesRaw)
                {
                    var company = new Organization
                    {
                        Name = (destination.Organizations.Count == 0) ? source.CompanyName : null,
                        Title = (destination.Organizations.Count == 0) ? source.JobTitle : null,
                        Department = (destination.Organizations.Count == 0) ? source.Department : null,
                        Metadata = new FieldMetadata { Primary = destination.Organizations.Count == 0 },
                        Type = WORK
                    };
                    destination.Organizations.Add(company);
                }
            }

            if (destination.Organizations.Count == 0 && (!string.IsNullOrEmpty(source.CompanyName) || !string.IsNullOrEmpty(source.JobTitle) || !string.IsNullOrEmpty(source.Department)))
            {
                var company = new Organization
                {
                    Name = source.CompanyName,
                    Title = source.JobTitle,
                    Department = source.Department,
                    Metadata = new FieldMetadata { Primary = true },
                    Type = WORK
                };
                destination.Organizations.Add(company);
            }
        }

        /// <summary>
        /// Updates Google contact from Outlook (but without groups/categories)
        /// </summary>
	    public static void UpdateContact(Outlook.ContactItem master, Person slave, bool useFileAs)
        {
            #region FileAs            
            if (useFileAs)
            {
                var fileAs = ContactPropertiesUtils.GetGoogleFileAs(slave);
                if (fileAs == null)
                {
                    fileAs = new FileAs();
                    if (slave.FileAses == null)
                        slave.FileAses = new List<FileAs>();
                    slave.FileAses.Add(fileAs);
                }
                if (!string.IsNullOrEmpty(master.FileAs))
                {
                    fileAs.Value = master.FileAs;
                }
                else if (!string.IsNullOrEmpty(master.FullName))
                {
                    fileAs.Value = master.FullName;
                }
                else if (!string.IsNullOrEmpty(master.CompanyName))
                {
                    fileAs.Value = master.CompanyName;
                }
                else if (!string.IsNullOrEmpty(master.Email1Address))
                {
                    fileAs.Value = master.Email1Address;
                }
            }

            #endregion FileAs

            #region Name
            var name = ContactPropertiesUtils.GetGooglePrimaryName(slave);
            if (name == null)
            {
                name = new Name(); // { Metadata = new FieldMetadata() { Source = new Source() { Type = "CONTACT" } } }; //ToDo: Check
                if (slave.Names == null)
                    slave.Names = new List<Name>();
                slave.Names.Add(name);
            }
            name.HonorificPrefix = master.Title;
            name.GivenName = master.FirstName;
            name.MiddleName = master.MiddleName;
            name.FamilyName = master.LastName;
            name.HonorificSuffix = master.Suffix;

            //Use the Google's full name to save a unique identifier. When saving the FullName, it always overwrites the Google Title
            if (!string.IsNullOrEmpty(master.FullName)) //Only if master.FullName has a value, i.e. not only a company or email contact
            {
                if (useFileAs)
                {
                    name.UnstructuredName = master.FileAs;
                }
                else
                {
                    name.UnstructuredName = OutlookContactInfo.GetTitleFirstLastAndSuffix(master);
                    if (!string.IsNullOrEmpty(name.UnstructuredName))
                    {
                        name.UnstructuredName = name.UnstructuredName.Trim().Replace("  ", " ");
                    }
                }
            }
            #endregion Name

            #region Birthday
            try
            {
                if (slave.Birthdays != null)
                    slave.Birthdays.Clear();
                else
                    slave.Birthdays = new List<Birthday>();

                if (!master.Birthday.Equals(outlookDateNone))
                {
                    slave.Birthdays.Add(new Birthday()
                    {
                        Date = new Date()
                        {
                            Day = master.Birthday.Day,
                            Month = master.Birthday.Month,
                            Year = master.Birthday.Year
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                Log.Error($"Birthday couldn't be updated from Outlook to Google for '{master.ToLogString()}': {ex.Message}");
            }
            #endregion Birthday

            var nickname = ContactPropertiesUtils.GetGoogleNickName(slave);
            if (nickname == null && !string.IsNullOrEmpty(master.NickName))
            {
                nickname = new Nickname() { Type = ContactPropertiesUtils.DEFAULT }; //ToDo: Find proper enum
                if (slave.Nicknames == null)
                    slave.Nicknames = new List<Nickname>();
                slave.Nicknames.Add(nickname);
                nickname.Value = master.NickName;
            }
            else if (nickname != null)
                nickname.Value = master.NickName;


            var location = ContactPropertiesUtils.GetGoogleOfficeLocation(slave);
            if (location == null && !string.IsNullOrEmpty(master.OfficeLocation))
            {
                location = new Location() { Type = DESK }; //ToDo: Find proper enum
                if (slave.Locations == null)
                    slave.Locations = new List<Location>();
                slave.Locations.Add(location);
                location.Value = master.OfficeLocation;
            }
            else if (location != null)
                location.Value = master.OfficeLocation;


            //Categories are synced separately in Syncronizer.OverwriteContactGroups: slave.Categories = master.Categories;
            var initials = ContactPropertiesUtils.GetGoogleInitials(slave);
            if (initials == null && !string.IsNullOrEmpty(master.Initials))
            {
                initials = new Nickname() { Type = ContactPropertiesUtils.INITIALS }; //ToDo: Find proper enum
                if (slave.Nicknames == null)
                    slave.Nicknames = new List<Nickname>();
                slave.Nicknames.Add(initials);
                initials.Value = master.Initials;
            }
            else if (initials != null)
                initials.Value = master.Initials;

            var language = ContactPropertiesUtils.GetGoogleUserDefined(slave, "language"); //ToDo: Create an enum or const for language
            if (language == null && !string.IsNullOrEmpty(master.Language))
            {
                language = new UserDefined() { Key = "language" }; //ToDo: Find proper enum
                if (slave.UserDefined == null)
                    slave.UserDefined = new List<UserDefined>();
                slave.UserDefined.Add(language);
                language.Value = master.Language;
            }
            else if (language != null)
                language.Value = master.Language;

            SetEmails(master, slave);

            SetAddresses(master, slave);

            SetPhoneNumbers(master, slave);

            SetCompanies(master, slave);

            SetImClients(master, slave);

            #region anniversary
            if (slave.Events == null)
                slave.Events = new List<Event>();
            //First remove anniversary
            foreach (var ev in slave.Events)
            {
                if (!string.IsNullOrEmpty(ev.Type) && ev.Type.Equals(EVENT_ANNIVERSARY))
                {
                    slave.Events.Remove(ev);
                    break;
                }
            }
            try
            {
                //Then add it again if existing
                if (!master.Anniversary.Equals(outlookDateNone)) //earlier also || master.Birthday.Year < 1900
                {
                    var ev = new Event
                    {
                        Type = EVENT_ANNIVERSARY,
                        Date = new Date
                        {
                            Day = master.Anniversary.Date.Day,
                            Month = master.Anniversary.Date.Month,
                            Year = master.Anniversary.Date.Year
                        }
                    };
                    slave.Events.Add(ev);
                }
            }
            catch (Exception ex)
            {
                Log.Error($"Anniversary couldn't be updated from Outlook to Google for '{master.ToLogString()}': {ex.Message}");
            }
            #endregion anniversary

            #region relations (spouse, child, manager and assistant)
            if (slave.Relations == null)
                slave.Relations = new List<Relation>();
            //First remove spouse, child, manager and assistant
            for (var i = slave.Relations.Count - 1; i >= 0; i--)
            {
                var rel = slave.Relations[i];
                if (rel.Type != null && (rel.Type.Equals(REL_SPOUSE) || rel.Type.Equals(REL_CHILD) || rel.Type.Equals(REL_MANAGER) || rel.Type.Equals(REL_ASSISTANT)))
                {
                    slave.Relations.RemoveAt(i);
                }
            }
            //Then add spouse again if existing
            if (!string.IsNullOrEmpty(master.Spouse))
            {
                var rel = new Relation
                {
                    Type = REL_SPOUSE,
                    Person = master.Spouse
                };
                slave.Relations.Add(rel);
            }
            //Then add children again if existing
            if (!string.IsNullOrEmpty(master.Children))
            {
                var rel = new Relation
                {
                    Type = REL_CHILD,
                    Person = master.Children
                };
                slave.Relations.Add(rel);
            }
            //Then add manager again if existing
            if (!string.IsNullOrEmpty(master.ManagerName))
            {
                var rel = new Relation
                {
                    Type = REL_MANAGER,
                    Person = master.ManagerName
                };
                slave.Relations.Add(rel);
            }
            //Then add assistant again if existing
            if (!string.IsNullOrEmpty(master.AssistantName))
            {
                var rel = new Relation
                {
                    Type = REL_ASSISTANT,
                    Person = master.AssistantName
                };
                slave.Relations.Add(rel);
            }
            #endregion relations (spouse, child, manager and assistant)

            #region HomePage
            if (slave.Urls == null)
                slave.Urls = new List<Url>();
            else
                slave.Urls.Clear();
            //Just copy the first URL, because Outlook only has 1
            if (!string.IsNullOrEmpty(master.WebPage))
            {
                var url = new Url
                {
                    Value = master.WebPage,
                    Type = URL_HOMEPAGE,
                    Metadata = new FieldMetadata { Primary = true }
                };
                slave.Urls.Add(url);
            }
            #endregion HomePage

            //CH - Fixed error with invalid xml being sent to google... This may need to be added to everything
            //slave.Content = $"<![CDATA[{master.Body}]]>";
            //floriwan: Maybe better to just escape the XML instead of putting it in CDATA, because this causes a CDATA added to all my contacts
            var bio = ContactPropertiesUtils.GetGoogleBiography(slave);
            if (bio == null && !string.IsNullOrEmpty(master.Body))
            {
                if (slave.Biographies == null)
                    slave.Biographies = new List<Biography>();
                bio = new Biography();
                slave.Biographies.Add(bio);
                bio.Value = System.Web.HttpUtility.HtmlEncode(master.Body);
            }
            else if (bio != null)
            {
                if (string.IsNullOrEmpty(master.Body))
                    bio.Value = string.Empty;
                else
                    bio.Value = System.Web.HttpUtility.HtmlEncode(master.Body);
            }

        }

        private enum OutlookFileAsFormat { CannotDetect, CompanyAndFullName, Company, FullNameAndCompany, LastNameAndFirstName, FirstMiddleLastSuffix };

        /// <summary>
        /// Updates Outlook contact from Google (but without groups/categories)
        /// </summary>
		public static void UpdateContact(Person master, Outlook.ContactItem slave, bool useFileAs)
        {
            if (master == null)
                throw new Exception("Google Person to update into Outlook is null");
            if (slave == null)
                throw new Exception("Outlook Contact to update from Google is null");


            #region DetectOutlookFileAsFormat

            var fmt = OutlookFileAsFormat.CannotDetect;

            if (!string.IsNullOrEmpty(slave.FileAs))
            {
                if (slave.CompanyAndFullName == slave.FileAs)
                {
                    fmt = OutlookFileAsFormat.CompanyAndFullName;
                }
                else if (slave.FullNameAndCompany == slave.FileAs)
                {
                    fmt = OutlookFileAsFormat.FullNameAndCompany;
                }
                else if (slave.CompanyName == slave.FileAs)
                {
                    fmt = OutlookFileAsFormat.Company;
                }
                else if (slave.LastNameAndFirstName == slave.FileAs)
                {
                    fmt = OutlookFileAsFormat.LastNameAndFirstName;
                }
                else if (slave.Subject == slave.FileAs)
                {
                    fmt = OutlookFileAsFormat.FirstMiddleLastSuffix;
                }
            }
            #endregion DetectOutlookFileAsFormat

            var name = ContactPropertiesUtils.GetGooglePrimaryName(master);
            var fileAs = ContactPropertiesUtils.GetGoogleFileAsValue(master);
            var org = ContactPropertiesUtils.GetGooglePrimaryOrganizationName(master);
            var email = ContactPropertiesUtils.GetGooglePrimaryEmailValue(master);

            #region Name                  
            slave.Title = (name == null) ? string.Empty : name.HonorificPrefix;
            slave.FirstName = (name == null) ? string.Empty : name.GivenName;
            slave.MiddleName = (name == null) ? string.Empty : name.MiddleName;
            slave.LastName = (name == null) ? string.Empty : name.FamilyName;
            slave.Suffix = (name == null) ? string.Empty : name.HonorificSuffix;
            if (string.IsNullOrEmpty(slave.FullName)) //The Outlook fullName is automatically set, so don't assign it from Google, unless the structured properties were empty
            {
                slave.FullName = (name == null) ? string.Empty : name.UnstructuredName;
            }

            #endregion Name

            #region Title/FileAs

            if (fmt == OutlookFileAsFormat.CompanyAndFullName)
            {
                var s = string.Empty;

                if (!string.IsNullOrEmpty(org))
                {
                    s = org;
                }

                if (name != null && !string.IsNullOrEmpty(name.FamilyName))
                {
                    s = !string.IsNullOrEmpty(s) ? s + "\r\n" + name.FamilyName : name.FamilyName;
                }

                if (name != null && !string.IsNullOrEmpty(name.GivenName))
                {
                    s = !string.IsNullOrEmpty(s)
                        ? !string.IsNullOrEmpty(name.FamilyName) ? s + ", " + name.GivenName : s + "\r\n" + name.GivenName
                        : name.GivenName;
                }

                if (name != null && !string.IsNullOrEmpty(name.MiddleName))
                {
                    if (!string.IsNullOrEmpty(s))
                    {
                        if (!string.IsNullOrEmpty(name.GivenName))
                        {
                            s = s + " " + name.MiddleName;
                        }
                        else if (!string.IsNullOrEmpty(name.FamilyName))
                        {
                            s = s + " " + name.MiddleName;
                        }
                    }
                    else
                    {
                        s = name.MiddleName;
                    }
                }

                slave.FileAs = s;
            }
            else if (fmt == OutlookFileAsFormat.Company)
            {
                if (!string.IsNullOrEmpty(org))
                {
                    slave.FileAs = org;
                }
            }
            else if (fmt == OutlookFileAsFormat.FullNameAndCompany)
            {
                var s = string.Empty;

                if (name != null && !string.IsNullOrEmpty(name.FamilyName))
                {
                    s = name.FamilyName;
                }

                if (name != null && !string.IsNullOrEmpty(name.GivenName))
                {
                    s = !string.IsNullOrEmpty(s) ? s + ", " + name.GivenName : name.GivenName;
                }

                if (name != null && !string.IsNullOrEmpty(name.MiddleName))
                {
                    if (!string.IsNullOrEmpty(s))
                    {
                        if (!string.IsNullOrEmpty(name.GivenName))
                        {
                            s = s + " " + name.MiddleName;
                        }
                        else if (!string.IsNullOrEmpty(name.FamilyName))
                        {
                            s = s + " " + name.MiddleName;
                        }
                    }
                    else
                    {
                        s = name.MiddleName;
                    }
                }

                if (!string.IsNullOrEmpty(org))
                {
                    s = !string.IsNullOrEmpty(s) ? s + "\r\n" + org : org;
                }

                slave.FileAs = s;
            }
            else if (fmt == OutlookFileAsFormat.LastNameAndFirstName)
            {
                var s = string.Empty;

                if (name != null && !string.IsNullOrEmpty(name.FamilyName))
                {
                    s = name.FamilyName;
                }

                if (name != null && !string.IsNullOrEmpty(name.GivenName))
                {
                    s = !string.IsNullOrEmpty(s) ? s + ", " + name.GivenName : name.GivenName;
                }

                if (name != null && !string.IsNullOrEmpty(name.MiddleName))
                {
                    if (!string.IsNullOrEmpty(s))
                    {
                        if (!string.IsNullOrEmpty(name.GivenName))
                        {
                            s = s + " " + name.MiddleName;
                        }
                        else if (!string.IsNullOrEmpty(name.FamilyName))
                        {
                            s = s + " " + name.MiddleName;
                        }
                    }
                    else
                    {
                        s = name.MiddleName;
                    }
                }

                slave.FileAs = s;
            }
            else if (fmt == OutlookFileAsFormat.FirstMiddleLastSuffix)
            {
                var s = string.Empty;

                if (name != null && !string.IsNullOrEmpty(name.GivenName))
                {
                    s = name.GivenName;
                }

                if (name != null && !string.IsNullOrEmpty(name.MiddleName))
                {
                    s = !string.IsNullOrEmpty(s) ? s + " " + name.MiddleName : name.MiddleName;
                }

                if (name != null && !string.IsNullOrEmpty(name.FamilyName))
                {
                    s = !string.IsNullOrEmpty(s) ? s + " " + name.FamilyName : name.FamilyName;
                }

                slave.FileAs = s;
            }
            else
            {
                if (string.IsNullOrEmpty(slave.FileAs) || useFileAs)
                {
                    if (name != null && !string.IsNullOrEmpty(name.UnstructuredName))
                    {
                        slave.FileAs = name.UnstructuredName.Replace("\r\n", "\n").Replace("\n", "\r\n"); //Replace twice to not replace a \r\n by \r\r\n. This is necessary because \r\n are saved as \n only to google and \r\n is saved on Outlook side to separate the single parts of the FullName
                    }
                    else if (!string.IsNullOrEmpty(fileAs))
                    {
                        slave.FileAs = fileAs.Replace("\r\n", "\n").Replace("\n", "\r\n"); //Replace twice to not replace a \r\n by \r\r\n. This is necessary because \r\n are saved as \n only to google and \r\n is saved on Outlook side to separate the single parts of the FullName
                    }
                    else if (!string.IsNullOrEmpty(org))
                    {
                        slave.FileAs = org;
                    }
                    else if (!string.IsNullOrEmpty(email))
                    {
                        slave.FileAs = email;
                    }
                }
                if (string.IsNullOrEmpty(slave.FileAs))
                {
                    if (!string.IsNullOrEmpty(slave.Email1Address))
                    {
                        var emailAddress = ContactPropertiesUtils.GetOutlookEmailAddress1(slave);
                        Log.Warning($"Google Person '{ContactPropertiesUtils.GetGoogleUniqueIdentifierName(master)}' has neither name nor email address. Setting email address of Outlook contact: {emailAddress}");
                        if (master.EmailAddresses == null)
                            master.EmailAddresses = new List<EmailAddress>();
                        master.EmailAddresses.Add(new EmailAddress() { Value = emailAddress });
                        slave.FileAs = emailAddress;
                        master.ToDebugLog();
                    }
                    else
                    {
                        Log.Error($"Google Person '{ContactPropertiesUtils.GetGoogleUniqueIdentifierName(master)}' has neither name nor email address. Cannot merge with Outlook contact: {slave.FileAs}");
                        master.ToDebugLog();
                        return;
                    }
                }
            }
            #endregion Title/FileAs

            #region birthday
            var birthday = ContactPropertiesUtils.GetGoogleBirthday(master);
            try
            {
                if (birthday != null && birthday.Date != null)
                {
                    var bd = new DateTime(birthday.Date.Year ?? DateTime.Now.Year, birthday.Date.Month ?? DateTime.Now.Month, birthday.Date.Day ?? DateTime.Now.Day);
                    if (bd != DateTime.MinValue)
                    {
                        if (!birthday.Date.Equals(slave.Birthday.Date)) //Only update if not already equal to avoid recreating the calendar item again and again
                        {
                            slave.Birthday = bd.Date;
                        }
                    }
                    else
                    {
                        slave.Birthday = outlookDateNone;
                    }
                }
                else
                {
                    slave.Birthday = outlookDateNone;
                }
            }
            catch (Exception ex)
            {
                Log.Error($"Birthday ({birthday}) couldn't be updated from Google to Outlook for '{slave.ToLogString()}': {ex.Message}");
            }
            #endregion birthday

            slave.NickName = ContactPropertiesUtils.GetGoogleNickNameValue(master);

            slave.Initials = ContactPropertiesUtils.GetGoogleInitialsValue(master);

            slave.OfficeLocation = ContactPropertiesUtils.GetGoogleOfficeLocationValue(master);
            //Categories are synced separately in Syncronizer.OverwriteContactGroups: slave.Categories = master.Categories;

            var languages = ContactPropertiesUtils.GetGoogleUserDefined(master, "language"); //ToDo: Find proper enum
            slave.Language = languages == null ? String.Empty : languages.Value;

            SetEmails(master, slave);

            SetPhoneNumbers(master, slave);

            SetAddresses(master, slave);

            #region companies
            slave.Companies = string.Empty;
            slave.CompanyName = string.Empty;
            slave.JobTitle = string.Empty;
            slave.Department = string.Empty;
            if (master.Organizations != null)
            {
                foreach (var company in master.Organizations)
                {
                    if (string.IsNullOrEmpty(company.Name) && string.IsNullOrEmpty(company.Title) && string.IsNullOrEmpty(company.Department))
                    {
                        continue;
                    }

                    if (company.Metadata != null && (company.Metadata.Primary ?? false) || company.Equals(master.Organizations[0]))
                    {//Per default copy the first company, but if there is a primary existing, use the primary
                        slave.CompanyName = company.Name;
                        slave.JobTitle = company.Title;
                        slave.Department = company.Department;
                    }
                    if (!string.IsNullOrEmpty(slave.Companies))
                    {
                        slave.Companies += "; ";
                    }

                    slave.Companies += company.Name;
                }
            }
            #endregion companies

            #region IM
            slave.IMAddress = string.Empty;
            if (master.ImClients != null)
                foreach (var im in master.ImClients)
                {
                    if (!string.IsNullOrEmpty(slave.IMAddress))
                    {
                        slave.IMAddress += "; ";
                    }

                    if (!string.IsNullOrEmpty(im.Protocol) && !im.Protocol.Trim().Equals("None", StringComparison.InvariantCultureIgnoreCase))
                    {
                        slave.IMAddress += im.Protocol + ": " + im.Username;
                    }
                    else
                    {
                        slave.IMAddress += im.Username;
                    }
                }
            #endregion IM

            #region anniversary
            var found = false;
            try
            {
                if (master.Events != null)
                    foreach (var ev in master.Events)
                    {
                        if (ev.Type != null && ev.Type.Equals(EVENT_ANNIVERSARY))
                        {
                            var an = new DateTime(ev.Date.Year ?? DateTime.Now.Year, ev.Date.Month ?? DateTime.Now.Month, ev.Date.Day ?? DateTime.Now.Day);
                            if (!an.Equals(slave.Anniversary.Date)) //Only update if not already equal to avoid recreating the calendar item again and again
                            {
                                slave.Anniversary = an;
                            }

                            found = true;
                            break;
                        }
                    }
                if (!found)
                {
                    slave.Anniversary = outlookDateNone; //set to empty in the end
                }
            }
            catch (Exception ex)
            {
                Log.Error($"Anniversary couldn't be updated from Google to Outlook for '{slave.ToLogString()}': {ex.Message}");
            }
            #endregion anniversary

            #region relations (spouse, child, manager, assistant)
            slave.Children = string.Empty;
            slave.Spouse = string.Empty;
            slave.ManagerName = string.Empty;
            slave.AssistantName = string.Empty;
            if (master.Relations != null)
                foreach (var rel in master.Relations)
                {
                    if (rel.Type != null && rel.Type.Equals(REL_CHILD))
                    {
                        slave.Children = rel.Person;
                    }
                    else if (rel.Type != null && rel.Type.Equals(REL_SPOUSE))
                    {
                        slave.Spouse = rel.Person;
                    }
                    else if (rel.Type != null && rel.Type.Equals(REL_MANAGER))
                    {
                        slave.ManagerName = rel.Person;
                    }
                    else if (rel.Type != null && rel.Type.Equals(REL_ASSISTANT))
                    {
                        slave.AssistantName = rel.Person;
                    }
                }
            #endregion relations (spouse, child, manager, assistant)

            slave.WebPage = string.Empty;
            if (master.Urls != null)
                foreach (var website in master.Urls)
                {
                    if (website != null && (website.Type == null || website.Type.Equals(URL_HOMEPAGE)) && (website.Metadata != null && (website.Metadata.Primary ?? false) || website.Equals(master.Urls[0])))
                    {//Per default copy the first website, but if there is a primary existing, use the primary
                        slave.WebPage = website.Value;
                    }
                }

            var bio = ContactPropertiesUtils.GetGoogleBiographyValue(master);

            try
            {
                var nonRTF = string.Empty;

                // RTFBody was introduced in later version of Outlook
                // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
                try
                {
                    if (slave.Body != null && slave.RTFBody != null)
                    {
                        nonRTF = Utilities.ConvertToText(slave.RTFBody as byte[]);
                    }
                }
                catch (AccessViolationException)
                {
                }



                //only update, if plain text is different between master and slave
                if (!nonRTF.Equals(System.Web.HttpUtility.HtmlDecode(bio)))
                {
                    //only update, if RTF text is same as plain text
                    if (string.IsNullOrEmpty(nonRTF) || nonRTF.Equals(slave.Body))
                    {
                        slave.Body = System.Web.HttpUtility.HtmlDecode(bio);
                    }
                    else
                    {
                        if (Synchronizer.SyncContactsForceRTF)
                        {
                            slave.Body = System.Web.HttpUtility.HtmlDecode(bio);
                        }
                        else
                        {
                            Log.Warning($"Outlook contact notes body not updated, because it is RTF, otherwise it will overwrite it by plain text: {slave.ToLogString()}");
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Log.Debug(e, $"Error when converting RTF to plain text, updating Google Person '{ContactPropertiesUtils.GetGoogleUniqueIdentifierName(master)}' notes to Outlook without RTF check: {e.Message}");
                slave.Body = bio;
            }
        }
    }
}

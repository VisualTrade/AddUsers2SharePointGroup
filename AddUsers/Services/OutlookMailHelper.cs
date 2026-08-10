using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace AddUsers.Services
{
    /// <summary>
    /// Resolves the recipients of an Outlook mail item to SMTP addresses.
    /// Handles Exchange users, plain SMTP entries, and distribution lists
    /// (expanded recursively with a depth guard), falling back to the
    /// PR_SMTP_ADDRESS MAPI property when the object model has no answer.
    /// </summary>
    internal static class OutlookMailHelper
    {
        private const string PrSmtpAddress =
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        private const string PrSenderSmtpAddress =
            "http://schemas.microsoft.com/mapi/proptag/0x5D01001F";

        private const int MaxDistributionListDepth = 8;

        /// <summary>
        /// Returns the first selected MailItem in the active explorer, or null.
        /// The caller owns the returned COM object and must release it.
        /// </summary>
        public static Outlook.MailItem GetSelectedMailItem(Outlook.Application application)
        {
            int ignored;
            return GetSelectedMailItem(application, out ignored);
        }

        /// <summary>
        /// Returns the first selected MailItem and reports how many items are selected,
        /// so callers can tell the user when additional selected messages are ignored.
        /// </summary>
        public static Outlook.MailItem GetSelectedMailItem(Outlook.Application application, out int selectionCount)
        {
            selectionCount = 0;
            if (application == null)
            {
                return null;
            }

            Outlook.Explorer explorer = null;
            Outlook.Selection selection = null;
            try
            {
                explorer = application.ActiveExplorer();
                if (explorer == null)
                {
                    return null;
                }

                selection = explorer.Selection;
                if (selection == null || selection.Count == 0)
                {
                    return null;
                }

                selectionCount = selection.Count;
                object item = selection[1]; // Outlook collections are 1-based.
                var mail = item as Outlook.MailItem;
                if (mail == null && item != null)
                {
                    Marshal.ReleaseComObject(item);
                }

                return mail;
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                Release(selection);
                Release(explorer);
            }
        }

        /// <summary>
        /// Resolves the sender (when <paramref name="includeSender"/> is true) and
        /// every recipient of <paramref name="mail"/> to distinct SMTP addresses.
        /// Participants that cannot be resolved are reported by display name
        /// through <paramref name="unresolved"/>.
        /// </summary>
        public static List<string> ResolveRecipientSmtpAddresses(
            Outlook.MailItem mail,
            bool includeSender,
            out List<string> unresolved)
        {
            if (mail == null)
            {
                throw new ArgumentNullException(nameof(mail));
            }

            var resolved = new List<string>();
            var seenAddresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var unresolvedNames = new List<string>();
            var seenUnresolved = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (includeSender)
            {
                ResolveSender(mail, resolved, seenAddresses, unresolvedNames, seenUnresolved);
            }

            Outlook.Recipients recipients = null;
            try
            {
                recipients = mail.Recipients;
                int count = recipients.Count;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Recipient recipient = null;
                    Outlook.AddressEntry entry = null;
                    try
                    {
                        recipient = recipients[i];

                        // On items in Sent Items (and drafts) the Recipients collection still
                        // contains BCC entries; adding those to a group whose membership is
                        // visible to every site member would disclose them. Skip BCC entirely.
                        int recipientType;
                        try
                        {
                            recipientType = recipient.Type;
                        }
                        catch (COMException)
                        {
                            recipientType = (int)Outlook.OlMailRecipientType.olTo;
                        }
                        if (recipientType == (int)Outlook.OlMailRecipientType.olBCC)
                        {
                            continue;
                        }

                        string displayName = SafeGetRecipientName(recipient);
                        try
                        {
                            entry = recipient.AddressEntry;
                        }
                        catch (COMException)
                        {
                            entry = null;
                        }

                        if (entry != null)
                        {
                            ResolveAddressEntry(
                                entry, displayName, 0,
                                resolved, seenAddresses,
                                unresolvedNames, seenUnresolved);
                        }
                        else
                        {
                            // No address entry at all: the raw address is the last resort.
                            string address = SafeGetRecipientAddress(recipient);
                            if (LooksLikeSmtp(address))
                            {
                                AddResolved(address, resolved, seenAddresses);
                            }
                            else
                            {
                                AddUnresolved(displayName ?? address, unresolvedNames, seenUnresolved);
                            }
                        }
                    }
                    finally
                    {
                        Release(entry);
                        Release(recipient);
                    }
                }
            }
            finally
            {
                Release(recipients);
            }

            unresolved = unresolvedNames;
            return resolved;
        }

        private static void ResolveAddressEntry(
            Outlook.AddressEntry entry,
            string displayName,
            int depth,
            List<string> resolved,
            HashSet<string> seenAddresses,
            List<string> unresolvedNames,
            HashSet<string> seenUnresolved)
        {
            string name = displayName ?? SafeGetEntryName(entry);

            Outlook.OlAddressEntryUserType userType;
            try
            {
                userType = entry.AddressEntryUserType;
            }
            catch (COMException)
            {
                userType = Outlook.OlAddressEntryUserType.olOtherAddressEntry;
            }

            switch (userType)
            {
                case Outlook.OlAddressEntryUserType.olSmtpAddressEntry:
                    {
                        string address = SafeGetEntryAddress(entry);
                        if (!LooksLikeSmtp(address))
                        {
                            address = GetSmtpViaPropertyAccessor(entry);
                        }

                        CommitResult(address, name, resolved, seenAddresses, unresolvedNames, seenUnresolved);
                        break;
                    }

                case Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry:
                case Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry:
                    {
                        string address = GetExchangeUserSmtp(entry);
                        if (!LooksLikeSmtp(address))
                        {
                            address = GetSmtpViaPropertyAccessor(entry);
                        }

                        CommitResult(address, name, resolved, seenAddresses, unresolvedNames, seenUnresolved);
                        break;
                    }

                case Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry:
                    {
                        if (!ExpandExchangeDistributionList(
                                entry, depth, resolved, seenAddresses, unresolvedNames, seenUnresolved))
                        {
                            AddUnresolved(name, unresolvedNames, seenUnresolved);
                        }

                        break;
                    }

                case Outlook.OlAddressEntryUserType.olOutlookDistributionListAddressEntry:
                    {
                        if (!ExpandMembers(
                                entry, depth, resolved, seenAddresses, unresolvedNames, seenUnresolved))
                        {
                            AddUnresolved(name, unresolvedNames, seenUnresolved);
                        }

                        break;
                    }

                default:
                    {
                        string address = GetSmtpViaPropertyAccessor(entry);
                        if (!LooksLikeSmtp(address))
                        {
                            address = SafeGetEntryAddress(entry);
                        }

                        CommitResult(address, name, resolved, seenAddresses, unresolvedNames, seenUnresolved);
                        break;
                    }
            }
        }

        private static bool ExpandExchangeDistributionList(
            Outlook.AddressEntry entry,
            int depth,
            List<string> resolved,
            HashSet<string> seenAddresses,
            List<string> unresolvedNames,
            HashSet<string> seenUnresolved)
        {
            if (depth >= MaxDistributionListDepth)
            {
                return false;
            }

            Outlook.ExchangeDistributionList list = null;
            Outlook.AddressEntries members = null;
            try
            {
                list = entry.GetExchangeDistributionList();
                if (list == null)
                {
                    return false;
                }

                members = list.GetExchangeDistributionListMembers();
                if (members == null)
                {
                    return false;
                }

                EnumerateMembers(members, depth, resolved, seenAddresses, unresolvedNames, seenUnresolved);
                return true;
            }
            catch (COMException)
            {
                return false;
            }
            finally
            {
                Release(members);
                Release(list);
            }
        }

        private static bool ExpandMembers(
            Outlook.AddressEntry entry,
            int depth,
            List<string> resolved,
            HashSet<string> seenAddresses,
            List<string> unresolvedNames,
            HashSet<string> seenUnresolved)
        {
            if (depth >= MaxDistributionListDepth)
            {
                return false;
            }

            Outlook.AddressEntries members = null;
            try
            {
                members = entry.Members;
                if (members == null)
                {
                    return false;
                }

                EnumerateMembers(members, depth, resolved, seenAddresses, unresolvedNames, seenUnresolved);
                return true;
            }
            catch (COMException)
            {
                return false;
            }
            finally
            {
                Release(members);
            }
        }

        private static void EnumerateMembers(
            Outlook.AddressEntries members,
            int depth,
            List<string> resolved,
            HashSet<string> seenAddresses,
            List<string> unresolvedNames,
            HashSet<string> seenUnresolved)
        {
            int count = members.Count;
            for (int i = 1; i <= count; i++)
            {
                Outlook.AddressEntry member = null;
                try
                {
                    member = members[i];
                    if (member != null)
                    {
                        ResolveAddressEntry(
                            member, null, depth + 1,
                            resolved, seenAddresses,
                            unresolvedNames, seenUnresolved);
                    }
                }
                catch (COMException)
                {
                    // Skip members that fail to load; the rest are still valuable.
                }
                finally
                {
                    Release(member);
                }
            }
        }

        /// <summary>
        /// Resolves the sender through the same ladder as recipients: SMTP-type fast
        /// path, AddressEntry resolution (Exchange users etc.), then the
        /// PR_SENDER_SMTP_ADDRESS property on the item itself.
        /// </summary>
        private static void ResolveSender(
            Outlook.MailItem mail,
            List<string> resolved,
            HashSet<string> seenAddresses,
            List<string> unresolvedNames,
            HashSet<string> seenUnresolved)
        {
            // Unsent items (drafts, Outbox) have no sender yet; do not report one missing.
            try
            {
                if (!mail.Sent)
                {
                    return;
                }
            }
            catch (COMException)
            {
            }

            string senderName = null;
            try
            {
                senderName = mail.SenderName;
            }
            catch (COMException)
            {
            }

            string senderType = null;
            try
            {
                senderType = mail.SenderEmailType;
            }
            catch (COMException)
            {
            }

            if (string.Equals(senderType, "SMTP", StringComparison.OrdinalIgnoreCase))
            {
                string address = null;
                try
                {
                    address = mail.SenderEmailAddress;
                }
                catch (COMException)
                {
                }

                if (LooksLikeSmtp(address))
                {
                    AddResolved(address, resolved, seenAddresses);
                    return;
                }
            }

            Outlook.AddressEntry sender = null;
            try
            {
                try
                {
                    sender = mail.Sender;
                }
                catch (COMException)
                {
                }

                if (sender != null)
                {
                    Outlook.OlAddressEntryUserType senderEntryType;
                    try
                    {
                        senderEntryType = sender.AddressEntryUserType;
                    }
                    catch (COMException)
                    {
                        senderEntryType = Outlook.OlAddressEntryUserType.olOtherAddressEntry;
                    }

                    // A distribution-list sender is never expanded into its membership
                    // (one click would mass-add people unrelated to the conversation);
                    // only the list's own address, via the item property below, is used.
                    bool senderIsList =
                        senderEntryType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry ||
                        senderEntryType == Outlook.OlAddressEntryUserType.olOutlookDistributionListAddressEntry;

                    if (!senderIsList)
                    {
                        // Fall through to the item-level property when entry resolution
                        // yields nothing (e.g. a departed employee whose directory
                        // object is gone) instead of stopping at "unresolved".
                        int resolvedBefore = resolved.Count;
                        int unresolvedBefore = unresolvedNames.Count;
                        ResolveAddressEntry(
                            sender, senderName, 0,
                            resolved, seenAddresses,
                            unresolvedNames, seenUnresolved);
                        if (resolved.Count > resolvedBefore)
                        {
                            return;
                        }
                        while (unresolvedNames.Count > unresolvedBefore)
                        {
                            seenUnresolved.Remove(unresolvedNames[unresolvedNames.Count - 1]);
                            unresolvedNames.RemoveAt(unresolvedNames.Count - 1);
                        }
                    }
                }
            }
            finally
            {
                Release(sender);
            }

            string fallback = GetMailItemSmtpProperty(mail, PrSenderSmtpAddress);
            if (LooksLikeSmtp(fallback))
            {
                AddResolved(fallback, resolved, seenAddresses);
            }
            else
            {
                AddUnresolved(
                    string.IsNullOrWhiteSpace(senderName) ? "(unknown sender)" : senderName + " (sender)",
                    unresolvedNames, seenUnresolved);
            }
        }

        private static string GetMailItemSmtpProperty(Outlook.MailItem mail, string schema)
        {
            Outlook.PropertyAccessor accessor = null;
            try
            {
                accessor = mail.PropertyAccessor;
                return accessor.GetProperty(schema) as string;
            }
            catch (COMException)
            {
                return null;
            }
            catch (InvalidCastException)
            {
                return null;
            }
            finally
            {
                Release(accessor);
            }
        }

        private static string GetExchangeUserSmtp(Outlook.AddressEntry entry)
        {
            Outlook.ExchangeUser user = null;
            try
            {
                user = entry.GetExchangeUser();
                return user != null ? user.PrimarySmtpAddress : null;
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                Release(user);
            }
        }

        private static string GetSmtpViaPropertyAccessor(Outlook.AddressEntry entry)
        {
            Outlook.PropertyAccessor accessor = null;
            try
            {
                accessor = entry.PropertyAccessor;
                return accessor.GetProperty(PrSmtpAddress) as string;
            }
            catch (COMException)
            {
                return null;
            }
            catch (InvalidCastException)
            {
                return null;
            }
            finally
            {
                Release(accessor);
            }
        }

        private static void CommitResult(
            string address,
            string fallbackName,
            List<string> resolved,
            HashSet<string> seenAddresses,
            List<string> unresolvedNames,
            HashSet<string> seenUnresolved)
        {
            if (LooksLikeSmtp(address))
            {
                AddResolved(address, resolved, seenAddresses);
            }
            else
            {
                AddUnresolved(fallbackName, unresolvedNames, seenUnresolved);
            }
        }

        private static void AddResolved(string address, List<string> resolved, HashSet<string> seen)
        {
            string trimmed = address.Trim();
            if (seen.Add(trimmed))
            {
                resolved.Add(trimmed);
            }
        }

        private static void AddUnresolved(string name, List<string> unresolved, HashSet<string> seen)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                name = "(unknown recipient)";
            }

            if (seen.Add(name))
            {
                unresolved.Add(name);
            }
        }

        private static bool LooksLikeSmtp(string address)
        {
            return !string.IsNullOrWhiteSpace(address) && address.IndexOf('@') > 0;
        }

        private static string SafeGetRecipientName(Outlook.Recipient recipient)
        {
            try
            {
                return recipient.Name;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static string SafeGetRecipientAddress(Outlook.Recipient recipient)
        {
            try
            {
                return recipient.Address;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static string SafeGetEntryName(Outlook.AddressEntry entry)
        {
            try
            {
                return entry.Name;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static string SafeGetEntryAddress(Outlook.AddressEntry entry)
        {
            try
            {
                return entry.Address;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static void Release(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
    }
}

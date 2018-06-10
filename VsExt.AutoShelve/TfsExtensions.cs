using Microsoft.TeamFoundation.VersionControl.Client;
using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace VsExt.AutoShelve
{
    public static class TfsExtensions
    {
        public static bool DifferFrom(this PendingChange[] currentChanges, PendingChange[] shelvedChanges)
        {
            var shelveditems = shelvedChanges.ToDictionary(c => c.ServerItem);
            return currentChanges.OrderByDescending(GetLastChangeDate)
                .Any(currentItem => !shelveditems.TryGetValue(currentItem.ServerItem, out var shelvedItem)
                    || !shelvedItem.UploadHashValue.SequenceEqual(GetHashValue(currentItem)));
        }

        private static DateTime GetLastChangeDate(PendingChange change)
        {
            return change.IsDelete
                ? change.CreationDate
                : File.GetLastWriteTime(change.LocalItem);
        }

        private static readonly Lazy<HashAlgorithm> DefaultHashAlgorithm = new Lazy<HashAlgorithm>(MD5.Create);

        private static byte[] GetHashValue(PendingChange change)
        {
            return change.IsDelete || !File.Exists(change.LocalItem)
                ? change.UploadHashValue
                : DefaultHashAlgorithm.Value.ComputeHash(File.ReadAllBytes(change.LocalItem));
        }

        public static string GetDomain(this string name)
        {
            var stop = name.IndexOf("\\");
            return (stop > -1) ? name.Substring(0, stop) : string.Empty;
        }

        public static string GetLogin(this string name)
        {
            var stop = name.IndexOf("\\");
            return name.Substring(stop + 1);
        }

        public static bool IsNameSpecificToWorkspace(this string shelvesetName)
        {
            return shelvesetName.Contains("{0}");
        }

        public static bool IsNameSpecificToDate(this string shelvesetName)
        {
            return shelvesetName.Contains("{2}");
        }

        private const int NameLength = 64;

        private static readonly char[] NameBadCharacters = new[]
        {
            '/',
            ':',
            '<',
            '>',
            '\\',
            '|',
            '*',
            '?',
            ';',
        };

        public static string CleanShelvesetName(this string shelvesetName)
        {
            if (string.IsNullOrWhiteSpace(shelvesetName))
            {
                return Resources.DefaultShelvetsetName;
            }
            var cleanName = new StringBuilder(NameLength);
            var len = 0;
            foreach (var ch in shelvesetName)
            {
                if (NameBadCharacters.Contains(ch))
                {
                    continue;
                }
                cleanName.Append(ch);
                if (++len >= NameLength)
                {
                    break;
                }
            }
            return cleanName.ToString();
        }
    }
}

// Guids.cs
// MUST match guids.h
using System;

namespace LatishSehgal.ExtensionSync
{
    static class GuidList
    {
        public const string guidExtensionSyncPkgString = "887b3e11-719d-4912-94b2-f88168027643";
        public const string guidExtensionSyncCmdSetString = "7e357f67-d5a0-42bb-a105-a0e772476a5f";

        public static readonly Guid guidExtensionSyncCmdSet = new Guid(guidExtensionSyncCmdSetString);
    };
}
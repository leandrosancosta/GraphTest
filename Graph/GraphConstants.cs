// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <GraphConstantsSnippet>
namespace GraphTutorial
{
    public static class GraphConstants
    {
        // Escopos de leitura
        public readonly static string[] Scopes =
        {
            "User.Read",
            "MailboxSettings.Read",
            "Calendars.ReadWrite"
        };
    }
}
// </GraphConstantsSnippet>

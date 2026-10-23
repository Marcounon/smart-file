/* Smart File - configuration.
 * Fill in clientId and tenantId after registering the Entra app (see README, step 1).
 * Everything else is optional tuning.
 */
window.SmartFileConfig = {
  // Microsoft Entra app registration -> Overview page
  clientId: "d77dd0c7-a190-436a-b899-deacfcb0a504",
  tenantId: "6341f2fd-2752-4b39-8f62-d89fb04b231a",

  // Delegated Microsoft Graph permissions the add-in asks for.
  // Mail.ReadWrite covers listing folders, reading earlier mail, moving messages and creating folders.
  scopes: ["User.Read", "Mail.ReadWrite"],

  // Folders that are never offered as a filing target.
  // Well-known names are resolved by Graph, so this works in any mailbox language.
  excludedWellKnownFolders: [
    "inbox", "drafts", "sentitems", "deleteditems", "junkemail", "outbox", "archive",
    "conversationhistory", "syncissues", "recoverableitemsdeletions", "clutter",
    "scheduled", "searchfolders", "serverfailures", "localfailures", "conflicts"
  ],
  // Extra top-level folders to skip by display name (case-insensitive).
  excludedFolderNames: [
    "RSS Feeds", "RSS Subscriptions", "Notes", "Journal", "Tasks", "Calendar", "Contacts",
    "Quick Step Settings", "Yammer Root", "Files", "Snoozed", "Unwanted", "ExternalContacts"
  ],

  // Display name of the folder that new folders are created under by default (e.g. "Customers").
  // null = use the parent of the top suggestion if it has one, otherwise the mailbox root.
  newFolderDefaultParent: null,

  // Sender domains whose company name does not appear in the domain itself.
  // Add your own: "ab.com": "Allen-Bradley"
  domainAliases: {
    "cat.com": "Caterpillar"
  },

  maxFolderDepth: 6,        // how deep to walk the folder tree
  historySampleSize: 50,    // earlier emails to inspect per signal
  bodySampleChars: 6000,    // how much of the body to scan for folder names
  maxLearnedRules: 150,     // sender -> folder choices remembered in roaming settings

  // Scoring weights. Each signal adds to a folder's score; the highest score is proposed first.
  weights: {
    learnedSender: 100,     // you filed mail from this exact address here before
    learnedDomain: 60,      // you filed mail from this domain here before
    conversation: 60,       // other emails in this conversation live here
    senderHistory: 50,      // earlier emails from this address live here (scaled by share)
    domainHistory: 35,      // earlier emails from this domain live here
    orgName: 45,            // folder name matches the sender's company (from the domain)
    orgPrefix: 15,          // weak prefix match for very short domains (cat.com -> Caterpillar)
    senderName: 30,         // folder name matches a word in the sender's display name
    recipientOrg: 25,       // folder name matches another external company on the thread
    subjectToken: 18, subjectCap: 45,
    bodyToken: 6,     bodyCap: 18,
    pathToken: 8,     pathCap: 16,
    confidentThreshold: 25  // below this the prompt says "best guess" instead of a confident match
  },

  publicEmailDomains: [
    "gmail.com", "googlemail.com", "outlook.com", "hotmail.com", "live.com", "msn.com",
    "yahoo.com", "icloud.com", "me.com", "aol.com", "proton.me", "protonmail.com",
    "comcast.net", "att.net", "verizon.net", "bellsouth.net", "sbcglobal.net", "charter.net"
  ]
};

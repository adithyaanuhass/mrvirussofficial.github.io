/*
  CloudNest configuration
  ----------------------------------------------------------------
  Paste your OAuth Client IDs here (they are public identifiers, NOT secrets,
  so it is fine to commit this file to GitHub).

  You can also skip editing this file and paste the IDs inside the app:
  header -> gear icon -> Settings. (Saved in your browser only.)

  How to get the IDs: see README.md
*/
window.CLOUDNEST_CONFIG = {
  // Google Cloud Console -> APIs & Services -> Credentials -> OAuth client ID (Web application)
  GOOGLE_CLIENT_ID: "",

  // Azure Portal -> App registrations -> Application (client) ID
  MICROSOFT_CLIENT_ID: "",

  // Advanced (leave as is unless you know what you are changing)
  GOOGLE_SCOPES: "https://www.googleapis.com/auth/drive",
  MICROSOFT_SCOPES: ["User.Read", "Files.ReadWrite.All"],
  MICROSOFT_AUTHORITY: "https://login.microsoftonline.com/common"
};

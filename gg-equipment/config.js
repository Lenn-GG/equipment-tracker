// G&G Equipment Requests: site settings.
// Fill these in once (see README.md, steps 1 and 2). None of these values are secret.
window.GG_CONFIG = {
  // Microsoft Entra ID > App registrations > your app > Overview
  clientId: "PASTE-APPLICATION-CLIENT-ID",
  tenantId: "PASTE-DIRECTORY-TENANT-ID",

  // The SharePoint site that holds the request lists. Everyone who uses the app
  // needs Edit (Member) access to this site.
  sharePointSite: "https://ggdoor.sharepoint.com/sites/Operations",

  // These people are always superintendents and can change roles in Settings.
  admins: ["lenn@ggdoor.net"],

  // SharePoint list names. The app creates them on first run.
  requestsList: "EquipmentRequests",
  configList: "EquipmentConfig",

  // How often open screens check for changes, in seconds.
  refreshSeconds: 20
};

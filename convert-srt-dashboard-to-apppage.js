// Converts the SRT dashboard Site Page into a single-part APP PAGE (no Edit
// button / no page chrome, edge-to-edge). Copy from the editor, paste into the
// browser DevTools console while logged in to the hpen-poc-manager site.
//
// It: (1) creates a new single-part app page hosting the SRT Dashboard web part,
// (2) renames the current regular page to ...-REG.aspx (kept as a backup),
// (3) renames the new app page into the canonical SRT-Resource-Dashboard.aspx.
// Watch for 200/201/204 statuses and a DONE line, then reload the dashboard.
(async () => {
  var site = "https://hpe.sharepoint.com/teams/hpen-poc-manager";
  var webPartId = "9fda1c5c-5ab2-476b-a6ab-5be927e950ec"; // SRT Dashboard
  var canonical = "SRT-Resource-Dashboard.aspx";
  var backup = "SRT-Resource-Dashboard-REG.aspx";
  var title = "SRT Resource Dashboard";
  var iid = crypto.randomUUID();

  var ctx = await (await fetch(site + "/_api/contextinfo",
    { method: "POST", headers: { Accept: "application/json;odata=nometadata" } })).json();
  var H = {
    Accept: "application/json;odata=nometadata",
    "Content-Type": "application/json;odata=nometadata",
    "X-RequestDigest": ctx.FormDigestValue
  };

  // 1. create the single-part app page (auto-named)
  var page = await (await fetch(site + "/_api/sitepages/pages",
    { method: "POST", headers: H, body: JSON.stringify({ PageLayoutType: "SingleWebPartAppPage", Title: title }) })).json();
  console.log("created app page:", page.Id, page.FileName);

  // 2. checkout + set the SRT Dashboard web part full-width
  await fetch(site + "/_api/sitepages/pages(" + page.Id + ")/checkoutpage", { method: "POST", headers: H }).catch(function () {});
  var canvas = [
    { position: { controlIndex: 1, sectionIndex: 1, sectionFactor: 12, layoutIndex: 1, zoneIndex: 1 },
      controlType: 3, id: iid, webPartId: webPartId,
      webPartData: { id: webPartId, instanceId: iid, title: title, description: title, dataVersion: "1.0", properties: { description: title } } },
    { controlType: 0, pageSettingsSlice: { isDefaultDescription: true, isDefaultThumbnail: true } }
  ];
  var sd = await fetch(site + "/_api/sitepages/pages(" + page.Id + ")/SavePageAsDraft",
    { method: "POST", headers: H, body: JSON.stringify({ CanvasContent1: JSON.stringify(canvas), Title: title }) });
  console.log("savedraft", sd.status);
  var pub = await fetch(site + "/_api/sitepages/pages(" + page.Id + ")/Publish", { method: "POST", headers: H });
  console.log("publish", pub.status);

  // 3. back up the current regular page, then promote the app page into its name
  var base = "/teams/hpen-poc-manager/SitePages/";
  var mv1 = await fetch(site + "/_api/web/getfilebyserverrelativeurl('" + base + canonical + "')/moveto(newurl='" + base + backup + "',flags=1)",
    { method: "POST", headers: H });
  console.log("backup current →", backup, ":", mv1.status);
  var mv2 = await fetch(site + "/_api/web/getfilebyserverrelativeurl('" + base + page.FileName + "')/moveto(newurl='" + base + canonical + "',flags=1)",
    { method: "POST", headers: H });
  console.log("promote app page →", canonical, ":", mv2.status);
  console.log("DONE — reload " + site + "/SitePages/" + canonical + " (no Edit button; full width). Backup at " + backup + ".");
})();

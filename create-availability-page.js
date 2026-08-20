// Paste this whole file into the browser DevTools Console while logged in
// to any page on https://hpe.sharepoint.com/teams/hpen-poc-manager
// Creates the single-part app page for the SSE Availability web part and
// renames it to SSE-Availability.aspx (the URL the dashboard pill expects).
(async () => {
  var s = "https://hpe.sharepoint.com/teams/hpen-poc-manager";
  var wp = "7e1b9a4c-3d52-4a6f-b8c1-2f9e0d5a7361"; // SseAvailability web part
  var fn = "SSE-Availability.aspx";
  var t = "SSE Availability";
  var iid = crypto.randomUUID();

  var c = await (await fetch(s + "/_api/contextinfo", {
    method: "POST",
    headers: { Accept: "application/json;odata=nometadata" }
  })).json();
  var H = {
    Accept: "application/json;odata=nometadata",
    "Content-Type": "application/json;odata=nometadata",
    "X-RequestDigest": c.FormDigestValue
  };

  var p = await (await fetch(s + "/_api/sitepages/pages", {
    method: "POST",
    headers: H,
    body: JSON.stringify({ PageLayoutType: "SingleWebPartAppPage", Title: t })
  })).json();
  console.log("created", p.Id, p.FileName);

  await fetch(s + "/_api/sitepages/pages(" + p.Id + ")/checkoutpage", {
    method: "POST", headers: H
  }).catch(function () {});

  var cv = [
    {
      position: { controlIndex: 1, sectionIndex: 1, sectionFactor: 12, layoutIndex: 1, zoneIndex: 1 },
      controlType: 3,
      id: iid,
      webPartId: wp,
      webPartData: {
        id: wp, instanceId: iid, title: t, description: t,
        dataVersion: "1.0", properties: { description: t }
      }
    },
    { controlType: 0, pageSettingsSlice: { isDefaultDescription: true, isDefaultThumbnail: true } }
  ];

  var sd = await fetch(s + "/_api/sitepages/pages(" + p.Id + ")/SavePageAsDraft", {
    method: "POST", headers: H,
    body: JSON.stringify({ CanvasContent1: JSON.stringify(cv), Title: t })
  });
  console.log("savedraft", sd.status);

  var pb = await fetch(s + "/_api/sitepages/pages(" + p.Id + ")/Publish", {
    method: "POST", headers: H
  });
  console.log("publish", pb.status);

  var ou = "/teams/hpen-poc-manager/SitePages/" + p.FileName;
  var nu = "/teams/hpen-poc-manager/SitePages/" + fn;
  var mv = await fetch(
    s + "/_api/web/getfilebyserverrelativeurl('" + ou + "')/moveto(newurl='" + nu + "',flags=1)",
    { method: "POST", headers: H }
  );
  console.log("rename", mv.status);
  console.log("DONE", s + "/SitePages/" + fn);
})();

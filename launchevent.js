/* Fluctuum File: Smart Alert on send.
 * Classic Outlook on Windows loads this file directly (JS only runtime, no imports, no Office.onReady).
 * Logic: if the draft already has a project (custom property), send. Otherwise try to infer it
 * from a known thread or a project code in the subject. If nothing fits, ask the user.
 */
/* eslint-disable no-var */
var FCD_INTERNAL_DOMAIN = "fluctuum.com";

function fcdNormConv(id) {
  return id ? String(id).replace(/\//g, "-").replace(/\+/g, "_") : "";
}

function fcdAsk(event) {
  event.completed({
    allowEvent: false,
    errorMessage: "Which project does this email belong to? Choose one in Fluctuum File so it is filed in 07 Comms, or send it without a project.",
    errorMessageMarkdown: "**Which project does this email belong to?**\n\nChoose one in Fluctuum File so it is filed in 07 Comms after sending, or select **Send Anyway** to send it without a project.",
    cancelLabel: "Choose project",
    commandId: "msgComposeOpenPaneButton",
    sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser
  });
}

function fcdAllow(event) {
  event.completed({ allowEvent: true });
}

function fcdSetAndAllow(event, props, code, how) {
  props.set("project", code);
  props.set("inferred", how);
  props.saveAsync(function () { fcdAllow(event); });
}

function fcdAllInternal(recipients) {
  if (!FCD_INTERNAL_DOMAIN) return false;
  if (!recipients.length) return false;
  for (var i = 0; i < recipients.length; i++) {
    var a = (recipients[i].emailAddress || "").toLowerCase();
    if (a.slice(-FCD_INTERNAL_DOMAIN.length - 1) !== "@" + FCD_INTERNAL_DOMAIN) return false;
  }
  return true;
}

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;
  var done = false;
  // Never hold the user longer than ~4 s: if anything hangs, let the email go.
  var guard = setTimeout(function () { if (!done) { done = true; fcdAllow(event); } }, 4000);
  function finish(fn) { if (done) return; done = true; clearTimeout(guard); fn(); }

  try {
    item.loadCustomPropertiesAsync(function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) { finish(function () { fcdAllow(event); }); return; }
      var props = r.value;
      if (props.get("project")) { finish(function () { fcdAllow(event); }); return; }

      var rs = Office.context.roamingSettings;
      var codes = ((rs && rs.get("codes")) || "").split(",");
      var conv = (rs && rs.get("conv")) || {};

      // 1. Known thread
      var key = fcdNormConv(item.conversationId).slice(-40);
      if (key && conv[key]) { finish(function () { fcdSetAndAllow(event, props, conv[key], "thread"); }); return; }

      // 2. Project code in subject
      item.subject.getAsync(function (s) {
        var subject = (s && s.value) || "";
        var m = subject.match(/\b\d{6}\b/g) || [];
        for (var i = 0; i < m.length; i++) {
          if (codes.indexOf(m[i]) >= 0) {
            var code = m[i];
            finish(function () { fcdSetAndAllow(event, props, code, "subject"); });
            return;
          }
        }
        // 3. Internal only emails go out without asking
        item.to.getAsync(function (t) {
          item.cc.getAsync(function (c) {
            var all = ((t && t.value) || []).concat((c && c.value) || []);
            if (fcdAllInternal(all)) { finish(function () { fcdAllow(event); }); return; }
            finish(function () { fcdAsk(event); });
          });
        });
      });
    });
  } catch (e) {
    finish(function () { fcdAllow(event); });
  }
}

// IMPORTANT: map the manifest FunctionName to the JS function (required in every client).
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

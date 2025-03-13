Office.onReady(function(info) {
    console.log("✅ Office.js er klar!");

    if (info.host === Office.HostType.Outlook) {
        console.log("✅ Outlook registreret!");
    }
});

// Gør forwardEmail globalt tilgængelig for Outlook
window.forwardEmail = function(event) {
    if (!Office.context.mailbox) {
        console.error("❌ Mailbox API er ikke tilgængelig.");
        return;
    }

    console.log("📩 forwardEmail() kaldt!");

    Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["it-afdelingen@nordenergi.dk"],
        subject: "[Spam Check] Mistænkelig e-mail",
        body: "Denne e-mail er blevet markeret som mulig spam. Venligst undersøg den."
    });

    if (event && event.completed) {
        event.completed();
    }
};

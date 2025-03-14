Office.onReady(function(info) {
    console.log("✅ Office.js er klar! 123");

    if (info.host === Office.HostType.Outlook) {
        console.log("✅ Outlook registreret!");

        // Sørg for at forwardEmail er global
        window.forwardEmail = function(event) {
            console.log("📩 forwardEmail() kaldt!");

            if (!Office.context || !Office.context.mailbox) {
                console.error("❌ Mailbox API er ikke tilgængelig.");
                return;
            }

            Office.context.mailbox.displayNewMessageForm({
                toRecipients: ["it-afdelingen@nordenergi.dk"],
                subject: "[Spam Check] Mistænkelig e-mail",
                body: "Denne e-mail er blevet markeret som mulig spam. Venligst undersøg den."
            });

            if (event && event.completed) {
                event.completed();
            }
        };

        console.log("📌 forwardEmail funktion er nu registreret:", typeof window.forwardEmail);
    }
});

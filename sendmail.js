Office.onReady(function(info) {
    console.log("✅ Office.js er klar!");

    document.addEventListener("DOMContentLoaded", function() {
        console.log("✅ DOM er klar!");
        
        let button = document.getElementById("spamCheckButton");
        if (button) {
            button.addEventListener("click", forwardEmail);
            console.log("✅ Knap registreret!");
        } else {
            console.warn("⚠️ Kunne ikke finde knappen 'spamCheckButton'");
        }
    });
});

function forwardEmail(event) {
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

    if (event) {
        event.completed();
    }
}

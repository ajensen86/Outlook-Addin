Office.onReady(function(info) {
    console.log("✅ Office.js er klar!");

    if (Office.context.mailbox) {
        console.log("📬 Mailbox API er tilgængelig!");

        // Knapregistrering, hvis vi har en UI-knap
        let button = document.getElementById("spamCheckButton");
        if (button) {
            button.addEventListener("click", forwardEmail);
            console.log("✅ Knap registreret!");
        } else {
            console.warn("⚠️ Kunne ikke finde knappen 'spamCheckButton'");
        }
    } else {
        console.error("❌ Mailbox API er ikke tilgængelig! Script vil ikke fungere.");
    }
});

function forwardEmail(event) {
    if (!Office.context.mailbox || !Office.context.mailbox.item) {
        console.error("❌ Mailbox eller Item API er ikke tilgængelig.");
        return;
    }

    console.log("📩 forwardEmail() kaldt!");

    // Forsøg at videresende mail
    Office.context.mailbox.item.forwardAsync(
        {
            toRecipients: ["it-afdelingen@nordenergi.dk"],
            subject: "[Spam Check] Mistænkelig e-mail",
            body: "Denne e-mail er blevet markeret som mulig spam. Venligst undersøg den."
        },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("❌ Fejl ved videresendelse: ", asyncResult.error.message);
            } else {
                console.log("✅ Mail videresendt succesfuldt!");
            }
        }
    );

    if (event) {
        event.completed();
    }
}

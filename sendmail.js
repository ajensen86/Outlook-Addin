Office.onReady(function(info) {
    console.log("Office.js er klar!");
    
    // Nu kan vi initialisere funktioner eller UI
    document.getElementById("spamCheckButton").addEventListener("click", forwardEmail);
});



function forwardEmail(event) {
    if (!Office.context.mailbox) {
        console.error("Mailbox API er ikke tilgængelig.");
        return;
    }

    console.log("forwardEmail() kaldt!");

    Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["it-afdelingen@nordenergi.dk"],
        subject: "[Spam Check] Mistænkelig e-mail",
        body: "Denne e-mail er blevet markeret som mulig spam. Venligst undersøg den."
    });

    event.completed();
}

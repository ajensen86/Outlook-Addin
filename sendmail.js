function forwardEmail(event) {
    Office.context.mailbox.item.forwardAsync(
        {
            toRecipients: ["it-afdelingen@nordenergi.dk"],
            subject: "[Spam Check] Mistænkelig e-mail",
            body: "Denne e-mail er blevet markeret som mulig spam. Venligst undersøg den."
        },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Fejl ved videresendelse: ", asyncResult.error.message);
            }
        }
    );
    event.completed();
}

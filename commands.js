
// Office is ready
Office.onReady(() => {
    console.log("Office.js is ready for event-based activation");
});

// Event handler for OnNewMessageCompose
async function onNewMessageComposeHandler(event) {
    console.log("OnNewMessageCompose event triggered");
    
    try {
        const email = Office.context.mailbox.userProfile.emailAddress;
        console.log("Current Email:", email);

        // Fetch signature from API
        const response = await fetch(
            `https://piercingly-cavernous-laura.ngrok-free.dev/api/signature/get?email=${email}`,
            {
                method: "GET",
                headers: {
                    "Content-Type": "application/json",
                    "ngrok-skip-browser-warning": "true",
                },
            }
        );

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        console.log("Signature data received:", data);
        
        const signatureHTML = data.data || "<p>--<br>Your Default Signature</p>";

        // Use setSignatureAsync for automatic insertion (requires Mailbox 1.10)
        Office.context.mailbox.item.body.setSignatureAsync(
            signatureHTML,
            { coercionType: "html" },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Signature inserted automatically!");
                    event.completed(); // Signal event completion
                } else {
                    console.error("Error inserting signature:", asyncResult.error);
                    event.completed({ allowEvent: true }); // Allow compose to continue
                }
            }
        );

    } catch (error) {
        console.error("Error in onNewMessageComposeHandler:", error);
        event.completed({ allowEvent: true }); // Allow compose to continue even on error
    }
}

// Register the function
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);

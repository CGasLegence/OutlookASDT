function insertTextAutomatically() {
    const item = Office.context.mailbox.item;

    // Check if the item is in compose mode
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
        item.body.setSelectedDataAsync(
            "This is a test",
            { coercionType: Office.CoercionType.Text },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Error inserting text:", asyncResult.error.message);
                }
            }
        );
    }
}

// Ensure the function runs when the add-in is loaded
Office.onReady(() => {
    if (Office.context.mailbox.item) {
        insertTextAutomatically();
    }
});

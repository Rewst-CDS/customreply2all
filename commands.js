Office.onReady((info) => {
    // Initialization logic
});

function specialReplyToAll(event) {
    // 1. Get the current mail item
    const item = Office.context.mailbox.item;
    
    // 2. Your core logic goes here
    console.log("Special Reply initiated...");

    // 3. Trigger the reply form
    item.displayReplyAllForm({
        // Add your specific recipient shuffling logic here
    });

    // 4. IMPORTANT: Only call event.completed() if this was triggered 
    // by a Ribbon button (event exists).
    if (event) {
        event.completed();
    }
}
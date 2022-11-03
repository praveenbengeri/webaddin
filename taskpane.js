Office.onReady(() => {
    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
    Office.context.document.settings.saveAsync();
    console.log("AutoShowTaskpaneWithDocument has been set");
});


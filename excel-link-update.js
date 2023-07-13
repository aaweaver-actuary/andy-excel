Excel.run(function(context) {
    // Get active workbook
    const workbook = context.workbook;

    // Get all worksheets
    let worksheets = workbook.worksheets.load('items');

    // Synchronize the state with Excel
    return context.sync().then(function() {
        // Find all linked workbooks
        let linkedWorkbookPromises = [];
        worksheets.items.forEach(sheet => {
            let links = sheet.hyperlinks.load('items');
            linkedWorkbookPromises.push(context.sync().then(function() {
                return links.items.map(link => link.address);
            }));
        });
        return Promise.all(linkedWorkbookPromises);
    }).then(function(linkedWorkbooks) {
        // Flatten the array of arrays
        linkedWorkbooks = [].concat.apply([], linkedWorkbooks);

        let updatePromises = [];
        linkedWorkbooks.forEach(workbookLink => {
            let newWorkbookLink = workbookLink.replace('findText', 'replaceText'); // Change these as needed
            updatePromises.push(tryOpenAndReplaceLink(workbook, workbookLink, newWorkbookLink));
        });
        return Promise.all(updatePromises);
    }).then(function(results) {
        // Create a new worksheet for the results
        let resultSheet = workbook.worksheets.add('JSLinkUpdate');
        let resultRange = resultSheet.getRange('A1:C' + (results.length + 1));
        resultRange.values = [['Original Link', 'Updated Link', 'Result']].concat(results);
        resultSheet.activate();
    }).catch(function(error) {
        console.log('Error: ' + error);
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
});

function tryOpenAndReplaceLink(workbook, oldLink, newLink) {
    return new Promise(function(resolve, reject) {
        Excel.run(function(context) {
            // Try to open the new workbook
            let newWorkbook = context.workbook.application.workbooks.open(newLink);
            return context.sync().then(function() {
                // Change the old link to the new link
                let worksheets = workbook.worksheets.load('items');
                return context.sync().then(function() {
                    worksheets.items.forEach(sheet => {
                        let links = sheet.hyperlinks.load('items');
                        context.sync().then(function() {
                            links.items.forEach(link => {
                                if (link.address === oldLink) {
                                    link.address = newLink;
                                }
                            });
                        });
                    });
                    return 'Updated Successfully';
                });
            }).catch(function() {
                return 'Error Opening Workbook';
            });
        }).then(function(result) {
            resolve([oldLink, newLink, result]);
        }).catch(function(error) {
            resolve([oldLink, newLink, 'Error: ' + error]);
        });
    });
}

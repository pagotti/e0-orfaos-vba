Office.initialize = function (reason) {
    switch (reason.toString()) {
        case "inserted":
            console.log("O add-in acaba de ser inserido!");
            break;
        case "documentOpened":
            console.log("O add-in é parte de um documento que acaba de ser aberto.");
            break;
    }
    $(document).ready(initializeButton);
};
function initializeButton() {
    $('#runMacro').click(Macro);
}
function Macro() {
    Excel.run(function (context) {
        var range = context.workbook.getSelectedRange();
        range.values = [["Estou Funcionando!"]];
        return context.sync();
    })["catch"](function (error) {
        console.log('Error: ' + error);
        if (error instanceof OfficeExtension.Error) {
            console.log('Informações para Debug: ' + JSON.stringify(error.debugInfo));
        }
    });
}

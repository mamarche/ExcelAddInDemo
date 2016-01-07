/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#crea-tabella').click(creaTabella);
        });
    };

    function creaTabella() {

        //creo un oggetto tabella
        var miaTabella = new Office.TableData();
        //aggiungo le intestazioni di colonna
        miaTabella.headers = ['Nome', 'Cognome'];
        //Aggiungo un array di righe con i dati di test
        miaTabella.rows = [['Mario', 'Rossi'], ['Franco', 'Bianchi'], ['Luca', 'Verdi']];

        //utilizzo l'API per generare la tabella nel foglio
        Office.context.document.setSelectedDataAsync(miaTabella, { coercionType: Office.CoercionType.Table },
            function (result) {
                var error = result.error
                if (result.status === Office.AsyncResultStatus.Failed) {
                    $('#error-log').text(error.name + ": " + error.message);
                }
            });
    }
    
})();
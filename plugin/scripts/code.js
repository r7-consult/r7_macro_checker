

window.Asc.plugin.init = function()
{
    console.log('MacroTester');
    
    var ws = new WebSocket('ws://127.0.0.1:8000/ws');


    ws.onmessage = function(event) {
        console.log('event', event);

        Asc.scope.macro = event.data;
        window.Asc.plugin.callCommand(function() {

            let macro = Asc.scope.macro;
            console.log(`Eval: ${macro}`);

            try {
                eval(macro);
            }
            catch(error) {
                console.log(`Error ${error.name}: ${error.message}`);
                return `${error.name}: ${error.message}`;
            }

            return null;
            
        }, false, false, function(error) {

            if (error == null) {
                ws.send('ok');
            }
            else {
                ws.send(error)
            }

        });
    };
}


window.Asc.plugin.button = function(id, windowId) {

    if (windowId) {
        window.Asc.plugin.executeMethod('CloseWindow', [windowId]);
    }
    else {
        this.executeCommand('close', '');
    }
}

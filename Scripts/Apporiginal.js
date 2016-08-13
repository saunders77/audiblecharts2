// This function is run when the app is ready to start interacting with the host application
// It ensures the DOM is ready before adding click handlers to buttons
Office.initialize = function (reason) {
    $(document).ready(function () {
        $('#getDataBtn').click(function () { getData('#selectedDataTxt'); });

        // If setSelectedDataAsync method is supported by the host application 
        // setDatabtn is hooked up to call the method else setDatabtn is removed
        if (Office.context.document.setSelectedDataAsync) {
            $('#setDataBtn').click(function () { setData('#selectedDataTxt'); });
        }
        else {
            $('#setDataBtn').remove();
        }
    });
    if (Office.context) {
        Office.context.document.bindings.getByIdAsync('MyBinding6', function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                
            appendMessage('Addeda old binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
            writeDataFromBinding(asyncResult.value);
            //    asyncResult.value.addHandlerAsync(Office.EventType.BindingDataChanged, updateData);
            //    asyncResult.value.getDataAsync(function (asyncResult) {
            //        displayFilters(asyncResult.value.slice(0));
            //        updateSettings();
            //        prepareData(asyncResult.value);
            //    });
            doEvent();
            } else {
                appendMessage("No existing bindings");
            }
        });
    }
};

// Writes data from textbox to the current selection in the document
function setData(elementId) {
    Office.context.document.setSelectedDataAsync($(elementId).val());
}

// Reads the data from current selection of the document and displays it in a textbox
function getData(elementId) {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
    function (result) {
        if (result.status === 'succeeded') {
            $(elementId).val(result.value);
        }
    });
}

function myTest2(){
    appendMessage("<p> hello. how are you? " + document.getElementById('music').innerHTML) + "<br> fine </p>";
}

function createTable(){
    // Build table.
    var myTable = new Office.TableData();
    myTable.headers = [['C tone','C# tone','D tone','D# tone','E tone','F tone','F# tone','G tone','G# tone','A tone','A# tone','B tone','C tone']];
    myTable.rows = [['','','','','','','','','','','','','']];
    rowNum = 1;
    while(rowNum < 64){
        myTable.rows.push(['','','','','','','','','','','','','']);
        rowNum = rowNum + 1;
    }

    // Write table.
    Office.context.document.setSelectedDataAsync(myTable, {coercionType: "table"},
        function (result) {
            var error = result.error
            if (result.status === "failed") {
                write(error.name + ": " + error.message);
            }
    });
}

function setBinding(){
    
  /*  Office.context.document.bindings.addFromSelectionAsync("matrix", { id: 'myBinding' }, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            appendMessage('Action faileda. Error: ' + asyncResult.error.message);
        } else {
            appendMessage('Added new binding with typea: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    });*/
  
  Office.context.document.bindings.addFromSelectionAsync("table", { id: 'MyBinding6' }, 
        function (asyncResult) {
            appendMessage('Addeda newa binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
            writeDataFromBinding(asyncResult.value);
            doEvent();
        }
    );
  
}

function writeNotesData()
{
    appendMessage(readNoteMatrix()[0]);
}

function writeDataFromBinding(binding) {
    if (binding) {

        binding.getDataAsync(function (dataResult) {
            var myMatrix = dataResult.value;
            Office.context.document.settings.set('noteMatrix', myMatrix.rows);
        });
    }
    else{
        appendMessage("binding is false");
    }
}

function readNoteMatrix(){
    return Office.context.document.settings.get('noteMatrix');
}

function MyHandler(evt) {
    writeDataFromBinding(evt.binding);
}

// Retrieve the "MyCities" binding, and add an event handler that redraws
// the pins on the map when a value is changed in the binding.
function doEvent() {
    Office.select("bindings#MyBinding6").addHandlerAsync("bindingDataChanged", MyHandler);
}

function playSound(){
    appendMessage("Play Notea");
    Office.context.document.setSelectedDataAsync("=Row()",
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed"){
            appendMessage(error.name + ": " + error.message);
            }
            getMyRow();
        }
    );

}
function getMyRow()
{
    Office.context.document.getSelectedDataAsync("text",
            function (asyncResult) {
                var error = asyncResult.error;
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    appendMessage(error.name + ": " + error.message);
                } 
                else {
                    // Get selected data.
                    var dataValue = asyncResult.value; 
                    Office.context.document.setSelectedDataAsync(dataValue,
                    function (asyncResult) {
                        var error = asyncResult.error;
                        if (asyncResult.status === "failed"){
                            appendMessage(error.name + ": " + error.message);
                        }
                    });
                }           
            }
        );
}

function pausecomp(ms) {
 ms += new Date().getTime();
 while (new Date() < ms){}
 } 

function startSong(){
    // first flag that a song is playing
//    if(Office.context.document.settings.get("songIsPlaying") == 1){
//        Office.context.document.settings.set("songIsPlaying",0);
//    }
//    else{
    Office.context.document.settings.set("songIsPlaying",1);
//    }
    
    // continue to do the following while the song is playing:
    Office.context.document.settings.set("loopLength",64);
    Office.context.document.settings.set("time",0);
    
    if(Office.context.document.settings.get("songIsPlaying") == 1){
//      playNext();
        Office.context.document.settings.set("intervalID",setInterval(playNext,250));
    }
}

function startTimedSong(){
    Office.context.document.settings.set("songIsPlaying",1);
    Office.context.document.settings.set("loopLength",64);
    playTimedNext();
}

function playTimedNext(){
    if(Office.context.document.settings.get("songIsPlaying") == 1){
        var chord = [["","","","","","","","","","","","",""]];
        var myNoteMatrix = Office.context.document.settings.get("noteMatrix");
        var myLoopLength = Office.context.document.settings.get("loopLength");
        //var myLoopLength = 8;
        chord[0] = getChordAtTime(0,myNoteMatrix);
        var myTime = 1;
        while(myTime < myLoopLength){
            chord.push(getChordAtTime(myTime,myNoteMatrix));
            myTime = myTime + 1;
        }

    
        myTime = 0;
        var period = 400;
        setTimeout(appendMessage("one"),900);
        setTimeout(appendMessage("two"),900);
        setTimeout(appendMessage("three"),900);
        setTimeout(appendMessage("four"),1300);
        setTimeout(appendMessage("five"),1700);
        while(myTime < myLoopLength){
            setTimeout(playTimedChord(chord[myTime]),myTime*period);
            setTimeout(appendMessage(chord[myTime]),myTime*period);
            myTime = myTime + 1;
        }
    
        setTimeout(playTimedNext,myTime*period);
    }
    
}

function playTimedChord(chord){
    var mynoteIndex = 0;
    while(mynoteIndex < 13){
        appendMusic(chord[mynoteIndex]);
        mynoteIndex += 1;
    }
}

function getChordAtTime(myTime,myNotes){
    var myChord = ["","","","","","","","","","","","",""];
    if(myNotes[myTime][0] == 1){
        myChord[0] = ("C4");
    }
    if(myNotes[myTime][1] == 1){
        myChord[1] = ("Csharp4");
    }
    if(myNotes[myTime][2] == 1){
        myChord[2] = ("D4");
    }
    if(myNotes[myTime][3] == 1){
        myChord[3] = ("Dsharp4");
    }
    if(myNotes[myTime][4] == 1){
        myChord[4] = ("E4");
    }
    if(myNotes[myTime][5] == 1){
        myChord[5] = ("F4");
    }
    if(myNotes[myTime][6] == 1){
        myChord[6] = ("Fsharp4");
    }
    if(myNotes[myTime][7] == 1){
        myChord[7] = ("G4");
    }
    if(myNotes[myTime][8] == 1){
        myChord[8] = ("Gsharp4");
    }
    if(myNotes[myTime][9] == 1){
        myChord[9] = ("A4");
    }
    if(myNotes[myTime][10] == 1){
        myChord[10] = ("Asharp4");
    }
    if(myNotes[myTime][11] == 1){
        myChord[11] = ("B4");
    }
    if(myNotes[myTime][12] == 1){
        myChord[12] = ("C5");
    }
    //appendMessage(myChord);
    return myChord;
}

function playNext(){
    //clearMusic();
    var chord = ["","","","","","","","","","","","",""];
    var myTime = Office.context.document.settings.get("time");
    var myNoteMatrix = Office.context.document.settings.get("noteMatrix");
    if(myNoteMatrix[myTime][0] == 1){
        chord[0] = ("C4");
    }
    if(myNoteMatrix[myTime][1] == 1){
        chord[1] = ("Csharp4");
    }
    if(myNoteMatrix[myTime][2] == 1){
        chord[2] = ("D4");
    }
    if(myNoteMatrix[myTime][3] == 1){
        chord[3] = ("Dsharp4");
    }
    if(myNoteMatrix[myTime][4] == 1){
        chord[4] = ("E4");
    }
    if(myNoteMatrix[myTime][5] == 1){
        chord[5] = ("F4");
    }
    if(myNoteMatrix[myTime][6] == 1){
        chord[6] = ("Fsharp4");
    }
    if(myNoteMatrix[myTime][7] == 1){
        chord[7] = ("G4");
    }
    if(myNoteMatrix[myTime][8] == 1){
        chord[8] = ("Gsharp4");
    }
    if(myNoteMatrix[myTime][9] == 1){
        chord[9] = ("A4");
    }
    if(myNoteMatrix[myTime][10] == 1){
        chord[10] = ("Asharp4");
    }
    if(myNoteMatrix[myTime][11] == 1){
        chord[11] = ("B4");
    }
    if(myNoteMatrix[myTime][12] == 1){
        chord[12] = ("C5");
    }
    
    var noteIndex = 0;
    appendMessage(chord);
    while(noteIndex < 13){
        appendMusic(chord[noteIndex]);
        
        noteIndex += 1;
    }
    
    
    
    if(myTime < Office.context.document.settings.get("loopLength") - 1){
        Office.context.document.settings.set("time",myTime + 1);
    }
    else{
        Office.context.document.settings.set("time",0);
    }
    
    //if(Office.context.document.settings.get("songIsPlaying") == 1){
    //    setTimeout(playNext,400);
    //}
}
function stopSong(){
    Office.context.document.settings.set("songIsPlaying",0);
    clearInterval(Office.context.document.settings.get("intervalID"));
    
}





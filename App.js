// This function is run when the app is ready to start interacting with the host application
// It ensures the DOM is ready before adding click handlers to buttons
var audios = [];

var noteNames = [];
var myData = [];
var soundCounter = 0;

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
        
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, ChangedHandler);
    });
    /*
    if (Office.context) {
        Office.context.document.bindings.getByIdAsync('MyBinding6', function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                
            //appendMessageUI('Addeda old binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
            writeDataFromBinding(asyncResult.value);
            //    asyncResult.value.addHandlerAsync(Office.EventType.BindingDataChanged, updateData);
            //    asyncResult.value.getDataAsync(function (asyncResult) {
            //        displayFilters(asyncResult.value.slice(0));
            //        updateSettings();
            //        prepareData(asyncResult.value);
            //    });
            doEvent();
            } else {
                //appendMessageUI("No existing bindings");
            }
        });
    }
    */
    for(var c = 0;c < 3;c++)
    {
        preloadNote("C3");
        preloadNote("Csharp3");
        preloadNote("D3");
        preloadNote("Dsharp3");
        preloadNote("E3");
        preloadNote("F3");
        preloadNote("Fsharp3");
        preloadNote("G3");
        preloadNote("Gsharp3");
        preloadNote("A3");
        preloadNote("Asharp3");
        preloadNote("B3");
        preloadNote("C4");
        preloadNote("Csharp4");
        preloadNote("D4");
        preloadNote("Dsharp4");
        preloadNote("E4");
        preloadNote("F4");
        preloadNote("Fsharp4");
        preloadNote("G4");
        preloadNote("Gsharp4");
        preloadNote("A4");
        preloadNote("Asharp4");
        preloadNote("B4");
        preloadNote("C5");
        preloadNote("Csharp5");
        preloadNote("D5");
        preloadNote("Dsharp5");
        preloadNote("E5");
        preloadNote("F5");
        preloadNote("Fsharp5");
        preloadNote("G5");
        preloadNote("Gsharp5");
        preloadNote("A5");
        preloadNote("Asharp5");
        preloadNote("B5");
        preloadNote("C6");
    }


};

function showHelp()
{
	//appendMessage("read to showing help");
	document.getElementById("transparencyd").style.visibility = "visible";
	//appendMessage("showing transparencyd");
    
    //appendMessage(document.getElementById("helpd").id);
    
    document.getElementById("helpd").style.visibility = "visible";
}

function hideHelp()
{
	document.getElementById("transparencyd").style.visibility = "hidden";
    document.getElementById("helpd").style.visibility = "hidden";
}

function appendMessageUI(ourText)
{
    hideHelp();
    document.getElementById("errors").innerHTML = "<br>" + ourText;
    document.getElementById("transparencye").style.visibility = "visible";
    document.getElementById("errors").style.visibility = "visible";
}

function vanishErrors()
{
    document.getElementById("errors").style.visibility = "hidden";
    //appendMessageUI("errors hide");
    document.getElementById("transparencye").style.visibility = "hidden";
}

function refreshButton()
{
    /*
    if(document.getElementById("autoplay").checked == true)
    {
        document.getElementById("playSelected").style.visibility = "hidden";
    }
    else
    {
        document.getElementById("playSelected").style.visibility = "visible";
    }
    */
}

function ChangedHandler(eventArgs)
{   
    vanishErrors();
    if(document.getElementById("autoplay").checked == true && document.getElementById("playSelected").disabled == "")
    {
        soundCounter = soundCounter + 1;
        if(soundCounter >= 3)
        {
            document.getElementById("autoplaydiv").style.visibility = "visible";
        }
        
        playSelected();
    }
}

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

function playSelected()
{
    stopPlaying();
    //appendMessageUI("playing selected.");
    myData = [];
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,playList); 
}

function stopPlaying()
{
    
}

function playList(asyncResult)
{
    var error = asyncResult.error;
    var goodData = true;
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        if(error.message.indexOf('large') > -1)
        {
        	appendMessageUI("Sorry, that's too long for me to play.");
        }
        else
        {
        	appendMessageUI(error.name + ": " + error.message);
        }
        
    } 
    else {
        // Get selected data.
        var dataValue = asyncResult.value; 
        
        if(dataValue.length == 1)
        {
            
            if(dataValue[0].length == 1)
            {
                //special case:
                /*
                if(dataValue[0] == '')
                {
                    
                }
                else if(isNaN(dataValue[0]))
                {
                    
                }
                else
                {
                    audios[12].play();
                }
                */
            }
            else
            {
                //normal case
                for(var c = 0;c < dataValue[0].length;c++)
                {
                    if(isNaN(dataValue[0][c]))
                    {
                        goodData = false;
                    }
                    myData.push(dataValue[0][c]);
                }
                //appendMessage("found notes");
                if(goodData)
                {
                   calculateNotes(); 
                }
                else
                {
                    appendMessageUI("Select only numbers.");
                }
                
            }
        }
        else
        {
            if(dataValue[0].length == 1)
            {
                for(var c = 0;c < dataValue.length;c++)
                {
                    if(isNaN(dataValue[c]))
                    {
                        goodData = false;
                    }
                    myData.push(dataValue[c]);
                }
                //appendMessage("found notes");
                if(goodData)
                {
                   calculateNotes(); 
                }
                else
                {
                    appendMessageUI("Select only numbers.");
                }
            }
            else
            {
                appendMessageUI("Choose just one row or column.");
            }
        }
        //appendMessageUI(myData);
        
    }          
}

function calculateNotes()
{
    //appendMessage("calculating notes.");
    //document.getElementById("playSelected").disabled = true;
    
    var myMax = parseFloat(Math.max.apply(Math,myData));
    var myMin = parseFloat(Math.min.apply(Math,myData));
    //appendMessageUI(myData);
    var numberNotes = [];
    for(var c = 0;c < myData.length;c++)
    {

        var myResult = Math.round((parseFloat(myData[c]) - myMin) * 36.00 / (myMax-myMin));
        if(isNaN(myResult))
        {
            numberNotes.push('x');
        }
        else
        {
            numberNotes.push(myResult);
        }


    }
    //appendMessageUI(numberNotes);
    
    if(numberNotes.indexOf(36) != -1)
    {
	    document.getElementById("playSelected").style.opacity = 0.5;
    	document.getElementById("playSelected").disabled = "disabled";
	    var theCurrentTime = 0;
	    //appendMessage(numberNotes);
	    playTheNextNote(numberNotes,theCurrentTime);
    }    
}

function playTheNextNote(numberNotes,theCurrentTime)
{

    var i = numberNotes[theCurrentTime];
    var played = 0;
    
    if(i == "x")
    {
        //appendMessageUI("hello!");
    }
    else
    {
        //appendMessageUI("goodbye");
        /*
        if(audios[i].readyState != 0)
        {
            audios[i].play();
            played = i;
            appendMessageUI(i + " " + audios[i].readyState);
            if(audios[i+37].ended)
            {
                audios[i+37].load();
            }
            if(audios[i+74].ended)
            {
                audios[i+74].load();
            }
        }
        else if(audios[i+37].readyState != 0)
        {
            audios[i+37].play();
            appendMessageUI(i+37 + " " + audios[i+37].readyState);
            played=i+37
            if(audios[i].ended)
            {
                audios[i].load();
            }
            if(audios[i+74].ended)
            {
                audios[i+74].load();
            }
        }
        else
        {
            audios[i+74].play();
            appendMessageUI(i+74 + " " + audios[i+74].readyState);
            played=i+74;
            if(audios[i].ended)
            {
                audios[i].load();
            }
            if(audios[i+37].ended)
            {
                audios[i+37].load();
            }
        } 
        //appendMessage(i);
        //audios[i].play();
        */
        appendMusic(noteNames[i]);
        //appendMessageUI("addedMusic");
    }
    //appendMessageUI(theCurrentTime + " " + numberNotes.length);
    theCurrentTime = theCurrentTime + 1;
    
    //appendMessageUI(theCurrentTime);
    
    if(theCurrentTime < numberNotes.length)
    {
        
        //appendMessage("setting timeout");
        setTimeout(function(){
            //appendMessage("goneback");
            playTheNextNote(numberNotes,theCurrentTime);
        },300);
    }
    else
    {
        //appendMessageUI(theCurrentTime);
        //document.getElementById("playSelected").disabled = false;
        
        setTimeout(function(){
            document.getElementById("music").innerHTML = "";
            document.getElementById("playSelected").style.opacity = 1;
            document.getElementById("playSelected").disabled = "";

        },1500);
        
        
    //    setTimeout(function(){
    //        audios[played].load();
    //    },300);        
    }
}

function clearMyMusic()
{
    document.getElementById("music").innerHTML = "";
}

function goBack(numberNotes,theCurrentTime)
{
    //appendMessage("goneback");
    playTheNextNote(numberNotes,theCurrentTime);
}



function myTest2(){
    //appendMessage("<p> hello. how are you? " + document.getElementById('music').innerHTML) + "<br> fine </p>";
}

function createTable(){
    // Build table.
    var myTable = new Office.TableData();
    myTable.headers = [['C3','C#3','D3','D#3','E3','F3','F#3','G3','G#3','A3','A#3','B3','C','C#','D','D#','E','F','F#','G','G#','A','A#','B','C5','C#5','D5','D#5','E5','F5','F#5','G5','G#5','A5','A#5','B5','C6']];
    myTable.rows = [['','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','']];
    rowNum = 1;
    while(rowNum < 64){
        myTable.rows.push(['','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','']);
        rowNum = rowNum + 1;
    }

    // Write table.
    Office.context.document.setSelectedDataAsync(myTable, {coercionType: "table"},
        function (result) {
            var error = result.error
            if (result.status === "failed") {
                write(error.name + ": " + error.message);
            }
        setBinding();
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
            //appendMessage('Addeda newa binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
            writeDataFromBinding(asyncResult.value);
            doEvent();
            
        }
    );
  
}

function onBindingSelectionChanged(eventArgs)
{
    //appendMessage(eventArgs.startRow);
    if(Office.context.document.settings.get("songIsPlaying") != 1 && eventArgs.startRow >= 0) //&& eventArgs.columnCount == 1 && eventArgs.rowCount == 1
    {
        audios[eventArgs.startColumn].play();
        
    }
}

function preloadNote(name)
{
    //var audio = document.createElement("audio");
    //audio.src = "Content/" + name + ".mp3";
    //audios.push(audio);
    noteNames.push(name);
    //if(audios.length == 111)
    //{
    //    audios[110].oncanplay=testReady();
    //}
}

function testReady()
{
    appendMessageUI(audios[0].readyState + " " + audios[8].readyState + " " + audios[9].readyState);
}

function addNote(noteName)
{
    var audio = document.createElement("audio");
    audio.src = "Content/" + noteName + ".mp3";
    audio.autoplay = true;
    
} 


function writeNotesData()
{
    //appendMessage(readNoteMatrix()[0]);
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
    Office.select("bindings#MyBinding6").addHandlerAsync(Office.EventType.BindingSelectionChanged, onBindingSelectionChanged);
}

function playSound(){
    //appendMessage("Play Notea");
    Office.context.document.setSelectedDataAsync("=Row()",
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed"){
            appendMessageUI(error.name + ": " + error.message);
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
                    appendMessageUI(error.name + ": " + error.message);
                } 
                else {
                    // Get selected data.
                    var dataValue = asyncResult.value; 
                    Office.context.document.setSelectedDataAsync(dataValue,
                    function (asyncResult) {
                        var error = asyncResult.error;
                        if (asyncResult.status === "failed"){
                            appendMessageUI(error.name + ": " + error.message);
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
    
    document.getElementById("stopsong").style.visibility = "visible";
    
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
        var chord = [["","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""]];
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
        setTimeout(appendMessageUI("one"),900);
        setTimeout(appendMessageUI("two"),900);
        setTimeout(appendMessageUI("three"),900);
        setTimeout(appendMessageUI("four"),1300);
        setTimeout(appendMessageUI("five"),1700);
        while(myTime < myLoopLength){
            setTimeout(playTimedChord(chord[myTime]),myTime*period);
            setTimeout(appendMessageUI(chord[myTime]),myTime*period);
            myTime = myTime + 1;
        }
    
        setTimeout(playTimedNext,myTime*period);
    }
    
}

function playTimedChord(chord){
    var mynoteIndex = 0;
    while(mynoteIndex < 13){
        addNote(chord[mynoteIndex]);
        mynoteIndex += 1;
    }
}

function getChordAtTime(myTime,myNotes){
    var myChord = ["","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""];
    /*for(var i = 0;i < myNotes[myTime].length;i++)
    {
        
    }*/
    
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
    var chord = ["","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""];
    var myTime = Office.context.document.settings.get("time");
    var myNoteMatrix = Office.context.document.settings.get("noteMatrix");
    
    for(var i = 0;i < 37;i++)
    {
        if(myNoteMatrix[myTime][i] == 1)
        {
            if(audios[i].currentTime == 0)
            {
                
                audios[i].play();
                if(audios[i+37].ended)
                {
                    audios[i+37].load();
                }
                if(audios[i+74].ended)
                {
                    audios[i+74].load();
                }
            }
            else if(audios[i+37].currentTime == 0)
            {
                audios[i+37].play();
                if(audios[i].ended)
                {
                    audios[i].load();
                }
                if(audios[i+74].ended)
                {
                    audios[i+74].load();
                }
            }
            else
            {
                audios[i+74].play();
                if(audios[i].ended)
                {
                    audios[i].load();
                }
                if(audios[i+37].ended)
                {
                    audios[i+37].load();
                }
            }
            
        }
    }
    
    /*
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
        addNote(chord[noteIndex]);
        
        noteIndex += 1;
    }
    */
    
    
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
    document.getElementById("stopsong").style.visibility = "hidden";
    clearInterval(Office.context.document.settings.get("intervalID"));
    
}





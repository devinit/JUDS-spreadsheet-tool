function onInstall() {
  onOpen()
  translateGeoCode()
  translateSectorCode();  
}
function onOpen() {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('Current version','version')    
    .addToUi();
}
function version() {
  var title = 'Joined-up Data Standards Navigator spreadsheet tool';
  var message = 'This is version 1.0. Please contact us at info@joinedupdata.org if you experience any problems using this tool or to suggest an improvement in its functionality.';
  var ui = SpreadsheetApp.getUi();
  ui.alert(title, message, ui.ButtonSet.OK);
}
/**
*
* Return sector code from Joined-up Data Standards Navigator 
*
* @param {crs}   from    The data standard to translate code from or select the cell where this information can be found
* @param {cofog} to      The data standard to translate code to or select the cell where this information can be found
* @param {32310} code    The code for translation or select the cell where this information can be found
* @return                The translated code 
* @customfunction
*/
function translateSectorCode(from, to, code){

  var project= new Array();
  if(to.toLowerCase()=="crs" || to.toLowerCase()=="cofog" ||from.toLowerCase()=="itep"  || to.toLowerCase()=="ntee" ||to.toLowerCase()=="world_bank_themes"  || to.toLowerCase()=="world_bank_sectors" || to.toLowerCase()=="isic"){
    project='Sectors';
  }
  if(to.toLowerCase()=="sdg" || to.toLowerCase()=="mdg" ||from.toLowerCase()=="wdi"){
    project='Indicators';
  }  
  if(to.toLowerCase()=="dhsq7" || to.toLowerCase()=="mics5" ||from.toLowerCase()=="lsms"){
    project='Surveys';
  }    
  var output2= new Array()
  var output3= new Array()
  var outputBroad= new Array()
  var UrlArrayExact= new Array();
    var url= "http://178.79.158.119:3030/Sectors/query?query=PREFIX+skos%3A%3Chttp%3A//www.w3.org/2004/02/skos/core%23%3E%0ASELECT+DISTINCT+%3FConcept+%3Fnotation++%3FexactMatch+%3Fnotation2%0AWHERE%7B%3FConcept+%3Fx+skos%3AConcept%7B%3FConcept+skos%3Anotation+%3Fnotation+.+Filter+%28regex%28str%28%3Fnotation%29%2C%27";
    url += code;
    url +="%27%2C+%27i%27%29%29%7D%0A%7B%3FConcept+skos%3AexactMatch+%3FexactMatch+.%7D%0ASERVICE+%3Chttp%3A//178.79.158.119%3A3030/";
    url += project;
    url += "/query%3E%7B%7B%3FexactMatch+skos%3Anotation+%3Fnotation2+.%7D%7D%0A%7DORDER+BY+%3FprefLabel+LIMIT+500+OFFSET+0&output=xml&results=xml&format=xml";
    UrlArrayExact.push([url]);     

  var UrlArrayClose= new Array();
    var url= "http://178.79.158.119:3030/Sectors/query?query=PREFIX+skos%3A%3Chttp%3A//www.w3.org/2004/02/skos/core%23%3E%0ASELECT+DISTINCT+%3FConcept+%3Fnotation++%3FcloseMatch+%3Fnotation2%0AWHERE%7B%3FConcept+%3Fx+skos%3AConcept%7B%3FConcept+skos%3Anotation+%3Fnotation+.+Filter+%28regex%28str%28%3Fnotation%29%2C%27";
    url += code;
    url +="%27%2C+%27i%27%29%29%7D%0A%7B%3FConcept+skos%3AcloseMatch+%3FcloseMatch+.%7D%0ASERVICE+%3Chttp%3A//178.79.158.119%3A3030/";
    url += project;
    url += "/query%3E%7B%7B%3FcloseMatch+skos%3Anotation+%3Fnotation2+.%7D%7D%0A%7DORDER+BY+%3FprefLabel+LIMIT+500+OFFSET+0&output=xml&results=xml&format=xml";
    UrlArrayClose.push([url]);     

  var UrlArrayBroad= new Array();
    var url= "http://178.79.158.119:3030/Sectors/query?query=PREFIX+skos%3A%3Chttp%3A//www.w3.org/2004/02/skos/core%23%3E%0ASELECT+DISTINCT+%3FConcept+%3Fnotation++%3FbroadMatch+%3Fnotation2%0AWHERE%7B%3FConcept+%3Fx+skos%3AConcept%7B%3FConcept+skos%3Anotation+%3Fnotation+.+Filter+%28regex%28str%28%3Fnotation%29%2C%27";
    url += code;
    url +="%27%2C+%27i%27%29%29%7D%0A%7B%3FConcept+skos%3AbroadMatch+%3FbroadMatch+.%7D%0ASERVICE+%3Chttp%3A//178.79.158.119%3A3030/"; 
    url += project;
    url += "/query%3E%7B%7B%3FcloseMatch+skos%3Anotation+%3Fnotation2+.%7D%7D%0A%7DORDER+BY+%3FprefLabel+LIMIT+500+OFFSET+0&output=xml&results=xml&format=xml";  
    UrlArrayBroad.push([url]);     

  var UrlArrayNarrow= new Array();
    var url= "http://178.79.158.119:3030/Sectors/query?query=PREFIX+skos%3A%3Chttp%3A//www.w3.org/2004/02/skos/core%23%3E%0ASELECT+DISTINCT+%3FConcept+%3Fnotation++%3FnarrowMatch+%3Fnotation2%0AWHERE%7B%3FConcept+%3Fx+skos%3AConcept%7B%3FConcept+skos%3Anotation+%3Fnotation+.+Filter+%28regex%28str%28%3Fnotation%29%2C%27";
    url += code;
    url +="%27%2C+%27i%27%29%29%7D%0A%7B%3FConcept+skos%3AnarrowMatch+%3FnarrowMatch+.%7D%0ASERVICE+%3Chttp%3A//178.79.158.119%3A3030/";
    url += project;
    url += "/query%3E%7B%7B%3FnarrowMatch+skos%3Anotation+%3Fnotation2+.%7D%7D%0A%7DORDER+BY+%3FprefLabel+LIMIT+500+OFFSET+0&output=xml&results=xml&format=xml";  
    UrlArrayNarrow.push([url]);
    
  var UrlIndicatorsArrayExact= new Array();
    var url= "http://178.79.158.119:3030/Indicators/query?query=PREFIX+skos%3A%3Chttp%3A//www.w3.org/2004/02/skos/core%23%3E%0ASELECT+DISTINCT+%3FConcept+%3Fnotation++%3FexactMatch+%3Fnotation2%0AWHERE%7B%3FConcept+%3Fx+skos%3AConcept%7B%3FConcept+skos%3Anotation+%3Fnotation+.+Filter+%28regex%28str%28%3Fnotation%29%2C%27";
    url += code;
    url +="%27%2C+%27i%27%29%29%7D%0A%7B%3FConcept+skos%3AexactMatch+%3FexactMatch+.%7D%0ASERVICE+%3Chttp%3A//178.79.158.119%3A3030/";
    url += project;
    url += "/query%3E%7B%7B%3FexactMatch+skos%3Anotation+%3Fnotation2+.%7D%7D%0A%7DORDER+BY+%3FprefLabel+LIMIT+500+OFFSET+0&output=xml&results=xml&format=xml";  
    UrlIndicatorsArrayExact.push([url]); 
    
  var UrlIndicatorsArrayClose= new Array();
    var url= "http://178.79.158.119:3030/Indicators/query?query=PREFIX+skos%3A%3Chttp%3A//www.w3.org/2004/02/skos/core%23%3E%0ASELECT+DISTINCT+%3FConcept+%3Fnotation++%3FcloseMatch+%3Fnotation2%0AWHERE%7B%3FConcept+%3Fx+skos%3AConcept%7B%3FConcept+skos%3Anotation+%3Fnotation+.+Filter+%28regex%28str%28%3Fnotation%29%2C%27";
    url += code;
    url +="%27%2C+%27i%27%29%29%7D%0A%7B%3FConcept+skos%3AcloseMatch+%3FcloseMatch+.%7D%0ASERVICE+%3Chttp%3A//178.79.158.119%3A3030/";
    url += project;
    url += "/query%3E%7B%7B%3FcloseMatch+skos%3Anotation+%3Fnotation2+.%7D%7D%0A%7DORDER+BY+%3FprefLabel+LIMIT+500+OFFSET+0&output=xml&results=xml&format=xml"; 
    UrlIndicatorsArrayClose.push([url]);    
    
  var UrlIndicatorsArrayNarrow= new Array();
    var url= "http://178.79.158.119:3030/Indicators/query?query=PREFIX+skos%3A%3Chttp%3A//www.w3.org/2004/02/skos/core%23%3E%0ASELECT+DISTINCT+%3FConcept+%3Fnotation++%3FnarrowMatch+%3Fnotation2%0AWHERE%7B%3FConcept+%3Fx+skos%3AConcept%7B%3FConcept+skos%3Anotation+%3Fnotation+.+Filter+%28regex%28str%28%3Fnotation%29%2C%27";
    url += code;
    url +="%27%2C+%27i%27%29%29%7D%0A%7B%3FConcept+skos%3AnarrowMatch+%3FnarrowMatch+.%7D%0ASERVICE+%3Chttp%3A//178.79.158.119%3A3030/";
    url += project;
    url += "/query%3E%7B%7B%3FnarrowMatch+skos%3Anotation+%3Fnotation2+.%7D%7D%0A%7DORDER+BY+%3FprefLabel+LIMIT+500+OFFSET+0&output=xml&results=xml&format=xml"; 
    UrlIndicatorsArrayNarrow.push([url]);      
 
  var UrlIndicatorsArrayBroad= new Array();
    var url= "http://178.79.158.119:3030/Indicators/query?query=PREFIX+skos%3A%3Chttp%3A//www.w3.org/2004/02/skos/core%23%3E%0ASELECT+DISTINCT+%3FConcept+%3Fnotation++%3FbroadMatch+%3Fnotation2%0AWHERE%7B%3FConcept+%3Fx+skos%3AConcept%7B%3FConcept+skos%3Anotation+%3Fnotation+.+Filter+%28regex%28str%28%3Fnotation%29%2C%27";
    url += code;
    url +="%27%2C+%27i%27%29%29%7D%0A%7B%3FConcept+skos%3AbroadMatch+%3FbroadMatch+.%7D%0ASERVICE+%3Chttp%3A//178.79.158.119%3A3030/";
    url += project;
    url += "/query%3E%7B%7B%3FbroadMatch+skos%3Anotation+%3Fnotation2+.%7D%7D%0A%7DORDER+BY+%3FprefLabel+LIMIT+500+OFFSET+0&output=xml&results=xml&format=xml"; 
    UrlIndicatorsArrayBroad.push([url]);    
    
  var UrlSurveysArrayExact= new Array();
    var url= "http://178.79.158.119:3030/Surveys/query?query=PREFIX+skos%3A%3Chttp%3A//www.w3.org/2004/02/skos/core%23%3E%0ASELECT+DISTINCT+%3FConcept+%3Fnotation++%3FexactMatch+%3Fnotation2%0AWHERE%7B%3FConcept+%3Fx+skos%3AConcept%7B%3FConcept+skos%3Anotation+%3Fnotation+.+Filter+%28regex%28str%28%3Fnotation%29%2C%27";
    url += code;
    url +="%27%2C+%27i%27%29%29%7D%0A%7B%3FConcept+skos%3AexactMatch+%3FexactMatch+.%7D%0ASERVICE+%3Chttp%3A//178.79.158.119%3A3030/";
    url += project;
    url += "/query%3E%7B%7B%3FexactMatch+skos%3Anotation+%3Fnotation2+.%7D%7D%0A%7DORDER+BY+%3FprefLabel+LIMIT+500+OFFSET+0&output=xml&results=xml&format=xml"; 
    UrlSurveysArrayExact.push([url]); 
    
  var UrlSurveysArrayClose= new Array();
    var url= "http://178.79.158.119:3030/Surveys/query?query=PREFIX+skos%3A%3Chttp%3A//www.w3.org/2004/02/skos/core%23%3E%0ASELECT+DISTINCT+%3FConcept+%3Fnotation++%3FcloseMatch+%3Fnotation2%0AWHERE%7B%3FConcept+%3Fx+skos%3AConcept%7B%3FConcept+skos%3Anotation+%3Fnotation+.+Filter+%28regex%28str%28%3Fnotation%29%2C%27";
    url += code;
    url +="%27%2C+%27i%27%29%29%7D%0A%7B%3FConcept+skos%3AcloseMatch+%3FcloseMatch+.%7D%0ASERVICE+%3Chttp%3A//178.79.158.119%3A3030/";
    url += project;
    url += "/query%3E%7B%7B%3FcloseMatch+skos%3Anotation+%3Fnotation2+.%7D%7D%0A%7DORDER+BY+%3FprefLabel+LIMIT+500+OFFSET+0&output=xml&results=xml&format=xml"; 
    UrlSurveysArrayClose.push([url]);    
    
  var UrlSurveysArrayNarrow= new Array();
    var url= "http://178.79.158.119:3030/Surveys/query?query=PREFIX+skos%3A%3Chttp%3A//www.w3.org/2004/02/skos/core%23%3E%0ASELECT+DISTINCT+%3FConcept+%3Fnotation++%3FnarrowMatch+%3Fnotation2%0AWHERE%7B%3FConcept+%3Fx+skos%3AConcept%7B%3FConcept+skos%3Anotation+%3Fnotation+.+Filter+%28regex%28str%28%3Fnotation%29%2C%27";
    url += code;
    url +="%27%2C+%27i%27%29%29%7D%0A%7B%3FConcept+skos%3AnarrowMatch+%3FnarrowMatch+.%7D%0ASERVICE+%3Chttp%3A//178.79.158.119%3A3030/";
    url += project;
    url += "/query%3E%7B%7B%3FnarrowMatch+skos%3Anotation+%3Fnotation2+.%7D%7D%0A%7DORDER+BY+%3FprefLabel+LIMIT+500+OFFSET+0&output=xml&results=xml&format=xml"; 
    UrlSurveysArrayNarrow.push([url]);      
 
  var UrlSurveysArrayBroad= new Array();
    var url= "http://178.79.158.119:3030/Surveys/query?query=PREFIX+skos%3A%3Chttp%3A//www.w3.org/2004/02/skos/core%23%3E%0ASELECT+DISTINCT+%3FConcept+%3Fnotation++%3FbroadMatch+%3Fnotation2%0AWHERE%7B%3FConcept+%3Fx+skos%3AConcept%7B%3FConcept+skos%3Anotation+%3Fnotation+.+Filter+%28regex%28str%28%3Fnotation%29%2C%27";
    url += code;
    url +="%27%2C+%27i%27%29%29%7D%0A%7B%3FConcept+skos%3AbroadMatch+%3FbroadMatch+.%7D%0ASERVICE+%3Chttp%3A//178.79.158.119%3A3030/";
    url += project;
    url += "/query%3E%7B%7B%3FbroadMatch+skos%3Anotation+%3Fnotation2+.%7D%7D%0A%7DORDER+BY+%3FprefLabel+LIMIT+500+OFFSET+0&output=xml&results=xml&format=xml"; 
    UrlSurveysArrayBroad.push([url]);        

  if(from.toLowerCase()=="crs" || from.toLowerCase()=="cofog" || from.toLowerCase()=="isic" || from.toLowerCase()== "ntee" || from.toLowerCase()== "world_bank_themes" || from.toLowerCase()== "world_bank_sectors" || from.toLowerCase()== "itep"){
    for (var b=0; b <UrlArrayExact.length; b++){
          var url2= UrlArrayExact[b];
          var url= String(url2);
          var xml = UrlFetchApp.fetch(url).getContentText();
          var document = XmlService.parse(xml);
          var namespace = XmlService.getNamespace('http://www.w3.org/2005/sparql-results#');
          var root = document.getRootElement();
          var output = new Array();
          var results = root.getChild('results', namespace);
          var columnList = results.getChildren('result', namespace);
          var Concept = '';
          var exactMatch = '';
          var prefLabel = '';
          var notation = '';
          var notation2 = '';    
          for (var i = 0; i < columnList.length; i++){
              var column = columnList[i];
              var bindings = column.getChildren("binding", namespace);
              for (var j=0; j < bindings.length; j++){
                var bind = bindings[j];
                var attrName = bind.getAttribute('name').getValue();
                var uri = "";
                var text="";
                if(bind.getChild('literal', namespace)){ 
                  text = bind.getChild('literal', namespace).getValue();
                }  
                if(bind.getChild('uri', namespace)) { 
                     uri = bind.getChild('uri', namespace).getValue();
                }
                if (attrName === 'Concept') {
                  Concept = uri;
                } else if (attrName === 'exactMatch'){
                  exactMatch = uri;        
                } else if (attrName === 'notation'){
                  notation= text;
                } else if (attrName === 'notation2'){
                  notation2= text;
                } 
              } 
              output.push([Concept, exactMatch, notation, notation2]);
            }   
          }    
         for (var b=0; b <UrlArrayClose.length; b++){
          var url2= UrlArrayClose[b];
          var url= String(url2);
          var xml = UrlFetchApp.fetch(url).getContentText();
          var document = XmlService.parse(xml);
          var namespace = XmlService.getNamespace('http://www.w3.org/2005/sparql-results#');
          var root = document.getRootElement();
          var results = root.getChild('results', namespace);
          var columnList = results.getChildren('result', namespace);
          var Concept = '';
          var closeMatch='';
          var prefLabel = '';
          var notation = '';
          var notation2 = '';    
          for (var i = 0; i < columnList.length; i++){
              var column = columnList[i];
              var bindings = column.getChildren("binding", namespace);
              for (var j=0; j < bindings.length; j++){
                var bind = bindings[j];
                var attrName = bind.getAttribute('name').getValue();
                var uri = "";
                var text="";
                if(bind.getChild('literal', namespace)){ 
                  text = bind.getChild('literal', namespace).getValue();
                }  
                if(bind.getChild('uri', namespace)) { 
                     uri = bind.getChild('uri', namespace).getValue();
                }
                if (attrName === 'Concept') {
                  Concept = uri;
                } else if (attrName === 'closeMatch'){
                  closeMatch = uri;        
                } else if (attrName === 'notation'){
                  notation= text;
                } else if (attrName === 'notation2'){
                  notation2= text;
                } 
              } 
              output.push([Concept, closeMatch, notation, notation2]);
            }   
           for (var k=0; k<output.length; k++){
             var object= output[k];
             var origin = String(object).split(',')[0].split('/')[4];
             var translation = String(object).split(',')[1].split('/')[4];
             if(origin==(from).toLowerCase() && translation==(to).toLowerCase()){
               output2.push(object);   
               }            
             }
           }
           for (var l=0; l<output2.length; l++){
               var object2= output2[l];
               var origin2 = String(object2).split(',')[2];
               var translation2= String(object2).split(',')[3]
               output3.push([translation2])
           } 
           
           if (output2.length==0){
            for (var b=0; b <UrlArrayBroad.length; b++){
              var url2= UrlArrayBroad[b];
              var url= String(url2);
              var xml = UrlFetchApp.fetch(url).getContentText();
              var document = XmlService.parse(xml);
              var namespace = XmlService.getNamespace('http://www.w3.org/2005/sparql-results#');
              var root = document.getRootElement();
              var results = root.getChild('results', namespace);
              var columnList = results.getChildren('result', namespace);
              var Concept = '';
              var outputBroad = new Array();      
              var broadMatch='';
              var prefLabel = '';
              var notation = '';
              var notation2 = '';    
              for (var i = 0; i < columnList.length; i++){
                  var column = columnList[i];
                  var bindings = column.getChildren("binding", namespace);
                  for (var j=0; j < bindings.length; j++){
                    var bind = bindings[j];
                    var attrName = bind.getAttribute('name').getValue();
                    var uri = "";
                    var text="";
                    if(bind.getChild('literal', namespace)){ 
                      text = bind.getChild('literal', namespace).getValue();
                    }  
                    if(bind.getChild('uri', namespace)) { 
                         uri = bind.getChild('uri', namespace).getValue();
                    }
                    if (attrName === 'Concept') {
                      Concept = uri;
                    } else if (attrName === 'broadMatch'){
                      broadMatch = uri;        
                    } else if (attrName === 'notation'){
                      notation= text;
                    } else if (attrName === 'notation2'){
                      notation2= text;
                    } 
                  } 
                  outputBroad.push([Concept, broadMatch, notation, notation2]);
                }   
                if (outputBroad.length==0){  
                  var error2 = "There is no exact or more general translation from the code you entered";
                 return error2; }
                }              
               for (var k=0; k<outputBroad.length; k++){
                 var outputBroad2 = new Array();                 
                 var object= outputBroad[k];
                 var origin = String(object).split(',')[0].split('/')[4];
                 var translation = String(object).split(',')[1].split('/')[4];
                 if(origin==(from).toLowerCase() && translation==(to).toLowerCase()){
                   outputBroad2.push(object);   
                   }            
                 }
                for (var l=0; l<outputBroad2.length; l++){
                   var object2= outputBroad2[l];
                   var origin2 = String(object2).split(',')[2];
                   var translation2= String(object2).split(',')[3]
                   var messageBroad="Please be aware that this is a broader match to your query";
                   output3.push([translation2, messageBroad])
                 } 
               }           
    
           return output3;
      } 
  if(from.toLowerCase()=="sdg" || from.toLowerCase()=="mdg" || from.toLowerCase()=="wdi"){
      for (var b=0; b <UrlIndicatorsArrayExact.length; b++){
          var url2= UrlIndicatorsArrayExact[b];
          var url= String(url2);
          var xml = UrlFetchApp.fetch(url).getContentText();
          var document = XmlService.parse(xml);
          var namespace = XmlService.getNamespace('http://www.w3.org/2005/sparql-results#');
          var root = document.getRootElement();
          var output = new Array();
          var results = root.getChild('results', namespace);
          var columnList = results.getChildren('result', namespace);
          var Concept = '';
          var exactMatch = '';
          var prefLabel = '';
          var notation = '';
          var notation2 = '';    
          for (var i = 0; i < columnList.length; i++){
              var column = columnList[i];
              var bindings = column.getChildren("binding", namespace);
              for (var j=0; j < bindings.length; j++){
                var bind = bindings[j];
                var attrName = bind.getAttribute('name').getValue();
                var uri = "";
                var text="";
                if(bind.getChild('literal', namespace)){ 
                  text = bind.getChild('literal', namespace).getValue();
                }  
                if(bind.getChild('uri', namespace)) { 
                     uri = bind.getChild('uri', namespace).getValue();
                }
                if (attrName === 'Concept') {
                  Concept = uri;
                } else if (attrName === 'exactMatch'){
                  exactMatch = uri;        
                } else if (attrName === 'notation'){
                  notation= text;
                } else if (attrName === 'notation2'){
                  notation2= text;
                } 
              } 
              output.push([Concept, exactMatch, notation, notation2]);
            }   
          }    
         for (var b=0; b <UrlIndicatorsArrayClose.length; b++){
          var url2= UrlIndicatorsArrayClose[b];
          var url= String(url2);
          var xml = UrlFetchApp.fetch(url).getContentText();
          var document = XmlService.parse(xml);
          var namespace = XmlService.getNamespace('http://www.w3.org/2005/sparql-results#');
          var root = document.getRootElement();
          var results = root.getChild('results', namespace);
          var columnList = results.getChildren('result', namespace);
          var Concept = '';
          var closeMatch='';
          var prefLabel = '';
          var notation = '';
          var notation2 = '';    
          for (var i = 0; i < columnList.length; i++){
              var column = columnList[i];
              var bindings = column.getChildren("binding", namespace);
              for (var j=0; j < bindings.length; j++){
                var bind = bindings[j];
                var attrName = bind.getAttribute('name').getValue();
                var uri = "";
                var text="";
                if(bind.getChild('literal', namespace)){ 
                  text = bind.getChild('literal', namespace).getValue();
                }  
                if(bind.getChild('uri', namespace)) { 
                     uri = bind.getChild('uri', namespace).getValue();
                }
                if (attrName === 'Concept') {
                  Concept = uri;
                } else if (attrName === 'closeMatch'){
                  closeMatch = uri;        
                } else if (attrName === 'notation'){
                  notation= text;
                } else if (attrName === 'notation2'){
                  notation2= text;
                } 
              } 
              output.push([Concept, closeMatch, notation, notation2]);
            }   
           for (var k=0; k<output.length; k++){
             var object= output[k];
             var origin = String(object).split(',')[0].split('/')[4];
             var translation = String(object).split(',')[1].split('/')[4];
             if(origin==(from).toLowerCase() && translation==(to).toLowerCase()){
               output2.push(object);   
               }            
             }
           }
           for (var l=0; l<output2.length; l++){
               var object2= output2[l];
               var origin2 = String(object2).split(',')[2];
               var translation2= String(object2).split(',')[3]
               output3.push([translation2])
           } 
           
           if (output2.length==0){
            for (var b=0; b <UrlIndicatorsArrayBroad.length; b++){
              var url2= UrlIndicatorsArrayBroad[b];
              var url= String(url2);
              var xml = UrlFetchApp.fetch(url).getContentText();
              var document = XmlService.parse(xml);
              var namespace = XmlService.getNamespace('http://www.w3.org/2005/sparql-results#');
              var root = document.getRootElement();
              var results = root.getChild('results', namespace);
              var columnList = results.getChildren('result', namespace);
              var Concept = '';
              var outputBroad = new Array();      
              var broadMatch='';
              var prefLabel = '';
              var notation = '';
              var notation2 = '';    
              for (var i = 0; i < columnList.length; i++){
                  var column = columnList[i];
                  var bindings = column.getChildren("binding", namespace);
                  for (var j=0; j < bindings.length; j++){
                    var bind = bindings[j];
                    var attrName = bind.getAttribute('name').getValue();
                    var uri = "";
                    var text="";
                    if(bind.getChild('literal', namespace)){ 
                      text = bind.getChild('literal', namespace).getValue();
                    }  
                    if(bind.getChild('uri', namespace)) { 
                         uri = bind.getChild('uri', namespace).getValue();
                    }
                    if (attrName === 'Concept') {
                      Concept = uri;
                    } else if (attrName === 'broadMatch'){
                      broadMatch = uri;        
                    } else if (attrName === 'notation'){
                      notation= text;
                    } else if (attrName === 'notation2'){
                      notation2= text;
                    } 
                  } 
                  outputBroad.push([Concept, broadMatch, notation, notation2]);
                }   
                if (outputBroad.length==0){  
                  var error2 = "There is no exact or more general translation from the code you entered";
                 return error2; }
                }              
               for (var k=0; k<outputBroad.length; k++){
                 var outputBroad2 = new Array();                 
                 var object= outputBroad[k];
                 var origin = String(object).split(',')[0].split('/')[4];
                 var translation = String(object).split(',')[1].split('/')[4];
                 if(origin==(from).toLowerCase() && translation==(to).toLowerCase()){
                   outputBroad2.push(object);   
                   }            
                 }
                for (var l=0; l<outputBroad2.length; l++){
                   var object2= outputBroad2[l];
                   var origin2 = String(object2).split(',')[2];
                   var translation2= String(object2).split(',')[3]
                   var messageBroad="Please be aware that this is a broader match to your query";
                   output3.push([translation2, messageBroad])
                 } 
               }           
    
           return output3;
      }       
  if(from.toLowerCase()=="mics5" || from.toLowerCase()=="dhsq7" || from.toLowerCase()=="lsms"){
      for (var b=0; b <UrlSurveysArrayExact.length; b++){
          var url2= UrlSurveysArrayExact[b];
          var url= String(url2);
          var xml = UrlFetchApp.fetch(url).getContentText();
          var document = XmlService.parse(xml);
          var namespace = XmlService.getNamespace('http://www.w3.org/2005/sparql-results#');
          var root = document.getRootElement();
          var output = new Array();
          var results = root.getChild('results', namespace);
          var columnList = results.getChildren('result', namespace);
          var Concept = '';
          var exactMatch = '';
          var prefLabel = '';
          var notation = '';
          var notation2 = '';    
          for (var i = 0; i < columnList.length; i++){
              var column = columnList[i];
              var bindings = column.getChildren("binding", namespace);
              for (var j=0; j < bindings.length; j++){
                var bind = bindings[j];
                var attrName = bind.getAttribute('name').getValue();
                var uri = "";
                var text="";
                if(bind.getChild('literal', namespace)){ 
                  text = bind.getChild('literal', namespace).getValue();
                }  
                if(bind.getChild('uri', namespace)) { 
                     uri = bind.getChild('uri', namespace).getValue();
                }
                if (attrName === 'Concept') {
                  Concept = uri;
                } else if (attrName === 'exactMatch'){
                  exactMatch = uri;        
                } else if (attrName === 'notation'){
                  notation= text;
                } else if (attrName === 'notation2'){
                  notation2= text;
                } 
              } 
              output.push([Concept, exactMatch, notation, notation2]);
            }   
          }    
         for (var b=0; b <UrlSurveysArrayClose.length; b++){
          var url2= UrlSurveysArrayClose[b];
          var url= String(url2);
          var xml = UrlFetchApp.fetch(url).getContentText();
          var document = XmlService.parse(xml);
          var namespace = XmlService.getNamespace('http://www.w3.org/2005/sparql-results#');
          var root = document.getRootElement();
          var results = root.getChild('results', namespace);
          var columnList = results.getChildren('result', namespace);
          var Concept = '';
          var closeMatch='';
          var prefLabel = '';
          var notation = '';
          var notation2 = '';    
          for (var i = 0; i < columnList.length; i++){
              var column = columnList[i];
              var bindings = column.getChildren("binding", namespace);
              for (var j=0; j < bindings.length; j++){
                var bind = bindings[j];
                var attrName = bind.getAttribute('name').getValue();
                var uri = "";
                var text="";
                if(bind.getChild('literal', namespace)){ 
                  text = bind.getChild('literal', namespace).getValue();
                }  
                if(bind.getChild('uri', namespace)) { 
                     uri = bind.getChild('uri', namespace).getValue();
                }
                if (attrName === 'Concept') {
                  Concept = uri;
                } else if (attrName === 'closeMatch'){
                  closeMatch = uri;        
                } else if (attrName === 'notation'){
                  notation= text;
                } else if (attrName === 'notation2'){
                  notation2= text;
                } 
              } 
              output.push([Concept, closeMatch, notation, notation2]);
            }   
           for (var k=0; k<output.length; k++){
             var object= output[k];
             var origin = String(object).split(',')[0].split('/')[4].split('_')[2];
             var translation = String(object).split(',')[1].split('/')[4].split('_')[2];
             if(origin=="u5"){
               origin=String(origin).replace('u5','mics5')
             }
             if(translation=="u5"){
               translation=String(translation).replace('u5','mics5')
             }
                     
             if(origin==(from).toLowerCase() && translation==(to).toLowerCase()){
               output2.push(object);  
               }            
             }
           }
           for (var l=0; l<output2.length; l++){
               var object2= output2[l];
               var origin2 = String(object2).split(',')[2];
               var translation2= String(object2).split(',')[3];
               output3.push([translation2])
           } 
           
           if (output2.length==0){
            for (var b=0; b <UrlSurveysArrayBroad.length; b++){
              var url2= UrlSurveysArrayBroad[b];
              var url= String(url2);
              var xml = UrlFetchApp.fetch(url).getContentText();
              var document = XmlService.parse(xml);
              var namespace = XmlService.getNamespace('http://www.w3.org/2005/sparql-results#');
              var root = document.getRootElement();
              var results = root.getChild('results', namespace);
              var columnList = results.getChildren('result', namespace);
              var Concept = '';
              var outputBroad = new Array();      
              var broadMatch='';
              var prefLabel = '';
              var notation = '';
              var notation2 = '';    
              for (var i = 0; i < columnList.length; i++){
                  var column = columnList[i];
                  var bindings = column.getChildren("binding", namespace);
                  for (var j=0; j < bindings.length; j++){
                    var bind = bindings[j];
                    var attrName = bind.getAttribute('name').getValue();
                    var uri = "";
                    var text="";
                    if(bind.getChild('literal', namespace)){ 
                      text = bind.getChild('literal', namespace).getValue();
                    }  
                    if(bind.getChild('uri', namespace)) { 
                         uri = bind.getChild('uri', namespace).getValue();
                    }
                    if (attrName === 'Concept') {
                      Concept = uri;
                    } else if (attrName === 'broadMatch'){
                      broadMatch = uri;        
                    } else if (attrName === 'notation'){
                      notation= text;
                    } else if (attrName === 'notation2'){
                      notation2= text;
                    } 
                  } 
                  outputBroad.push([Concept, broadMatch, notation, notation2]);
                }   
                if (outputBroad.length==0){  
                  var error2 = "There is no exact or more general translation from the code you entered";
                 return error2; }
                }              
               for (var k=0; k<outputBroad.length; k++){
                 var outputBroad2 = new Array();                 
                 var object= outputBroad[k];
                 var origin = String(object).split(',')[0].split('/')[4];
                 var translation = String(object).split(',')[1].split('/')[4];
                 if(origin==(from).toLowerCase() && translation==(to).toLowerCase()){
                   outputBroad2.push(object);   
                   }            
                 }
                for (var l=0; l<outputBroad2.length; l++){
                   var object2= outputBroad2[l];
                   var origin2 = String(object2).split(',')[2];
                   var translation2= String(object2).split(',')[3]
                   var messageBroad="Please be aware that this is a broader match to your query";
                   output3.push([translation2, messageBroad])
                 } 
               }           
    
           return output3;
      }             
}


/**
*
* Return specific country code from Joined-up Data Standards Navigator 
*
* @param  {australia}  country  Country name of interest or select the cell where this information can be found
* @param  {dac}     to    The country classification to be returned or select the cell where this information can be found
* @return                 Specific country code requested
* @customfunction
*/
function translateGeoCode(country, to){
var output2= new Array()
var output3= new Array()
var outputBroad= new Array()
var UrlArrayExact= new Array();
  var url= "http://178.79.158.119:3030/Supranational/query?query=PREFIX+skos%3A%3Chttp%3A//www.w3.org/2004/02/skos/core%23%3E%0APREFIX+CountryTerms%3A%3Chttp%3A//joinedupdata.org/CountryTerms/%3E%0ASELECT+DISTINCT+%3FConcept+%3FprefLabel+%3Fdac_recip+%3Fdac_donor+%3Ffaostat_num+%3Fgaul_num+%3Fimf_ifs+%3Fiso_alpha_2+%3Fiso_alpha_3+%3Fiso_num+%3Fun_num+%3FSDG_indicators_global_database%0AWHERE%7B%3FConcept+%3Fx+skos%3AConcept%7B%3FConcept+skos%3AprefLabel+%3FprefLabel+.+FILTER+%28regex%28str%28%3FprefLabel%29%2C+%27%5E";
  url += country;
  url +="%27%2C+%27i%27%29%29+%7D%0AOptional%7B%3FConcept+CountryTerms%3Adac_donor+%3Fdac_donor+.%7D+%0AOptional%7B%3FConcept+CountryTerms%3Adac_recip+%3Fdac_recip+.%7D+%0A%7B%3FConcept+CountryTerms%3Afaostat_num+%3Ffaostat_num+.%7D%0A%7B%3FConcept+CountryTerms%3Agaul_num+%3Fgaul_num+.%7D%0A%7B%3FConcept+CountryTerms%3Aimf_ifs+%3Fimf_ifs+.%7D%0A%7B%3FConcept+CountryTerms%3Aiso_alpha-2+%3Fiso_alpha_2+.%7D%0A%7B%3FConcept+CountryTerms%3Aiso_alpha-3+%3Fiso_alpha_3+.%7D%0A%7B%3FConcept+CountryTerms%3Aiso_num+%3Fiso_num+.%7D%0A%7B%3FConcept+CountryTerms%3Aun_num+%3Fun_num+.%7D%0A%7B%3FConcept+CountryTerms%3ASDG_indicators_global_database+%3FSDG_indicators_global_database+.%7D+%0A%7DORDER+BY+%3FprefLabel+LIMIT+500+OFFSET+0+&output=xml&results=xml&format=xml"; 
  UrlArrayExact.push([url]);     

  for (var b=0; b <UrlArrayExact.length; b++){
    var url2= UrlArrayExact[b];
    var url= String(url2);
    var xml = UrlFetchApp.fetch(url).getContentText();
    var document = XmlService.parse(xml);
    var namespace = XmlService.getNamespace('http://www.w3.org/2005/sparql-results#');
    var root = document.getRootElement();
    var output = new Array();
    var results = root.getChild('results', namespace);
    var columnList = results.getChildren('result', namespace);
    var Concept = '';
    var prefLabel='';
    var dac_recip = '';
    var faostat_num = '';
    var gaul_num='';
    var imf_ifs = '';
    var iso_alpha_2 = '';
    var iso_alpha_3 = '';
    var iso_num = '';
    var un_num = '';
    var SDG_indicators_global_database = '';    
    for (var i = 0; i < columnList.length; i++){
        var column = columnList[i];
        var bindings = column.getChildren("binding", namespace);
        for (var j=0; j < bindings.length; j++){
          var bind = bindings[j];
          var attrName = bind.getAttribute('name').getValue();
          var uri = "";
          var text="";
          if(bind.getChild('literal', namespace)){ 
            text = bind.getChild('literal', namespace).getValue();
          }  
          if(bind.getChild('uri', namespace)) { 
               uri = bind.getChild('uri', namespace).getValue();
          }
          if (attrName === 'Concept') {
            Concept = uri;
          } else if (attrName === 'prefLabel'){
            prefLabel = text;        
          } else if (attrName === 'dac_recip'){
            dac_recip= text;
          } else if (attrName === 'dac_donor'){
            dac_recip= text;
          } else if (attrName === 'faostat_num'){
            faostat_num= text;
          } else if (attrName === 'gaul_num'){
            gaul_num= text;
          } else if (attrName === 'imf_ifs'){
            imf_ifs= text;
          } else if (attrName === 'iso_alpha_2'){
            iso_alpha_2= text;
          } else if (attrName === 'iso_alpha_3'){
            iso_alpha_3= text;
          } else if (attrName === 'iso_num'){
            iso_num= text;
          } else if (attrName === 'un_num'){
            un_num= text;
          } else if (attrName === 'SDG_indicators_global_database'){
            SDG_indicators_global_database= uri;
          } 
        } 
        output.push([Concept, prefLabel, dac_recip, faostat_num, gaul_num, imf_ifs, iso_alpha_2, iso_alpha_3, iso_num, un_num, SDG_indicators_global_database]);
      }   
    }

     for (var l=0; l<output.length; l++){
         var object2= output[l];
         var url = String(object2).split(',')[0];
         var country = String(object2).split(',')[1];
         var dac_recip = String(object2).split(',')[2];
         var faostat_num = String(object2).split(',')[3];
         var gaul_num = String(object2).split(',')[4];
         var imf_ifs = String(object2).split(',')[5];
         var iso_alpha_2 = String(object2).split(',')[6];
         var iso_alpha_3 = String(object2).split(',')[7];
         var iso_num = String(object2).split(',')[8];
         var un_num = String(object2).split(',')[9];
         if (to === "dac"){
           output2.push([country,dac_recip]) 
         }
         if(to === "iso2"){
           output2.push([country,iso_alpha_2]) 
           Logger.log([country,iso_alpha_2])
         }
         if(to === "iso3"){
           output2.push([country,iso_alpha_3]) 
         }     
         if(to === "iso"){
           output2.push([country,iso_num]) 
         }  
         if(to === "un"){
           output2.push([country,un_num]) 
         }      
         if(to === "SDG"){
           output2.push([country,SDG_indicators_global_database]) 
         }               
     } 

     for (var l=0; l<output2.length; l++){
         var object2= output2[l];
         var origin2 = String(object2).split(',')[0];
         Logger.log(origin2)           
         var translation2= String(object2).split(',')[1]
         Logger.log([translation2])
         output3.push([translation2])
     }
     if (output3.length==0){  
        var error2 = "We cannot find what you are looking for.";
       return error2; }
      
     return output3;
}



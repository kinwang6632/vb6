function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v3.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; 
  document.MM_sr=new Array; 
  for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){
   	document.MM_sr[j++]=x; 
   	if(!x.oSrc) 
   	  x.oSrc=x.src; 
   	x.src=a[i+2];
   	
   }
}


function checkclick(num) {// sets tabs to active
  if ((toofast == 0) && (num != tabclicked)) {
    toofast = 1; lasttabclicked=tabclicked; 
    tabclicked=num; 
    objslide();
  }  
}

function objshow(showobj) {
  if (n) {document.layers[showobj].visibility = "show"}
  if (ie) {document.all[showobj].style.visibility = "visible"}
}
        
function objhide(hideobj) {
  if (n) {document.layers[hideobj].visibility = "hide"}
  if (ie) {document.all[hideobj].style.visibility = "hidden"}
}

function objslide() {
  if (n) {
    if (opentab == 1) {
      if (lasttabclicked == 0) {
        objhide("International")
      }
      mfactor = tablinks[lasttabclicked] * lfactor + 4  
      if (lasttabclicked == 0) {
        movetab = 0
      }
      else {
        movetab = lasttabclicked + 1
      } 
      moveobj = tabs[movetab]  
      document.layers[moveobj].ypos = parseInt(document.layers[moveobj].top)  
      if (document.layers[moveobj].ypos > (tabtops[movetab] - mfactor)) { 
        for(movetab; movetab < arraylen; movetab++) {  
          moveobj = tabs[movetab] 
          document.layers[moveobj].ypos = parseInt(document.layers[moveobj].top) 
          document.layers[moveobj].ypos -= 5  
          document.layers[moveobj].top = document.layers[moveobj].ypos
        } 
        setTimeout("objslide()",1)
      }
      else {
        for(movetab; movetab < arraylen; movetab++) {  
    	  moveobj = tabs[movetab] 
      	  tabtops[movetab] = parseInt(document.layers[moveobj].top) 
        }  
        opentab = 0  
        if (lasttabclicked != 0){
          objshow(tabs[lasttabclicked]);objhide(tabs2[lasttabclicked])
        }  
        objslide()
      }
    }
    else {
      if (tabclicked != 0 && tabclicked != -1) {
        objshow(tabs2[tabclicked]);objhide(tabs[tabclicked]) 
  	objslide2()
      }
      else {
        if (tabclicked == 0){
          objslide2()
        }
        else {
         toofast = 0
        }
      }
    }
  }
  if (ie) {
    if (opentab == 1) {
      if (lasttabclicked == 0) {
        objhide("International")
      }
      mfactor = tablinks[lasttabclicked] * lfactor + 4 
      if (lasttabclicked == 0) {
        movetab = 0
      }
      else {
	movetab = lasttabclicked + 1
      } 
      moveobj = tabs[movetab] 
      document.all[moveobj].style.ypos = parseInt(document.all[moveobj].style.top)  
      if (document.all[moveobj].style.ypos > (tabtops[movetab] - mfactor)) { 
        for(movetab; movetab < arraylen; movetab++) { 
          moveobj = tabs[movetab] 
       	  document.all[moveobj].style.ypos = parseInt(document.all[moveobj].style.top)
          document.all[moveobj].style.ypos -= 5 
          document.all[moveobj].style.top = document.all[moveobj].style.ypos
        } 
        setTimeout("objslide()",1)
      }
      else {
        for(movetab; movetab < arraylen; movetab++) {  
  	  moveobj = tabs[movetab] 
  	  tabtops[movetab] = parseInt(document.all[moveobj].style.top)
  	} 
  	opentab = 0  
  	if (lasttabclicked != 0) { 
          objshow(tabs[lasttabclicked]);objhide(tabs2[lasttabclicked])
  	}  
  	objslide()
      }
    }
    else {
      if (tabclicked != 0 && tabclicked != -1) {
  	objshow(tabs2[tabclicked]);
        objhide(tabs[tabclicked])  // switches clicked tab to open
  	objslide2()
      }
      else {
  	if (tabclicked == 0) {
          objslide2()
        }
  	else {
          toofast = 0
        }
      }
    }
  } 
}
       
function objslide2() {
  if (n) {
    mfactor = tablinks[tabclicked] * lfactor + 4 
    if (tabclicked == 0) {
      movetab = 0
    }
    else {
      movetab = tabclicked + 1
    } 
    moveobj = tabs[movetab]  
    document.layers[moveobj].ypos = parseInt(document.layers[moveobj].top) 
    if (document.layers[moveobj].ypos < (tabtops[movetab] + mfactor)) {  
      for(movetab; movetab < arraylen; movetab++) { 
        moveobj = tabs[movetab] 
        document.layers[moveobj].ypos = parseInt(document.layers[moveobj].top) 
        document.layers[moveobj].ypos += 5  
        document.layers[moveobj].top = document.layers[moveobj].ypos
      } 
      setTimeout("objslide2()",1)
    }
    else {
      for(movetab; movetab < arraylen; movetab++) {  
  	moveobj = tabs[movetab]  
  	tabtops[movetab] = parseInt(document.layers[moveobj].top) 
      } 
      if (tabclicked == 0) {
        objshow("International")
      }
      opentab = 1; toofast = 0;
    } 	
  }
  if (ie) {
    mfactor = tablinks[tabclicked] * lfactor + 4 
    if (tabclicked == 0) {
      movetab = 0
    }
    else {
      movetab = tabclicked + 1
    } 
    moveobj = tabs[movetab] 
    document.all[moveobj].style.ypos = parseInt(document.all[moveobj].style.top) 
    if (document.all[moveobj].style.ypos < (tabtops[movetab] + mfactor)) { 
      for(movetab; movetab < arraylen; movetab++) {  
      	moveobj = tabs[movetab] 
      	document.all[moveobj].style.ypos = parseInt(document.all[moveobj].style.top) 
      	document.all[moveobj].style.ypos += 5 
        document.all[moveobj].style.top = document.all[moveobj].style.ypos
      } 
      setTimeout("objslide2()",1)
    }
    else {
      for(movetab; movetab < arraylen; movetab++) { 
  	moveobj = tabs[movetab]  
  	tabtops[movetab] = parseInt(document.all[moveobj].style.top) 
      } 
      if (tabclicked == 0) {
        objshow("International")
      }
      opentab = 1; 
      toofast = 0;
     }  		
   }
}
        
function GoToCountry(countryid) {
  win=window.open("","newwin");
  if (countryid==1) {win.location="a"}
  if (countryid==2) {win.location="b"}
  if (countryid==3) {win.location="c"}
  if (countryid==4) {win.location="d"}
  if (countryid==5) {win.location="e"}
  if (countryid==6) {win.location="f"}
  if (countryid==7) {win.location="g"}
  if (countryid==8) {win.location="h"}
}

function GoToColor(locstring) {
  win=window.open("", "win1","scrollbars=0,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=0,width=633,height=444")
  win.location = locstring
}

function chkusr() {
  document.logina.submit()
}


function newWindow(locstring){
  win=window.open("", "win1","scrollbars=0,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=0,width=633,height=444")
  win.location = locstring
}

//**********************************************************

function MM_swapImage1() { //v1.0
  var i,theObj,j=0,swapArray=new Array,oldArray=document.MM_swapImgData;
  for (i=0; i < (MM_swapImage.arguments.length-2); i+=3) {
    theObj = eval(MM_swapImage.arguments[(navigator.appName == 'Netscape')?i:i+1])
    if (theObj != null) {
      swapArray[j++] = theObj;
      swapArray[j++] = (oldArray==null || oldArray[j-1]!=theObj)?theObj.src:oldArray[j];
      theObj.src = MM_swapImage.arguments[i+2];
  } }
  document.MM_swapImgData = swapArray; //used for restore
}


function MM_showHideLayers() { //v1.0
  var i, visStr, args;
  args = MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) { //with arg triples (objNS,objIE,visStr)
    visStr   = args[i+2];
    if (navigator.appName == 'Netscape') {
      if (document.layers != null) eval(args[i]+".visibility = '"+visStr+"'");
    } else { //IE
      if (visStr == 'show') visStr = 'visible'; //convert vals
      if (visStr == 'hide') visStr = 'hidden';
      if (document.all != null) eval(args[i+1]+".style.visibility = '"+visStr+"'");
  } }
}


function MM_swapImgRestore1() { //v1.0
  if (document.MM_swapImgData != null)
    for (var i=0; i<(document.MM_swapImgData.length-1); i+=2)
      document.MM_swapImgData[i].src = document.MM_swapImgData[i+1];
}


var fTransRunning
function ShowLogo()
{
	if (fTransRunning) return; fTransRunning=1
	// Reveal transition type: 0~23
	Logo.filters.item(0).Transition= Math.floor(Math.random() * 24)
	Logo.filters.item(0).Apply(); 
	Logo.style.visibility="visible"; 
	Logo.filters.item(0).play(1); 

	// start the clock
	//startTimePiece("Clock")
}
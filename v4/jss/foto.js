function getQueryVariable(variable)
{
       var query = window.location.search.substring(1);
       var vars = query.split("&");
       for (var i=0;i<vars.length;i++) {
               var pair = vars[i].split("=");
               if(pair[0] == variable){return pair[1];}
       }
       return(false);
}

var centerEl = document.querySelector('center');
var pictureSrc = getQueryVariable('nome').replace(/\\/g, '/').replace('imgs/index.html/', 'imgs/');
var pictureText = window.decodeURIComponent(getQueryVariable('texto') || '');
centerEl.innerHTML = '<p>' + pictureText + '</p><img src="' + pictureSrc + '">'

// De-fphover
// version 0.1
// 2005-12-15
// Copyright: public domain
// --------------------------------------------------------------------
//
// This is a Greasemonkey user script.
//
// To install, you need Greasemonkey: http://greasemonkey.mozdev.org/
// Then restart Firefox and revisit this script.
// Under Tools, there will be a new menu item to "Install User Script".
// You should configure the Included and Excluded pages in the GreaseMonkey
//      configuration pane.
//
// To uninstall, go to Tools/Manage User Scripts,
// select "defphover", and click Uninstall.
//
// --------------------------------------------------------------------
// DESCRIPTION
// Many sites use a Microsoft Frontpage Java applet to generate
// hyperlinks with fancy (for 1995) rollovers. This script converts
// them into regular HTML buttons. It tries to retain as much styling
// as it can from the applet parameters.
//
// NOTE: You'll still sit around and wait for the applets to load
// before they get converted unless you disable Java.
// --------------------------------------------------------------------
// ==UserScript==
// @name          defphover
// @description   Replaces annoying Frontpage fphover java applets with ordinary buttons.
// @include *
// ==/UserScript==


//Convert an fphover applet object to a link by extracting the
//name and text.
function makeButton(appletObj) {
  var params = appletObj.getElementsByTagName("param");
  var label, url, backgroundColor, hoverColor, font, fontSize, fontColor;
  var color, effect;

  //peel out everything embedded in a PARAM.
  for (var i = 0; i < params.length; ++i) {
    var param = params[i];

    if (param.name=="text")  label = param.value;
    if (param.name=="url")   url = param.value;
    if (param.name=="color") color = param.value;
    if (param.name=="bgcolor") backgroundColor = param.value;
    if (param.name=="fontsize") fontSize = param.value+"pt";
    if (param.name=="textcolor") fontColor = param.value;
    if (param.name=="font") font = param.value;
    if (param.name=="effect") effect = param.value;
    if (param.name=="hovercolor") hoverColor = param.value;
  }


  //create a span to hold everything, and style it.
  var button = document.createElement("button");
  if (!!appletObj.getAttribute('width') &&
    !!appletObj.getAttribute('height')) {
    button.style.height = appletObj.getAttribute('height') + "px";
    button.style.width = appletObj.getAttribute('width') + "px";
  }
  button.style.padding = "0;";
  button.style.border = "0";
  button.style.backgroundColor = backgroundColor;
  switch (effect.toLowerCase()) {
    case 'glow':
    case 'average':
      button.style.transition = 'all 150ms ease-out';
      button.style.backgroundImage = 'radial-gradient(ellipse closest-corner, ' + backgroundColor + ' 0%, ' + backgroundColor + ' 100%)';
      button.addEventListener('mouseenter', function(e) {
        button.style.backgroundImage = 'radial-gradient(ellipse closest-corner, ' + backgroundColor + ' 0%, ' + color + ' 100%)';
      });
      button.addEventListener('mouseleave', function(e) {
        button.style.backgroundImage = '';
      })
      break;
    default:
  }

  //only way to do this.
  if (typeof url !== 'undefined') {
    button.addEventListener("click", function() { window.location=url; }, true );
    button.style.cursor = 'pointer';
  }

  //insert styled text into the button.
  var span = document.createElement("span");
  span.style.fontSize = fontSize;
  span.style.color = fontColor;
  span.style.fontFamily = font;
  span.innerHTML = label;

  button.appendChild(span);

  return button;
}


//Program starts here.
var allFPHovers, currentFPHover, button;

allFPHovers = document.evaluate(
    "//applet[@code='fphover.class']",
    document,
    null,
    XPathResult.UNORDERED_NODE_SNAPSHOT_TYPE,
    null);

for (var i = 0; i < allFPHovers.snapshotLength; ++i) {
    currentFPHover = allFPHovers.snapshotItem(i);

    // Replace the applet with a normal link.
    button = makeButton(currentFPHover);
    currentFPHover.parentNode.replaceChild(button, currentFPHover);
}

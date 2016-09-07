// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
// See full license at the bottom of this file.

Office.initialize = function () {
};

function showMessage(message, icon, event) {
  Office.context.mailbox.item.notificationMessages.replaceAsync('msg', {
    type: 'informationalMessage',
    icon: icon,
    message: message,
    persistent: false
  }, function(result){
    event.completed();
  });
}






// This is the function executed by uilessButton1
function buttonFunction1(event) {
  showMessage('uilessButton1 clicked!', 'uilessButtonIcon1-16', event);
}

// This is the function executed by uilessButton2
function buttonFunction2(event) {
  showMessage('uilessButton2 clicked!', 'uilessButtonIcon2-16', event);
}

// This is the function executed by uilessButton3
function buttonFunction3(event) {
  showMessage('uilessButton3 clicked!', 'uilessButtonIcon3-16', event);
}

// This is the function executed by uilessButton4
function buttonFunction4(event) {
  showMessage('uilessButton4 clicked!', 'uilessButtonIcon4-16', event);
}

// This is the function executed by uilessButton5
function buttonFunction5(event) {
  showMessage('uilessButton5 clicked!', 'uilessButtonIcon5-16', event);
}

// This is the function executed by uilessButton6
function buttonFunction6(event) {
  showMessage('uilessButton6 clicked!', 'uilessButtonIcon6-16', event);
}

// This is the function executed by uilessButton7
function buttonFunction7(event) {
  showMessage('uilessButton7 clicked!', 'uilessButtonIcon7-16', event);
}

// This is the function executed by uilessButton8
function buttonFunction8(event) {
  showMessage('uilessButton8 clicked!', 'uilessButtonIcon8-16', event);
}

// This is the function executed by uilessButton9
function buttonFunction9(event) {
  showMessage('uilessButton9 clicked!', 'uilessButtonIcon9-16', event);
}

// This is the function executed by uilessButton10
function buttonFunction10(event) {
  showMessage('uilessButton10 clicked!', 'uilessButtonIcon10-16', event);
}

// This is the function executed by uilessButton11
function buttonFunction11(event) {
  showMessage('uilessButton11 clicked!', 'uilessButtonIcon11-16', event);
}

// This is the function executed by uilessButton12
function buttonFunction12(event) {
  showMessage('uilessButton12 clicked!', 'uilessButtonIcon12-16', event);
}

// This is the function executed by uilessButton13
function buttonFunction13(event) {
  showMessage('uilessButton13 clicked!', 'uilessButtonIcon13-16', event);
}

// This is the function executed by uilessButton14
function buttonFunction14(event) {
  showMessage('uilessButton14 clicked!', 'uilessButtonIcon14-16', event);
}

// This is the function executed by uilessButton15
function buttonFunction15(event) {
  showMessage('uilessButton15 clicked!', 'uilessButtonIcon15-16', event);
}

// This is the function executed by uilessButton16
function buttonFunction16(event) {
  showMessage('uilessButton16 clicked!', 'uilessButtonIcon16-16', event);
}


/*
  MIT License:

  Permission is hereby granted, free of charge, to any person obtaining
  a copy of this software and associated documentation files (the
  "Software"), to deal in the Software without restriction, including
  without limitation the rights to use, copy, modify, merge, publish,
  distribute, sublicense, and/or sell copies of the Software, and to
  permit persons to whom the Software is furnished to do so, subject to
  the following conditions:

  The above copyright notice and this permission notice shall be
  included in all copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
  EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
  NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
  LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
  OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
  WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/
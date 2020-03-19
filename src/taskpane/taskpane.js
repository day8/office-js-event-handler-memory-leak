/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var day8 = {};
day8.curves = {};
day8.curves.main = {};

day8.curves.main.on_click = (function day8$curves$main$on_click(e){
  return Excel.run((function (context){
    console.info("Empty batch function");

    return Promise.resolve(true);
  }));
});
day8.curves.main.mount_app = (function day8$curves$main$mount_app(){
  var button = document.createElement("button");
  (button.innerHTML = "Run");

  (button.onclick = day8.curves.main.on_click);

  Office.onReady((function (info_js){
    return console.info(info_js);
  }));

  return document.body.appendChild(button);
});
document.addEventListener("DOMContentLoaded",day8.curves.main.mount_app);


// Office.onReady(info => {
//   console.info("v5");
//   if (info.host === Office.HostType.Excel) {
//     document.getElementById("sideload-msg").style.display = "none";
//     document.getElementById("app-body").style.display = "flex";
//     document.getElementById("run").onclick = run;
//   }
// });
//
// export var run = (function day8$run(e){
//   return Excel.run((function (context){
//     console.info("Empty batch function");
//
//     return Promise.resolve(true);
//   }));
// });

// export function run() {
//   return Excel.run((function (context ) {
//     console.info("empty batch function");
//     return Promise.resolve(true);
//   }))
// }

// export async function run() {
//   try {
//     await Excel.run(async context => {
//       console.info("empty batch function");
//       return Promise.resolve(true);
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }

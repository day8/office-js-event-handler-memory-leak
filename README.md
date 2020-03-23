# Office-Addin-ContentApp-JS

This repository contains the simplest Content Add-in that is possible to
demonstrate that Content Add-ins effectively leak an alarming amount of memory
for every single call to `Excel.run()`. 

This is reported as [office-js issue #1054](https://github.com/OfficeDev/office-js/issues/1054).

## Setup

1. Open a command prompt as administrator.
2. `CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"` 
3. Open a command prompt as your user.
4. `git clone https://github.com/day8/office-js-memory-leak`
5. `cd office-js-memory-leak`
6. `npm install`
7. `npm start`
8. Excel should open automatically.
9. Click 'Insert'
10. Click 'My Add-ins down arrow' (the down arrow menu on the right, NOT the button on the left)
11. Under 'Developer Add-ins' click 'Content Add-in Memory Leak'
12. If you get the error "We can't open this add-in from localhost" quit Excel, run `npm stop` and go to 7.
13. If the Content Add-in loads successfully resize it so that you can see the 'Run' button.
14. Open Windows Task Manager and observe the memory use of Excel
15. Click the 'Run' button repeatedly.


## Debugging

This template supports debugging using any of the following techniques:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Additional resources

* [Office add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* More Office Add-in samples at [OfficeDev on Github](https://github.com/officedev)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Copyright

Copyright (c) 2019 Microsoft Corporation. All rights reserved.

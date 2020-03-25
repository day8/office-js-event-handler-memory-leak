
This repo provides a minimal demonstration of [office-js issue #1054](https://github.com/OfficeDev/office-js/issues/1054). It has been created using the Microsoft Yeoman office-js generator.


## To Show The Problem

1. Open a command prompt as administrator.
2. `CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"` 
3. Open a command prompt as your user.
4. `git clone https://github.com/day8/office-js-memory-leak`
5. `cd office-js-memory-leak`
6. `npm install`
7. `npm start`
8. Excel should open automatically.
9. Select the 'Insert' tab in the Excel ribbon
10. Click 'My Add-ins down arrow' (the down arrow menu on the right, NOT the button on the left)
11. Under 'Developer Add-ins' click 'Content Add-in Memory Leak'
12. If you get the error "We can't open this add-in from localhost" quit Excel, run `npm stop` and go to 7.
13. If the Content Add-in loads successfully resize it so that you can see the 'Run' button.
14. Open Windows Task Manager and observe the memory use of Excel
15. Click the 'Run' button repeatedly.
16. Observe the memory usage for Excel continuing to climb each time Run is clicked. 
![office-js 2020-03-20 13-44](https://user-images.githubusercontent.com/350450/77129910-a309a100-6aba-11ea-91b5-99abb5d4276f.gif)
17. After clicking the button LOTS of times, notice how memory is just contnuing to rise unbounded
18. Leave the app alone for 20 mins and notice that at some point the used memory returns to normal

## Now look at the onclick handler

When you click on that `Run` button, the [`onclick`](https://github.com/day8/office-js-memory-leak/blob/master/src/contentapp/contentapp.js#L16-L24) handler is [this minimal code](https://github.com/day8/office-js-memory-leak/blob/master/src/contentapp/contentapp.js#L16-L24).  It doesn't get any simpler. 



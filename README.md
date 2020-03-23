
This repo provides a minimal demonstration of [office-js issue #1054](https://github.com/OfficeDev/office-js/issues/1054). 


## To Show The Problem

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
16. Observe the memory usage for Excel continuing to climb each time Run is clicked. 
![office-js 2020-03-20 13-44](https://user-images.githubusercontent.com/350450/77129910-a309a100-6aba-11ea-91b5-99abb5d4276f.gif)

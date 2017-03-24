# Pocket Test Automation Framework(口袋自动化测试框架)
## 1.What's Pocket TA Framework?
Pocket TA framework like Sikuli, but easy to use than it, the framework is develped by AutoHotkey and also supports to automate anything you see on the screen. It uses image or color recognition to identify and control GUI components.
pocket means easy to develop, easy to maintain and easy to share your scripts to others. No need to learn much for a tester or developer to master it. Page UI test scripts support IE, Chrome and firefox run on windows, framework provides one key to pause, resume and stop script excution. Using third-party libraries of AutoHotKey can make a lot of incredible things. 

## 2. How to use?
step1: Go to [AutoHotKey site](https://autohotkey.com/) and download AutoHotkey Installer, install AutoHotkey v1.1.* 
step2: git clone https://github.com/wofeng/PocketTAFramework.git
step3: unzip and install IDE - SciTE4AHK300601_Install.zip, Log console - DebugView.zip and screen ruler - ScreenRuler.zip 
note: all can be found in softwares directory

## 3. Write "hello world"
1. create a directory "helloworld" under "Modules" directory
2. create a helloworld.apk 
<pre><code>
#include ../../Common/scriptheader.ahk
;main function
Gosub, Helloworld
ExitApp
;case steps
Helloworld:
MsgBox, Hello world! 
return

#include ../../Common/scriptfooter.ahk
</code></pre>

## 4. Specified steps execution
create a <filename>.ini under "common" directory, open DebugView to see test case execution record
anyname.ini:
<code>
[Cases]
c1=OpenSite
c2=SignIn
c3=SignOut
</code>

login.ahk:
<pre><code>
#include ../../Common/scriptheader.ahk
Gosub, RunAllCases ; do not change it, the label was defined in globalVar.ahk
ExitApp
IgoreStep:
MsgBox, ignore the step
return

OpenSite:
OutputDebug, Open a site
return

SignIn:
OutputDebug, sign in 
return

SignOut:
OutputDebug, sign out
return
#include ../../Common/scriptfooter.ahk
</code></pre>

## 5. Specified test cases execution
1. create TestSuite.ahk in each test suite (sub-directory under Modules)to include test files
2. create <anyname>.ini to specified steps 
3. Register test suite in "[Modules]" section in "sysconfig.ini"
4. Run __BatchRun__/startBatchRun.ahk
5. see test case execution record in DebugView
refer to batchRunExample1/batchRunExample2/batchRunExample3 test suites

## FAQ
1. Screen resolution difference will affect the image identification?
After test, My answer is NO, but small resolution will reduce visual area of browser, UI elements outside of visual area can not
be identified. screenshot image from zoom in or out browser can not be identified by script run on normal browser

2. Why image can be identified by Chrome but not identified by IE?
 Occasionally, the issue can be occurred, but I don't know the root cause, but you can try to get the image from IE,
 Chrome can identified it

3. How to stop script execution?
framework provides one key to pause, resume and stop script excution.
press "Pause" key to pause or resume script execution
press "ESC" key to stop script execution
press "PrintScreen" key to generate screenshot file

4. How to distribute my test scripts to others?
Using Convert .ahk to .exe tool from AutoHotkey. compress the directory to zip file , notice
keep the directory structure, because it depends on "ini" files include sysconfig.ini.
others only double clicks the executable file to run it without any complicated configuration.
you can use TestCaseManager tool to manage your test cases.

Enjoy it :)

Reference: http://wofeng.github.io/blog/2016/06/09/2016-06-09-PocketTA/  (Chinese)

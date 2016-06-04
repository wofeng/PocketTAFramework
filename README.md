# Pocket TA FrameWork(口袋自动化测试框架)
1.Pocket TA Framework简述
口袋自动化测试框架意为开发环境简单，易于移植，脚本开发简单，易于维护和分享。它是基于AutoHotKey开发的针对Windows自动化测试框架，测试或开发人员无需掌握太多专业知识就可以快速编写自动化测试脚本，辅助QA日常手工测试。TA脚本可以编译成可执行文件方便分享。一套脚本支持Windows系统下Chrome, Firefox和IE浏览器下运行， 支持页面及时调试，一键暂停，继续，终止和截屏等功能，利用它可以完成基于UI测试，HTTP API 测试等。

2. AutoHotkey简述
本框架是基于Autohotkey_L 脚本语言开发，首先先来了解下AutoHotKey。AutoHotKey (Basic) 由Chris Mallett开发维护，是一款免费的、Windows平台下开放源代码的热键脚本语言。从名字可以想象出它是通过热键执行所编写的宏，使用者可以为键盘，游戏杆和鼠标创建热键，通过发送键盘或鼠标的键击动作命令来实现几乎所有操作的自动化。一些高级玩家会利用它开发游戏自动化脚本。由于AutoHotKey Basic并不支持Unicode, 并且2009年推出的1.0.48.05版本后就停止更新，AutoHotkey_L是Lexikos维护AutoHotkey定制编译版本支持Unicode，其在AutoHotkey Basic基础上增加很多特性，具体不再熬述，目前大部分开发者都采用AutoHotKey_L版本。

3.Pocket TA Framework介绍
3.1背景AutoHotKey脚本语言提供了大量的API，除此之外具有强大的扩展能力，支持Windows COM, Win32 API等，但是它并不是个自动化测试框架而更像是一种开发语言，仅仅利用原生的API编写UI自动化测试脚本并不容易，需要针对UI自动化测试的特点，基于AutoHotKey语法进行二次开发。要实现UI自动化测试框架，我认为需要解决以下几个问题。
1.  规范脚本结构，最好有个统一的模板；
2.  能够很好地组织TA Case 和Test Suite，最好具备选择执行测试脚本的功能；
3.  方便执行单条Case和批量Cases；
4.  通用简单易用的方法实现元素的定位，操作，和结果判断；
5.  记录脚本详细的运行信息，方便使用者了解Case的执行结果

3.2 Pocket TA Framework 系统组成Pocket TA Framework 由三部分组成，即Common模块，Lib模块和Modules模块组成
3.2.1 Common 模块它由全局变量配置文件(sysconfig.ini)，TA Cases 运行的初始化程序GlobalVars.ahk和TA脚本模板引用文件(scriptheader.ahk, scriptfooter.ahk)组成。3.2.1.1全局配置文件全局配置文件(sysconfig.ini)含有两部分：[SysConfig]定义全局配置信息比如CaseLevel-指定当前TA case执行的级别；brexe -指定当前运行浏览器类型；region_global-机器屏幕大小；site_url-测试站点等等信息，开发的Library API依赖这部分全局变量。[Modules] 注册所需要运行的模块名，它位于Modules目录下，这部分信息仅用于批量脚本文件执行。

3.2.1.2系统初始化系统初始化程序GlobalVars主要负责初始化全局变量，同时指定快捷键支持脚本运行期的暂停(Pause)，继续(Pauseagain)，终止(ESC)和一键(PrintScreen)生成截屏文件等功能。

3.2.1.3脚本模板引用文件脚本模板引用文件(scriptheader, scriptfooter)把系统运行必要的文件和初始化过程封装起来，方便测试人员编写TA 脚本。下面是脚本文件模板
#include  ../../common/scriptheader.ahk;
;----------------------------------Main  Process----------------------------------
Gosub, DataPreparationGosub, RunCasesExitApp;
;----------------------------------Load User Defined Veriables--------------------
DataPreparation:
appini := GetIniFile("demo.ini")                    ;模块用户自定义变量
appobj := getKVObjFromIni(appini)return;
;----------------------------------Detail Cases-----------------------------------
;Case name需要注册到所在模块的common目录下配置文件中当前定义Case Level下面
Case1:
; TODO
return

Case2:
;  TODO
return

CaseN:
;  TODO
return
#include  ../../common/scriptfooter.ahk

3.2.2Lib 模块Lib模块主要含有两部分 1. AHK标准库
2.自己扩展的TA的关键字和一些常用方法。引入的部分AHK标准库是为了更方便地开发适合项目需求的功能。Scriptlib.ahk是自己开发的API集，
目前V1.2已开发40多个方法。
工具方法
GetIniFile, GetAllVarsFromIniFile, GetKVObjFromIni, GetSplittedObj, UriEncode, UriDecode,EvalString常用关键字Debug,Highlight, GetHTMLCodeFromClient, GetHTMLCodeFromServ, GetTextFromHtml,GetTextFromPage,CopyURLFromAddressBar,GoToURLByPaste, PrtScreen, PrtFromScreen,  ShowTooltip,WaitForImage, WaitForImageVanish, WaitForPixelColor,WaitForPixelColorVanish, OpenSiteWithBrowser,ActiveBrowserWin, ClickElement,ClickImage, ClickDragMouse,MoveMouseToCoor,     MoveMouseToImage

常用关键字
Debug,Highlight, GetHTMLCodeFromClient, GetHTMLCodeFromServ, GetTextFromHtml,GetTextFromPage,CopyURLFromAddressBar,GoToURLByPaste, PrtScreen, PrtFromScreen,  ShowTooltip,WaitForImage, WaitForImageVanish, WaitForPixelColor,WaitForPixelColorVanish, OpenSiteWithBrowser,ActiveBrowserWin, ClickElement,ClickImage, ClickDragMouse,MoveMouseToCoor,     MoveMouseToImage

扩展用于项目的方法(Email)
SendEmailFromOutLook, DeleteEmail, DeleteAccount, GetEmailBody,GetEmailAttachmentName, VerifyEmailKeyInfo, VerifyEmailAttachmentName

3.2.3
Modules模块Modules模块主要放置用户编写脚本，其中BatchRun支持脚本的批量运行，它根据sysconfig.ini注册的模块和每个注册模块下TestSuite中注册的脚本文件执行。Modes1..n 是具体的模块名，名称必须与sysconfig,ini中注册的模块名保持一致。每个模块含有common目录，存放一个或多个配置文件用于注册有效的Test Cases。Test Cases用到的变量信息都配置在一个或多个用户配置文件中( Module Configuration File(s) )，位置可自由指定，通过模板Load User Defined Veriables脚本自动加载。每个脚本文件可以存放一个或多条Test Case, 它可以由一个或多个Label 脚本片段组成。Test suite文件只用于Batch run模式下，它注册了哪些脚本文件需要被执行。

3.3 脚本文件执行过程通过框架介绍，了解了Pocket TA Framework结构，下面介绍下执行TA Case的过程，分单脚本文件执行和批量脚本文件执行。
3.3.1 单脚本文件执行
开始执行脚本文件，系统会把脚本内容逐行加载到内存中，从上往下按顺序执行其中的命令，直到遇到Return，Exit，热键/热字串标签或脚本的底部，这是由AutoHotKey系统控制，脚本加载完毕首先执行scriptheader脚本中initialGlobalVars脚本片段，完成全局变量初始化，然后执行模板中DataPreparation脚本片段，加载模块用户自定义变量，接下来执行当前脚本文件注册过的所有Test Cases（位于当前模块下common目录下所有配置文件）, 这个过程是根据Case Level收集当前模块下所有注册的Test Cases，然后依次执行，如果注册Test Case （Label）名称不在当前脚本文件中，框架会忽略，如果Test Case有相互依赖关系，则存在依赖关系的Test Case 名称(Label名)需要用下划线(“_”)开始，框架在运行期会自动处理，如果被依赖的脚本失败则所有有依赖关系的Test Cases都不再执行。

3.3.2 批量脚本执行
首先通过sysconfig.ini获取所有注册模块名称，然后遍历每个模块，加载每个模块下TestSute，依次执行Test Suite中注册的脚本文件，每个脚本文件会执行它所含有的注册的Test Cases，与脚本文件执行过程一样。当前模块Test Suite执行完毕，进入下个注册模块，直到所有注册模块都被执行。

4. UI自动化相关技术
4.1元素定位Pocket TA framework主要采用如下4种元素定位方法。图像定位技术: 通过截图对当前可视化屏幕区域或指定区域进行匹配，找到返回截图所在屏幕左上角坐标位置。颜色定位技术: 通过像素颜色(RGB值)搜素指定区域，找到则返回第一个满足该色彩值像素点坐标。快捷键定位技术: 通过快捷键实现指定元素的定位，适合系统功能，浏览器组件定位，Web Form组件定位等。元素相对坐标定位技术: Web 2.0 通常采用DIV技术进行页面布局，为了保持页面在各个浏览器效果一致性，通常页面采用固定宽度(比如常见960 pixels)并且UI元素距离保持不变。Pocket TA Framework推荐使用图像定位和元素定位相结合的方式。我们把整个窗口看成坐标系，横轴自左向右为正，纵轴自上而下为正，反之为负，首先在当前页面找到一个参考图像，通过图像定位获取改图片的左上角坐标，通过屏幕尺子等测量工具测量出所要操作的元素与参考图片的距离，通过框架提供的命令可以完成鼠标的定位。

4.2操作元素关键字ClickElement,ClickImage,ClickDragMouse, MoveMouseToCoor,  MoveMouseToImage, 覆盖常见鼠标操作，移动，单击，双击，拖拽。4.3判断结果可以通过指定图片是否出现完成结果页面验证，可以利用WaitForImage,WaitForImageVanish, WaitForPixelColor和WaitForPixelColorVanish命令。

4.3判断结果可以通过指定图片是否出现完成结果页面验证，可以利用WaitForImage,WaitForImageVanish, WaitForPixelColor和WaitForPixelColorVanish命令。
对于文字验证，一种方式在文字所在区域通过拖拽把文字复制到粘贴板后进行文字比较，可以通过GetTextFromPage 方法; 另一种通过获取结果页面HTML，获取任意HTML标签内文本信息,然后进行比较，可以通过GetTextFromHtml方法。

5. 编写和Debug脚本
Pocket TA Framework 最大的亮点支持脚本即时调试运行，借助AutoHotKey实现一键脚本调试运行，Web脚本无需重复执行登录操作，即可实现所见即所得的效果，大大提高脚本调试，运行效率，而且调试成功的脚本无需任何语法转换就可以使用。借助Debug  View工具可以观察Debug信息，非常方便。以下是Debug 脚本模板，通过按ALT+WIN+r (快捷键可随意定制)执行该脚本，ESC终止脚本运行。

对于文字验证，一种方式在文字所在区域通过拖拽把文字复制到粘贴板后进行文字比较，可以通过GetTextFromPage 方法; 另一种通过获取结果页面HTML，获取任意HTML标签内文本信息,然后进行比较，可以通过GetTextFromHtml方法。5. 编写和Debug脚本Pocket  TA Framework 最大的亮点支持脚本即时调试运行，借助AutoHotKey实现一键脚本调试运行，Web脚本无需重复执行登录操作，即可实现所见即所得的效果，大大提高脚本调试，运行效率，而且调试成功的脚本无需任何语法转换就可以使用。借助Debug  View工具可以观察Debug信息，非常方便。以下是Debug 脚本模板，通过按ALT+WIN+r (快捷键可随意定制)执行该脚本，ESC终止脚本运行。
#include  ../../common/scriptheader.ahk
;----------------------------------Main  Process----------------------------------------
!#r::; 按ALT + WIN + r  运行此脚本
Gosub, DataPreparation
Gosub, TestCase1           ;设置需要Debug Case的label 名称
ExitApp
;----------------------------------Load User Defined Veriables---------------------
DataPreparation:
appini := GetIniFile("demo.ini"); 加载用户自定义变量
appobj := getKVObjFromIni(appini)
return;
----------------------------------Detail Cases-------------------------------------------
TestCase1:
;  TODO
return
#include  ../../common/scriptfooter.ahk

6. 总结麻雀虽小，五脏六腑俱全，Pocket TA framework 相对目前自动化测试工具的优势在于
1. 开发环境安装配置简单，体积小(<10M)；
2. 强大的快捷键支持一键调试，提高脚本调试运行效率；
3. 具有TA Case管理功能；
4. 有统一的脚本开发调试模板，易于掌握；
5. 图像结合相对坐标的定位方法，不依赖于代码，维护简单
6.  针对项目定制的一些方法可以实现UI和逻辑分层测试和混合测试，提高UI自动化测试效率。

Selenium具有系统和语言的优势，但限制在于元素定位方式依赖于HTML结构和标签属性，所以它更适合作为产品成型后自动化测试平台。Pocket TA framework可以调用Windows API，更胜任Windows平台下程序的自动化测试，它的元素定位方式不依赖于HTML结构，当然Pocket TA framework也有自身的局限性1. 支持平台有限仅适合Windows；2. 由于采用图像和相对坐标定位方式，往往受某些因素影响，例如屏幕分辨率大小不同会导致可视化区域大小的变化，当前可视化区域以外元素无法定位；图像匹配技术采用像素点颜色比较(pixel-by-pixel)，插值像素会影响匹配结果，比如图像的缩放

7. 项目实践下面针对WebEx11项目,通过实际脚本展示如何自动化测试UI。
SignIn:
;站点URL, 当前使用的浏览器类型等已经配置在全局变量文件中(sysconfig.ini)，模板文件自动处理
OpenSiteWithBrowser()         ;用指定类型浏览器打开站点首页
xy := WaitForPixelColor("0x4B8600", "100, 200, 500, 500")     ;等待绿色Sign In 按钮出现
Send, ^0           ; 按Ctrl+0  -重置浏览器内容缩放比例到100%
WaitForImage("wxball.png")      ; 等待首页加载完毕
Send % "477875@qq.com". "{Tab}" ; 输入用户名
Sleep, 200Send % "P@ss123"      ; 输入用户密码
ClickElement(xy[1]+30 . "," . xy[2]+5,,"W")     ; 通过颜色定位方法获得Sign In按钮位置，点击登录
WaitForImage("upload.png")      ; 等待Home页面加载完毕
Return

;---完----

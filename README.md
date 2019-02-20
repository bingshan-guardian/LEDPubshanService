# LEDPublishService

用于在Windows系统上向LED屏推送文本  
安装后为Windows服务  

## 安装方法
* 运行--〉cmd：打开cmd命令框
* 在命令行里定位到`InstallUtil.exe`所在的位置
    * InstallUtil.exe 默认的安装位置是在`C:/Windows/Microsoft.NET/Framework/v2.0.50727`里面
    * 通过cd定位到该位置 `cd C:/Windows/Microsoft.NET/Framework/v2.0.50727`
* 操作命令：
    * 安装服务命令：在命令行里输入下面的命令：
        `InstallUtil.exe  PATH_TO/LEDPublishService.exe`
    * 启动服务命令
        `net start LEDPublishService`
    * 停止服务命令
        `net stop LEDPublishService`
    * 卸载服务命令：在命令行里输入下面的命令：
        `InstallUtil.exe /u  PATH_TO/LEDPublishService.exe`

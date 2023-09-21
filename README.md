需要自己把 `.xls` 文件用 `Excel` 打开，然后另存为 `.xlsx` 格式。

钉钉信息写到 `/src/config.jl` 里。

调整参数后，运行 `./src/_main_.jl` 即可。

成果如图：

![image](https://github.com/Xlin0mu/SDUCourseTable2DingTalkCalender/assets/99205570/7e0b2a26-9cf4-4820-b1c2-bae5c4a478e9)

# `config.jl` 配置

钉钉企业内部应用开发的参数，还有改权限，哪里卡住发issue

首先注册企业，流程很简单。

成为开发者：

https://open.dingtalk.com/document/orgapp/become-a-dingtalk-developer

创建应用并获取 `AppKey` 和 `AppSecret`：

https://open.dingtalk.com/document/orgapp/create-orgapp

权限：

https://open.dingtalk.com/document/orgapp/permission-management
